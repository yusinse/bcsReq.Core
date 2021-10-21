using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aras.IOM;
using System.Xml;
using Newtonsoft.Json;
using bwInnovatorCore;

namespace bcsReq.Core
{
    public class Compare
    {
        #region "                   宣告區"

        protected Innovator Cinn;
        protected CInnovator.bwGeneric CbwGeneric = new CInnovator.bwGeneric(); //Modify by kenny 2019/04/11
        protected CGeneric.Common CoCommon;//Modify by kenny 2019/04/11
        private string CstrErrMessage = "";

        protected Innovator innovator { get; private set; }
        protected string LangCode { get; private set; }

        #endregion

        #region "                   進入點"

        public Compare()
        {
            Cinn = new Innovator(null);
        }


        public Compare(Innovator getInnovator)
        {
            //System.Diagnostics.Debugger.Break();

            innovator = getInnovator;
            Cinn = getInnovator;
            CbwGeneric.bwIOMInnovator = Cinn;
            CoCommon = new CGeneric.Common();

            //Modify by kenny 2016/04/01 ------
            string LanCode = Cinn.getI18NSessionContext().GetLanguageCode();
            if (LanCode == null) LanCode = "";
            LangCode = LanCode;
            LanCode = LanCode.ToLower();

            if ((LanCode.IndexOf("zt") > -1) || (LanCode.IndexOf("tw") > -1))
            {
                LanCode = "zh-tw";
            }
            else if ((LanCode.IndexOf("zc") > -1) || (LanCode.IndexOf("cn") > -1))
            {
                LanCode = "zh-cn";
            }
            else if ((LanCode.IndexOf("kr") > -1) || (LanCode.IndexOf("ko") > -1))
            {
                LanCode = "ko-kr";
            }
            else
            {
                LanCode = "en";
            }
            CoCommon.SetLanguage = LanCode;
            CbwGeneric.SetLanguage = LanCode;
            //----------------------------------
        }

        #endregion

        #region "                   屬性區"

        protected virtual string CstrLicenseCode
        {
            get { return "AO-09012"; }
        }

        //Add Property by kenny 2016/03/02
        public string ErrMessage
        {
            get { return CstrErrMessage; }
        }


        #endregion

        #region "                   方法區"

        public Item CompareItems(string itemId1, string itemId2, string itemtypeName)
        {
            if (CheckLicense() == false)
            {
                return Cinn.newError(CstrErrMessage);
            }

            try
            {
                string result="";
                switch (itemtypeName.ToLower())
                {
                    case "re_requirement_document":
                        result= CompareRequirementDoc(itemId1, itemId2);
                        break;
                    case "re_requirement":
                        result = CompareRequirement(itemId1, itemId2);
                        break;
                }

                return innovator.newResult(result);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private string CompareRequirementDoc(string itemId1, string itemId2)
        {
            ResultCompares result = new ResultCompares();

            //读取需求文档数据关系数据
            Item reqDocRels1 = innovator.newItem("re_Req_Doc_Content", "get");
            reqDocRels1.setAttribute("select", "chapter,related_id(id,config_id,content,generation)");
            reqDocRels1.setProperty("source_id", itemId1);
            reqDocRels1 = reqDocRels1.apply();
            Item reqDocRels2 = innovator.newItem("re_Req_Doc_Content", "get");
            reqDocRels2.setAttribute("select", "chapter,related_id(id,config_id,content,generation)");
            reqDocRels2.setProperty("source_id", itemId2);
            reqDocRels2 = reqDocRels2.apply();

            for (int i = 0; i < reqDocRels1.getItemCount(); i++)
            {
                Item req1 = reqDocRels1.getItemByIndex(i).getRelatedItem();
                string req_id = req1.getID();
                string req_configid1 = req1.getProperty("config_id");
                string req_gen = req1.getProperty("generation");

                RequirementCompareInfo compareReq1 = new RequirementCompareInfo();
                compareReq1.reqConfigId = req_configid1;
                compareReq1.reqId = req_id;

                Item req2 = reqDocRels2.getItemsByXPath("//Item[config_id='" + req_configid1 + "']");
                if (req2.getItemCount() < 1)
                {
                    compareReq1.state = "delete";
                }
                else
                {
                    string req_id2 = req2.getID();
                    RequirementCompareInfo compareReq2 = new RequirementCompareInfo();
                    compareReq2.reqConfigId = req_configid1;
                    compareReq2.reqId = req_id2;
                    if (req_gen == req2.getProperty("generation"))
                    {
                        compareReq1.state = "same";
                        compareReq2.state = "same";
                    }
                    else
                    {
                        string content1 = req1.getProperty("content");
                        string content2 = req2.getProperty("content");
                        XmlDocument contentDom1 = new XmlDocument();
                        contentDom1.LoadXml(transformXml(content1));
                        XmlDocument contentDom2 = new XmlDocument();
                        contentDom2.LoadXml(transformXml(content2));

                        bool is_different = false;

                        foreach (XmlNode field1 in contentDom1.DocumentElement.ChildNodes)
                        {
                            string fieldName = field1.Name;
                            if (fieldName == "Requirement-Info")
                            {
                                continue;
                            }
                            string uid = field1.Attributes["id"].Value;
                            ElementCompareInfo compareElement1 = new ElementCompareInfo();
                            compareElement1.elementId = uid;

                            XmlNode field2 = contentDom2.SelectSingleNode("//*[@id='" + uid + "']");
                            if (field2 == null)
                            {
                                compareElement1.state = "delete";
                                compareReq1.elements.Add(compareElement1);
                                is_different = true;
                                continue;
                            }

                            ElementCompareInfo compareElement2 = new ElementCompareInfo();
                            compareElement2.elementId = uid;
                            if (field1.InnerXml != field2.InnerXml)
                            {
                                compareElement1.state = "change";
                                compareElement2.state = "change";
                                is_different = true;
                            }
                            else
                            {
                                compareElement1.state = "same";
                                compareElement2.state = "same";
                            }
                            compareReq1.elements.Add(compareElement1);
                            compareReq2.elements.Add(compareElement2);
                        }

                        foreach (XmlNode field2 in contentDom2.DocumentElement.ChildNodes)
                        {
                            string uid = field2.Attributes["id"].Value;
                            string fieldName = field2.Name;
                            if (fieldName == "Requirement-Info" || compareReq2.elements.Exists(t=>t.elementId==uid))
                            {
                                continue;
                            }
                            if (contentDom1.SelectSingleNode("//*[@id='" + uid + "']") != null)
                            {
                                continue;
                            }
                            ElementCompareInfo compareElement2 = new ElementCompareInfo();
                            compareElement2.elementId = uid;
                            compareElement2.state = "add";
                            compareReq2.elements.Add(compareElement2);
                            is_different = true;
                        }

                        if (is_different)
                        {
                            compareReq1.state = "change";
                            compareReq2.state = "change";
                        }
                        else
                        {
                            compareReq1.state = "same";
                            compareReq2.state = "same";
                        }
                    }
                    result.reqDoc2.Add(compareReq2);
                }
                result.reqDoc1.Add(compareReq1);
            }

            for (int i = 0; i < reqDocRels2.getItemCount(); i++)
            {
                Item req2 = reqDocRels2.getItemByIndex(i).getRelatedItem();
                string req_id2 = req2.getID();
                string req_configid2 = req2.getProperty("config_id");
                if (result.reqDoc2.Exists(t => t.reqConfigId == req_configid2))
                {
                    continue;
                }
                RequirementCompareInfo compareReq2 = new RequirementCompareInfo();
                compareReq2.reqConfigId = req_configid2;
                compareReq2.reqId = req_id2;
                compareReq2.state = "add";
                result.reqDoc2.Add(compareReq2);
            }

            string json = JsonConvert.SerializeObject(result);
            return json;
        }

        private string CompareRequirement(string itemId1, string itemId2)
        {
            ResultCompares result = new ResultCompares();

            //读取需求内容数据
            Item req1 = innovator.newItem("re_requirement", "get");
            req1.setAttribute("select", "id,config_id,content");
            req1.setAttribute("language", "en");
            req1.setID( itemId1);
            req1 = req1.apply();
            Item req2 = innovator.newItem("re_requirement", "get");
            req2.setAttribute("select", "id,config_id,content");
            req2.setAttribute("language", "en");
            req2.setID(itemId2);
            req2 = req2.apply();

            string content1 = req1.getProperty("content","","en");
            string content2 = req2.getProperty("content","","en");
            XmlDocument contentDom1 = new XmlDocument();
            contentDom1.LoadXml(transformXml(content1));
            XmlDocument contentDom2 = new XmlDocument();
            contentDom2.LoadXml(transformXml(content2));

            RequirementCompareInfo compareReq1 = new RequirementCompareInfo();
            compareReq1.reqConfigId = req1.getProperty("config_id");
            compareReq1.reqId = itemId1;
            RequirementCompareInfo compareReq2 = new RequirementCompareInfo();
            compareReq2.reqConfigId = req2.getProperty("config_id");
            compareReq2.reqId = itemId2;
            foreach (XmlNode field1 in contentDom1.DocumentElement.ChildNodes)
            {
                string fieldName = field1.Name;
                if (fieldName == "Requirement-Info")
                {
                    continue;
                }
                string uid = field1.Attributes["id"].Value;
                ElementCompareInfo compareElement1 = new ElementCompareInfo();
                compareElement1.elementId = uid;

                XmlNode field2 = contentDom2.SelectSingleNode("//*[@id='" + uid + "']");
                if (field2 == null)
                {
                    compareElement1.state = "delete";
                    compareReq1.elements.Add(compareElement1);
                    continue;
                }

                ElementCompareInfo compareElement2 = new ElementCompareInfo();
                compareElement2.elementId = uid;
                if (field1.InnerXml != field2.InnerXml)
                {
                    compareElement1.state = "change";
                    compareElement2.state = "change";
                }
                else
                {
                    compareElement1.state = "same";
                    compareElement2.state = "same";
                }
                compareReq1.elements.Add(compareElement1);
                compareReq2.elements.Add(compareElement2);
            }

            foreach (XmlNode field2 in contentDom2.DocumentElement.ChildNodes)
            {
                string uid = field2.Attributes["id"].Value;
                string fieldName = field2.Name;
                if (fieldName == "Requirement-Info" || compareReq2.elements.Exists(t => t.elementId == uid))
                {
                    continue;
                }
                if (contentDom1.SelectSingleNode("//*[@id='" + uid + "']") != null)
                {
                    continue;
                }
                ElementCompareInfo compareElement2 = new ElementCompareInfo();
                compareElement2.elementId = uid;
                compareElement2.state = "add";
                compareReq2.elements.Add(compareElement2);
            }

            result.reqDoc1.Add(compareReq1);
            result.reqDoc2.Add(compareReq2);
            string json = JsonConvert.SerializeObject(result);
            return json;
        }

        private string transformXml(string xml)
        {
            string newXml = System.Text.RegularExpressions.Regex.Replace(
                        xml,
                        @"(xmlns:?[^=]*=[""][^""]*[""])", "",
                        System.Text.RegularExpressions.RegexOptions.IgnoreCase |
                        System.Text.RegularExpressions.RegexOptions.Multiline);
            newXml = newXml.Replace("aras:", "");
            return newXml;
        }

        private class ResultCompares
        {
            public List<RequirementCompareInfo> reqDoc1 = new List<RequirementCompareInfo>();
            public List<RequirementCompareInfo> reqDoc2 = new List<RequirementCompareInfo>();
        }

        private class RequirementCompareInfo
        {
            public string reqConfigId;
            public string reqId;
            public string state;
            public List<ElementCompareInfo> elements = new List<ElementCompareInfo>();
        }

        private class ElementCompareInfo
        {
            public string elementId;
            public string state;
        }

        #endregion

        #region "                   方法區(內部使用)"

        private bool CheckLicense()
        {
            try
            {
                if (!CbwGeneric.IsLicenseUseRuntimeFunctionByName(CstrLicenseCode))
                {
                    CstrErrMessage = CbwGeneric.ErrorException;

                    if (CstrErrMessage == "") CstrErrMessage = CoCommon.GetMessageByMsgId("msg_gen_000024", "授權碼驗證不正確");

                    return false;
                }

                return true;
            }
            catch (Exception ex)
            {
                CstrErrMessage = ex.Message;
                return false;
            }
        }

        #endregion

    }
}
