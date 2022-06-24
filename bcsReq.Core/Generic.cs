using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using Aras.IOM;
using Aras.Server.Core;
using bwInnovatorCore;
using HtmlAgilityPack;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace bcsReq.Core
{
    public class Generic
    {
        #region "                   宣告區"

        protected Innovator Cinn;
        protected CInnovator.bwGeneric CbwGeneric = new CInnovator.bwGeneric(); //Modify by kenny 2019/04/11
        protected CGeneric.Common CoCommon;//Modify by kenny 2019/04/11
        private string CstrErrMessage = "";

        protected Innovator innovator { get; private set; }
        protected string LangCode { get; private set; }

        protected CallContext CCO;
        protected StringBuilder gridStyle;

        public static List<string> iconList = new List<string>();

        #endregion

        #region "                   進入點"

        public Generic()
        {
            Cinn = new Innovator(null);
        }


        public Generic(Innovator getInnovator, CallContext cco)
        {
            //System.Diagnostics.Debugger.Break();

            innovator = getInnovator;
            CCO = cco;
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

        public Item GetReqDocGrid(Item reqDoc)
        {
            if (CheckLicense() == false)
            {
                return Cinn.newError(CstrErrMessage);
            }

            string reqdoc_id = reqDoc.getID();
            if (String.IsNullOrEmpty(reqdoc_id))
            {
                return Cinn.newError("未读取到需求文档ID");
            }

            gridStyle = new StringBuilder();

            GetReqDocBaseTamplate();

            Item reqdoc_relationReqs = Cinn.newItem("re_Req_Doc_Content", "get");
            reqdoc_relationReqs.setAttribute("select", "id,chapter,related_id(req_title)");
            reqdoc_relationReqs.setAttribute("orderBy", "chapter");
            reqdoc_relationReqs.setProperty("source_id", reqdoc_id);
            reqdoc_relationReqs = reqdoc_relationReqs.apply();
            if (reqdoc_relationReqs.getItemCount() > 0)
            {
                for (int i = 0; i < reqdoc_relationReqs.getItemCount(); i++)
                {
                    string chapter = (i + 1).ToString();
                    Item reqdoc_relationReq = reqdoc_relationReqs.getItemsByXPath("//Item[chapter='" + chapter + "']");
                    if (reqdoc_relationReq.getItemCount() < 1)
                    {
                        break;
                    }
                    if (reqdoc_relationReq.isCollection())
                    {
                        return Cinn.newError("数据结构错误!");
                    }
                    GetReqDocStructure(reqdoc_relationReqs, reqdoc_relationReq, chapter, 0);
                }
            }

            gridStyle.Append("</table>");
            return Cinn.newResult(gridStyle.ToString());
        }

        private void GetReqDocStructure(Item reqdoc_relationReqs, Item parentReqdocRel, string chapter, int level)
        {
            Item parentReq = parentReqdocRel.getRelatedItem();

            gridStyle.Append("<tr level=\"");
            gridStyle.Append(level.ToString());
            gridStyle.Append("\" icon0=\"../Solutions/RE/Images/Requirement.svg\" icon1=\"../Solutions/RE/Images/Requirement.svg\" id=\"" + parentReqdocRel.getID().ToString() + "\"><userdata key=\"gridData_rowItemID\" value=\"");
            gridStyle.Append(parentReq.getID());
            gridStyle.Append("\" />");

            gridStyle.Append("<td>" + Escape(parentReq.getProperty("req_title", "")) + "</td>");
            gridStyle.Append("<td></td>");
            gridStyle.Append("<td>" + Escape("<checkbox state='0'/>") + "</td>");

            for (int i = 0; i < reqdoc_relationReqs.getItemCount(); i++)
            {
                string subChapter = chapter + "." + (i + 1).ToString();
                Item reqdoc_relationReq = reqdoc_relationReqs.getItemsByXPath("//Item[chapter='" + subChapter + "']");
                if (reqdoc_relationReq.getItemCount() < 1)
                {
                    break;
                }
                if (reqdoc_relationReq.isCollection())
                {
                    throw new Exception("数据结构错误!");
                }
                GetReqDocStructure(reqdoc_relationReqs, reqdoc_relationReq, subChapter, level + 1);
            }

            gridStyle.Append("</tr>");
        }

        private void GetReqDocBaseTamplate()
        {
            //需求对象类ID
            string itemtypeId = "55515617CB224C90AB5A9DAC0F061C2A";

            gridStyle.Append("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
            gridStyle.Append("<table");
            gridStyle.Append(" font=\"Microsoft Sans Serif-8\"");
            gridStyle.Append(" sel_bgColor=\"steelbue\"");
            gridStyle.Append(" sel_TextColor=\"white\"");
            gridStyle.Append(" header_BgColor=\"buttonface\"");
            gridStyle.Append(" expandroot=\"true\"");
            gridStyle.Append(" expandall=\"false\"");
            gridStyle.Append(" treelines=\"1\"");
            gridStyle.Append(" editable=\"true\"");
            gridStyle.Append(" draw_grid=\"true\"");
            gridStyle.Append(" multiselect=\"true\"");
            gridStyle.Append(" column_draggable=\"true\"");
            gridStyle.Append(" enableHtml=\"false\"");
            gridStyle.Append(" enterAsTab=\"false\"");
            gridStyle.Append(" bgInvert=\"true\"");
            gridStyle.Append(" xmlns:msxsl=\"urn:schemas-microsoft-com:xslt\"");
            gridStyle.Append(" xmlns:aras=\"http://www.aras.com\"");
            gridStyle.Append(" xmlns:usr=\"urn:the-xml-files:xslt\">");

            //网格列
            gridStyle.Append("<thead>");
            gridStyle.Append("  <th align=\"c\">" + CCO.Cache.GetPropertyFromCache(itemtypeId, "req_title").GetAttribute("label") + "</th>");
            gridStyle.Append("  <th align=\"c\">执行动作</th>");
            gridStyle.Append("  <th align=\"c\">勾选</th>");
            gridStyle.Append("</thead>");
            gridStyle.Append("<columns>");
            gridStyle.Append("  <column width=\"400\" edit=\"NOEDIT\"  align=\"l\" order=\"0\" />");
            gridStyle.Append("  <column width=\"100\" edit=\"COMBO:0\"  align=\"c\" order=\"1\" />");
            gridStyle.Append("  <column width=\"50\" edit=\"FIELD\"  align=\"c\" order=\"2\" />");
            gridStyle.Append("</columns>");

            gridStyle.Append("<list id='0'>");
            gridStyle.Append("<listitem value='0' label='使用模板' />");
            gridStyle.Append("<listitem value='1' label='另存为' />");
            gridStyle.Append("</list>");
        }

        public Item GetTrackReportGrid(Item reqDoc)
        {
            if (CheckLicense() == false)
            {
                return Cinn.newError(CstrErrMessage);
            }

            string reqdoc_id = reqDoc.getProperty("itemid");
            if (String.IsNullOrEmpty(reqdoc_id))
            {
                return Cinn.newError("未读取到需求文档ID");
            }

            gridStyle = new StringBuilder();

            GetTrackReportBaseTamplate();

            Item reqdoc_relationReqs = Cinn.newItem("re_Req_Doc_Content", "get");
            reqdoc_relationReqs.setAttribute("select", "id,chapter,related_id(id,req_title,major_rev,managed_by_id,current_state)");
            reqdoc_relationReqs.setAttribute("orderBy", "chapter");
            reqdoc_relationReqs.setProperty("source_id", reqdoc_id);
            reqdoc_relationReqs = reqdoc_relationReqs.apply();
            if (reqdoc_relationReqs.getItemCount() > 0)
            {
                for (int i = 0; i < reqdoc_relationReqs.getItemCount(); i++)
                {
                    Item reqdoc_relation = reqdoc_relationReqs.getItemByIndex(i);
                    Item req = reqdoc_relation.getRelatedItem();
                    req.fetchRelationships("re_Requirement_UseCase", "related_id(id,title,major_rev,owned_by_id,current_state,cn_estimate_duration,description,cn_notes)");
                    req.fetchRelationships("re_Requirement_TestCase", "related_id(id,title,major_rev,owned_by_id,current_state,cn_estimate_duration,description,cn_notes)");
                    GetReqDocStructure(req, reqdoc_relation.getProperty("chapter") + " - ");
                }
            }

            gridStyle.Append("</table>");
            return Cinn.newResult(gridStyle.ToString());
        }

        private void GetReqDocStructure(Item req, string chapter)
        {
            string state = req.getProperty("current_state", "");
            string stateName = "";
            string owner = req.getProperty("managed_by_id", "");
            if (owner != "")
            {
                owner = Cinn.getItemById("Identity", owner).getProperty("keyed_name");
            }
            if (state != "")
            {
                Item stateItem = Cinn.getItemById("Life Cycle State", state);
                stateName = stateItem.getProperty("name");
                state = stateItem.getProperty("label", stateName);
            }
            gridStyle.Append("<tr level=\"");
            gridStyle.Append("0");
            gridStyle.Append("\" icon0=\"../Solutions/RE/Images/Requirement.svg\" icon1=\"../Solutions/RE/Images/Requirement.svg\" id=\"" + req.getNewID().ToString() + "\"><userdata key=\"gridData_rowItemID\" value=\"");
            gridStyle.Append(req.getID());
            gridStyle.Append("\" />");
            gridStyle.Append("<td>" + Escape(req.getProperty("req_title", "")) + "</td>");
            gridStyle.Append("<td>" + Escape(req.getProperty("major_rev", "")) + "</td>");
            gridStyle.Append("<td>" + Escape(owner) + "</td>");
            gridStyle.Append("<td></td><td></td><td></td>");
            gridStyle.Append("<td bgColor='" + getStateColor(stateName) + "'>" + Escape(state) + "</td>");

            Item useCaseItems = req.getRelationships("re_Requirement_UseCase");
            for (int i = 0; i < useCaseItems.getItemCount(); i++)
            {
                Item useCaseItem = useCaseItems.getItemByIndex(i).getRelatedItem();
                state = useCaseItem.getProperty("current_state", "");
                stateName = "";
                owner = useCaseItem.getProperty("owned_by_id", "");
                if (owner != "")
                {
                    owner = Cinn.getItemById("Identity", owner).getProperty("keyed_name");
                }
                if (state != "")
                {
                    Item stateItem = Cinn.getItemById("Life Cycle State", state);
                    stateName = stateItem.getProperty("name");
                    state = stateItem.getProperty("label", stateName);
                }
                gridStyle.Append("<tr level=\"");
                gridStyle.Append("1");
                gridStyle.Append("\" icon0=\"../Images/List_2.svg\" icon1=\"../Images/List_2.svg\" id=\"" + useCaseItem.getNewID().ToString() + "\"><userdata key=\"gridData_rowItemID\" value=\"");
                gridStyle.Append(useCaseItem.getID());
                gridStyle.Append("\" />");
                gridStyle.Append("<td>" + Escape(useCaseItem.getProperty("title", "")) + "</td>");
                gridStyle.Append("<td>" + Escape(useCaseItem.getProperty("major_rev", "")) + "</td>");
                gridStyle.Append("<td>" + Escape(owner) + "</td>");
                gridStyle.Append("<td>" + useCaseItem.getProperty("cn_estimate_duration", "") + "</td>");
                gridStyle.Append("<td>" + Escape(useCaseItem.getProperty("description", "")) + "</td>");
                gridStyle.Append("<td>" + Escape(useCaseItem.getProperty("cn_notes", "")) + "</td>");
                gridStyle.Append("<td bgColor='" + getStateColor(stateName) + "'>" + Escape(state) + "</td>");

                useCaseItem.fetchRelationships("re_UseCase_TestCase", "related_id(id,title,major_rev,owned_by_id,current_state,cn_estimate_duration,description,cn_notes)");
                GetUseCaseStructure(useCaseItem.getRelationships("re_UseCase_TestCase"), 2);

                gridStyle.Append("</tr>");
            }

            GetUseCaseStructure(req.getRelationships("re_Requirement_TestCase"), 1);

            gridStyle.Append("</tr>");
        }

        private void GetUseCaseStructure(Item testCaseItems, int level)
        {
            for (int i = 0; i < testCaseItems.getItemCount(); i++)
            {
                Item testCaseItem = testCaseItems.getItemByIndex(i).getRelatedItem();
                string state = testCaseItem.getProperty("current_state", "");
                string stateName = "";
                string owner = testCaseItem.getProperty("owned_by_id", "");
                if (owner != "")
                {
                    owner = Cinn.getItemById("Identity", owner).getProperty("keyed_name");
                }
                if (state != "")
                {
                    Item stateItem = Cinn.getItemById("Life Cycle State", state);
                    stateName = stateItem.getProperty("name");
                    state = stateItem.getProperty("label", stateName);
                }

                gridStyle.Append("<tr level=\"");
                gridStyle.Append(level.ToString());
                gridStyle.Append("\" icon0=\"../Images/WatermarkSet.svg\" icon1=\"../Images/WatermarkSet.svg\" id=\"" + testCaseItem.getNewID().ToString() + "\"><userdata key=\"gridData_rowItemID\" value=\"");
                gridStyle.Append(testCaseItem.getID());
                gridStyle.Append("\" />");
                gridStyle.Append("<td>" + Escape(testCaseItem.getProperty("title", "")) + "</td>");
                gridStyle.Append("<td>" + Escape(testCaseItem.getProperty("major_rev", "")) + "</td>");
                gridStyle.Append("<td>" + Escape(owner) + "</td>");
                gridStyle.Append("<td>" + testCaseItem.getProperty("cn_estimate_duration", "") + "</td>");
                gridStyle.Append("<td>" + Escape(testCaseItem.getProperty("description", "")) + "</td>");
                gridStyle.Append("<td>" + Escape(testCaseItem.getProperty("cn_notes", "")) + "</td>");
                gridStyle.Append("<td bgColor='" + getStateColor(stateName) + "'>" + Escape(state) + "</td>");
                gridStyle.Append("</tr>");
            }
        }

        private void GetTrackReportBaseTamplate()
        {
            gridStyle.Append("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
            gridStyle.Append("<table");
            gridStyle.Append(" font=\"Microsoft Sans Serif-8\"");
            gridStyle.Append(" sel_bgColor=\"steelbue\"");
            gridStyle.Append(" sel_TextColor=\"white\"");
            gridStyle.Append(" header_BgColor=\"buttonface\"");
            gridStyle.Append(" expandroot=\"true\"");
            gridStyle.Append(" expandall=\"false\"");
            gridStyle.Append(" treelines=\"1\"");
            gridStyle.Append(" editable=\"true\"");
            gridStyle.Append(" draw_grid=\"true\"");
            gridStyle.Append(" multiselect=\"true\"");
            gridStyle.Append(" column_draggable=\"true\"");
            gridStyle.Append(" enableHtml=\"false\"");
            gridStyle.Append(" enterAsTab=\"false\"");
            gridStyle.Append(" bgInvert=\"false\"");
            gridStyle.Append(" xmlns:msxsl=\"urn:schemas-microsoft-com:xslt\"");
            gridStyle.Append(" xmlns:aras=\"http://www.aras.com\"");
            gridStyle.Append(" xmlns:usr=\"urn:the-xml-files:xslt\">");

            //网格列
            gridStyle.Append("<thead>");
            gridStyle.Append("  <th align=\"c\">thead.name</th>");
            gridStyle.Append("  <th align=\"c\">thead.major</th>");
            gridStyle.Append("  <th align=\"c\">thead.leader</th>");
            gridStyle.Append("  <th align=\"c\">thead.days</th>");
            gridStyle.Append("  <th align=\"c\">thead.description</th>");
            gridStyle.Append("  <th align=\"c\">thead.remark</th>");
            gridStyle.Append("  <th align=\"c\">thead.state</th>");
            gridStyle.Append("</thead>");
            gridStyle.Append("<columns>");
            gridStyle.Append("  <column width=\"400\" edit=\"NOEDIT\"  align=\"l\" order=\"0\" />");
            gridStyle.Append("  <column width=\"30\" edit=\"NOEDIT\"  align=\"c\" order=\"1\" />");
            gridStyle.Append("  <column width=\"100\" edit=\"NOEDIT\"  align=\"c\" order=\"2\" />");
            gridStyle.Append("  <column width=\"60\" edit=\"NOEDIT\"  align=\"c\" order=\"3\" />");
            gridStyle.Append("  <column width=\"200\" edit=\"NOEDIT\"  align=\"l\" order=\"4\" />");
            gridStyle.Append("  <column width=\"200\" edit=\"NOEDIT\"  align=\"l\" order=\"5\" />");
            gridStyle.Append("  <column width=\"70\" edit=\"NOEDIT\"  align=\"c\" order=\"6\" />");
            gridStyle.Append("</columns>");
        }

        private string getStateColor(string state)
        {
            string color;
            switch (state)
            {
                case "Complete":
                case "Released":
                case "Superseded":
                    color = "#66FF00";
                    break;
                case "":
                    color = "";
                    break;
                default:
                    color = "yellow";
                    break;
            }
            return color;
        }

        public Item FlaggingMessage(Item message)
        {
            if (CheckLicense() == false)
            {
                return Cinn.newError(CstrErrMessage);
            }

            string flagOrUnFlag = message.getProperty("isflag", "true");//true:标记; false:取消标记
            string messageId = message.getProperty("messageid");
            string idenId = Cinn.getUserAliases();
            Item messageItem = Cinn.getItemById("BCS Discussion Item Message", messageId);
            string flagIdentities = messageItem.getProperty("bcs_flag_identities", "");
            if (flagOrUnFlag == "true")
            {
                if (!flagIdentities.Contains(idenId))
                {
                    messageItem.setProperty("bcs_flag_identities", flagIdentities == "" ? idenId : flagIdentities + "," + idenId);
                    return messageItem.apply("edit");
                }
            }
            else
            {
                if (flagIdentities.IndexOf(idenId) == 0)
                {
                    messageItem.setProperty("bcs_flag_identities", flagIdentities == idenId ? "" : flagIdentities.Replace(idenId + ",", ""));
                    return messageItem.apply("edit");
                }
                else if (flagIdentities.IndexOf(idenId) > 0)
                {
                    messageItem.setProperty("bcs_flag_identities", flagIdentities.Replace("," + idenId, ""));
                    return messageItem.apply("edit");
                }
            }
            return message;
        }

        public Item EraseMessage(Item message)
        {
            if (CheckLicense() == false)
            {
                return Cinn.newError(CstrErrMessage);
            }

            string messageId = message.getProperty("messageid");
            string userId = Cinn.getUserID();
            Item messageItem = Cinn.getItemById("BCS Discussion Item Message", messageId);
            if (messageItem == null)
            {
                return Cinn.newError("未查询到此留言!");
            }
            if (messageItem.getProperty("created_by_id") != userId)
            {
                return Cinn.newError("无法删除他人留言!");
            }
            Item result = messageItem.apply("delete");
            if (result.isError())
            {
                return result;
            }

            if (messageItem.getProperty("bcs_parent_message_id", "") == "")
            {
                messageItem = Cinn.newItem("BCS Discussion Item Message", "delete");
                messageItem.setAttribute("where", "bcs_parent_message_id='" + messageId + "'");
                return messageItem.apply();
            }
            return message;
        }

        public Item UseCaseOnPromote(Item thisItem)
        {
            if (CheckLicense() == false)
            {
                return Cinn.newError(CstrErrMessage);
            }

            string[] idlist = thisItem.getAttribute("idlist", "").Split(',');
            Item useCase;
            string assigned = "";

            foreach (string itemId in idlist)
            {
                useCase = Cinn.getItemById("re_UseCase", itemId);
                string state = useCase.getProperty("state");
                switch (state)
                {
                    case "Plan":
                        assigned = Cinn.getItemByKeyedName("Identity", Cinn.getItemById("User", useCase.getProperty("created_by_id")).getProperty("keyed_name")).getID();
                        break;
                    case "Working":
                        assigned = useCase.getProperty("owned_by_id");
                        break;
                    case "Draft":
                    case "InReview":
                        Item memberIden = getOrgLeader(useCase.getProperty("owned_by_id"));
                        if (memberIden.isError())
                        {
                            return memberIden;
                        }
                        assigned = memberIden.getProperty("related_id", "");
                        break;
                    default:
                        continue;
                }
                useCase.setProperty("bcs_assigned", assigned);
                useCase = useCase.apply("edit");
                if (useCase.isError())
                {
                    return useCase;
                }
            }
            return thisItem;
        }

        private Item getOrgLeader(string owner)
        {
            Item orgIden = Cinn.newItem("Identity", "get");
            orgIden.setProperty("bcs_is_org", "1");
            Item memberIden = Cinn.newItem("Member", "get");
            memberIden.setProperty("bcs_is_org", "1");
            memberIden.setProperty("related_id", owner);
            orgIden.addRelationship(memberIden);
            orgIden = orgIden.apply();
            if (orgIden.isError())
            {
                Item ownerItem = Cinn.getItemById("Identity", owner);
                return Cinn.newError(string.Format("無法取得角色所屬組織. 角色 = {0}", ownerItem.getProperty("name", "")));
            }
            memberIden = Cinn.newItem("Member", "get");
            memberIden.setProperty("source_id", orgIden.getID());
            memberIden.setProperty("bcs_is_leader", "1");
            memberIden = memberIden.apply();
            if (memberIden.isError())
            {
                return Cinn.newError("無法取得1階主管");
            }
            return memberIden;
        }

        public Item TestCaseOnPromote(Item thisItem)
        {
            if (CheckLicense() == false)
            {
                return Cinn.newError(CstrErrMessage);
            }

            string[] idlist = thisItem.getAttribute("idlist", "").Split(',');
            Item testCase;
            Item useCase;
            Item requirement;
            Item relation;
            int itemsCount;
            string assigned = "";

            foreach (string itemId in idlist)
            {
                testCase = Cinn.getItemById("re_TestCase", itemId);
                string state = testCase.getProperty("state");
                switch (state)
                {
                    case "Preliminary":
                        assigned = Cinn.getItemByKeyedName("Identity", Cinn.getItemById("User", testCase.getProperty("created_by_id")).getProperty("keyed_name")).getID();
                        break;
                    case "Approve":
                        assigned = testCase.getProperty("owned_by_id");
                        break;
                    case "Valiadate":
                        useCase = Cinn.newItem("re_UseCase", "get");
                        relation = useCase.createRelationship("re_UseCase_TestCase", "get");
                        relation.setRelatedItem(testCase);
                        useCase = useCase.apply();
                        itemsCount = useCase.getItemCount();
                        if (itemsCount < 1)
                        {
                            requirement = Cinn.newItem("re_Requirement", "get");
                            relation = requirement.createRelationship("re_Requirement_TestCase", "get");
                            relation.setRelatedItem(testCase);
                            requirement = requirement.apply();
                            itemsCount = requirement.getItemCount();
                            if (itemsCount < 1)
                            {
                                return Cinn.newError("测试案例未关联至需求或使用案例中");
                            }
                            else if (itemsCount == 1)
                            {
                                assigned = requirement.getProperty("managed_by_id");
                            }
                            else
                            {
                                return Cinn.newError("测试案例不能使用于多个需求中!");
                            }
                        }
                        else if (itemsCount == 1)
                        {
                            assigned = useCase.getProperty("owned_by_id");
                        }
                        else
                        {
                            return Cinn.newError("测试案例不能使用于多个使用案例中!");
                        }
                        break;
                    default:
                        continue;
                }
                testCase.setProperty("bcs_assigned", assigned);
                testCase = testCase.apply("edit");
                if (testCase.isError())
                {
                    return testCase;
                }
            }
            return thisItem;
        }

        public Item CompleteCaseTask(Item completeTask)
        {
            if (CheckLicense() == false)
            {
                return Cinn.newError(CstrErrMessage);
            }

            InnovatorDatabase conn = CCO.DB.InnDatabase;

            string ItemTypeName = completeTask.getType();
            string ItemId = completeTask.getID();
            string VoteAction = completeTask.getAttribute("voteaction");
            string Comment = completeTask.getProperty("bcs_voteComment", "");
            Item TaskItem = Cinn.getItemById(ItemTypeName, ItemId);
            string Assigned = TaskItem.getProperty("bcs_assigned");
            string userId = Cinn.getUserID();
            string userIdentities = Aras.Server.Security.Permissions.GetIdentitiesList(conn, userId);
            if (!userIdentities.Contains(Assigned))
            {
                return Cinn.newError("不能完成他人的任务!");
            }

            PromoteFactory Factory = new PromoteFactory(Cinn, ItemTypeName, ItemId, VoteAction, completeTask.getProperty("state"));
            Item result = Factory.Promote();
            if (result.isError())
            {
                return result;
            }


            return Factory.UpdateComment(conn.GetTableName(ItemTypeName), Comment);
        }

        private class PromoteFactory
        {
            public PromoteFactory(Innovator privateInn, string ItemTypeName, string ItemId, string VoteAction, string state)
            {
                innovator = privateInn;
                Type = ItemTypeName;
                Id = ItemId;
                Vote = VoteAction;
                FromState = state;

                PromoteItem = innovator.newItem(ItemTypeName);
                switch (ItemTypeName)
                {
                    case "re_TestCase":
                        PromoteItem.setPropertyAttribute("Preliminary", "Complete", "Approve");

                        PromoteItem.setPropertyAttribute("Approve", "Complete", "Valiadate");
                        PromoteItem.setPropertyAttribute("Approve", "Reject", "Preliminary");

                        PromoteItem.setPropertyAttribute("Valiadate", "Complete", "Complete");
                        PromoteItem.setPropertyAttribute("Valiadate", "Reject", "Approve");
                        break;
                    case "re_UseCase":
                        PromoteItem.setPropertyAttribute("Plan", "Complete", "Working");

                        PromoteItem.setPropertyAttribute("Working", "Complete", "Draft");
                        PromoteItem.setPropertyAttribute("Working", "Reject", "Plan");

                        PromoteItem.setPropertyAttribute("Draft", "Complete", "InReview");
                        PromoteItem.setPropertyAttribute("Draft", "Reject", "Working");

                        PromoteItem.setPropertyAttribute("InReview", "Complete", "Completed");
                        PromoteItem.setPropertyAttribute("InReview", "Reject", "Draft");
                        break;
                }
            }

            public Innovator innovator;
            public Item PromoteItem;
            public string Type;
            public string Id;
            public string Vote;
            public string FromState;

            public Item Promote()
            {
                Item result = checkRoles();
                if (result.isError())
                {
                    return result;
                }
                result = innovator.newItem(Type, "promote");
                result.setID(Id);
                return result.promote(PromoteItem.getPropertyAttribute(FromState, Vote), "");
            }

            public Item UpdateComment(string TableName, string Comment)
            {
                string lineBrek = "char(13)+char(10)";
                string voteInfo = innovator.getItemById("User", innovator.getUserID()).getProperty("keyed_name") + " " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                string SQL = String.Format("update innovator.{0} set cn_notes=CASE WHEN cn_notes IS NULL THEN N'{3}'+{2}+N'{1}' WHEN cn_notes='' THEN N'{3}'+{2}+N'{1}' ELSE cn_notes+{2}+N'{3}'+{2}+N'{1}' END where id='" + Id + "'", TableName, Comment, lineBrek, voteInfo);
                return innovator.applySQL(SQL);
            }

            public Item checkRoles()
            {
                switch (Type)
                {
                    case "re_UseCase":
                        if (Vote == "Complete" && FromState == "Working")
                        {
                            //负责人完成使用案例Working任务时检查所有子阶测试案例是否都已完成
                            Item testCases = innovator.newItem("re_UseCase_TestCase", "get");
                            testCases.setAttribute("select", "related_id(keyed_name,state)");
                            testCases.setProperty("source_id", Id);
                            testCases = testCases.apply();

                            string errorString = "";
                            for (int i = 0; i < testCases.getItemCount(); i++)
                            {
                                Item testCase = testCases.getItemByIndex(i).getRelatedItem();
                                if (testCase.getProperty("state") != "Complete")
                                {
                                    errorString += String.Format("[{0}]", testCase.getProperty("keyed_name"));
                                }
                            }
                            if (errorString != "")
                            {
                                return innovator.newError("以下测试案例未完成:" + errorString);
                            }
                        }
                        break;
                }

                return innovator.newResult("ok");
            }
        }


        private string Escape(string data)
        {
            return System.Security.SecurityElement.Escape(data);
        }

        /// <summary>
        /// 需求追踪报表
        /// </summary>
        /// <param name="req">需求</param>
        /// <returns></returns>
        public Item getTrackReportToGrid(Item req)
        {
            if (!CheckLicense())
            {
                return Cinn.newError(CstrErrMessage);
            }
            string reqdoc_id = req.getProperty("itemid");
            if (string.IsNullOrEmpty(reqdoc_id))
            {
                return Cinn.newError("未读取到需求文档ID");
            }
            string re_direction = req.getProperty("re_direction");
            gridStyle = new StringBuilder();
            GetTrackReportBaseTamplate();
            GetTrackReportBaseTamplates();
            if (re_direction == "0")
            {
                Item reqdoc_relation2 = Cinn.newItem("re_Requirement", "get");
                reqdoc_relation2.setID(reqdoc_id);
                reqdoc_relation2 = reqdoc_relation2.apply();
                GetReqDataXml(reqdoc_relation2);
            }
            else
            {
                Item reqdoc_relation = Cinn.newItem("re_Requirement", "get");
                reqdoc_relation.setID(reqdoc_id);
                reqdoc_relation = reqdoc_relation.apply();
                getActReDate(reqdoc_relation);
            }
            gridStyle.Append("</table>");
            return Cinn.newResult(gridStyle.ToString());
        }

        /// <summary>
        /// 需求上阶查询
        /// </summary>
        /// <param name="req">需求</param>
        /// <param name="actReqs">项目任务</param>
        private void GetReqDataXml(Item req)
        {
            setReqItemXml(req);
            req.fetchRelationships("re_Requirement_UseCase", "related_id(id,title,major_rev,owned_by_id,current_state,cn_estimate_duration,description,cn_notes)");
            req.fetchRelationships("re_Requirement_TestCase", "related_id(id,title,major_rev,owned_by_id,current_state,cn_estimate_duration,description,cn_notes)");
            req.fetchRelationships("re_Requirement_Out_Link", "related_id(id,req_title,major_rev,managed_by_id,current_state)");
            Item useCaseItems = req.getRelationships("re_Requirement_UseCase");
            for (int j = 0; j < useCaseItems.getItemCount(); j++)
            {
                setUseCaseItemXml(useCaseItems.getItemByIndex(j));
            }
            GetUseCaseStructure(req.getRelationships("re_Requirement_TestCase"), 1);
            Item re_Out_Link = req.getRelationships("re_Requirement_Out_Link");
            for (int i = 0; i < re_Out_Link.getItemCount(); i++)
            {
                GetReqDataXml(re_Out_Link.getItemByIndex(i).getRelatedItem());
            }
            gridStyle.Append("</tr>");
        }

        private void getActReDate(Item req)
        {
            setReqItemXml(req);
            Item actReqs = Cinn.applySQL("select * from innovator.ACTIVITY2 A left join innovator.ACTIVITY2_REQUIREMENT B on A.id=B.SOURCE_ID where b.RELATED_ID='" + req.getID() + "'");
            for (int j = 0; j < actReqs.getItemCount(); j++)
            {
                setActItemXml(actReqs.getItemByIndex(j));
            }
            Item reqdoc_relation = Cinn.newItem("re_Requirement_Out_Link", "get");
            reqdoc_relation.setProperty("related_id", req.getID());
            Item reqdoc_relations = Cinn.newItem("re_Requirement", "get");
            reqdoc_relations.addRelationship(reqdoc_relation);
            reqdoc_relations = reqdoc_relations.apply();
            for (int i = 0; i < reqdoc_relations.getItemCount(); i++)
            {
                getActReDate(reqdoc_relations.getItemByIndex(i));
            }
            gridStyle.Append("</tr>");
        }

        private string getWbsID(string webID)
        {
            Item wbsAct = Cinn.newItem("sub wbs", "get");
            wbsAct.setProperty("related_id", webID);
            wbsAct = wbsAct.apply();
            if (!wbsAct.isEmpty())
            {
                return getWbsID(wbsAct.getProperty("source_id"));
            }
            return webID;
        }


        public Item RE_Export_Excel(Item req, dynamic CCOS)
        {
            if (CheckLicense() == false)
            {
                return Cinn.newError(CstrErrMessage);
            }

            string reqdoc_id = req.getProperty("itemid");
            if (String.IsNullOrEmpty(reqdoc_id))
            {
                return Cinn.newError("未读取到需求文档ID");
            }
            string gridXMLS = req.node.LastChild.InnerXml;
            if (String.IsNullOrEmpty(gridXMLS))
            {
                return Cinn.newError("页面无内容导出");
            }
            string rurl = CCOS.Server.MapPath("../Client");
            DataTable dt = HtmlToDataTable(gridXMLS, rurl);
            string filePath = CCOS.Server.MapPath("../Client/scripts/WebEditor/ueditor/TemporaryFile/") + Cinn.getNewID() + ".xlsx";

            return NPOICreateExcel(dt, filePath);
        }

        public static DataTable HtmlToDataTable(string html, string rurl)
        {
            const string nulltxt = "everything is ok";
            DataTable dt = new DataTable();
            HtmlDocument doc = new HtmlDocument();
            doc.LoadHtml(html);
            var tList = doc.DocumentNode.SelectNodes("//table");
            if (tList != null)
            {
                var table = tList[0];
                var rows = table.SelectNodes("//tr");
                var Columns = table.SelectNodes("//th");
                if (rows != null)
                {
                    var colCount = 0;
                    foreach (var td in rows[0].ChildNodes.Where(m => m.OriginalName.ToLower() == "td"))
                    {
                        var attr = td.Attributes["colspan"];
                        var colspan = (attr != null) ? int.Parse(attr.Value) : 1;
                        colCount = colCount + colspan;
                    }
                    var rowCount = rows.Count;
                    var arr = new string[rowCount][];
                    for (var r = 0; r < rowCount; r++)
                    {
                        arr[r] = new string[colCount];
                    }
                    //填充数据
                    for (var row = 0; row < rowCount; row++)
                    {
                        //获取图标目录并存入list
                        string iconU = rows[row].GetAttributes("icon0").First().Value.Replace(".svg", ".png").Replace("../", "").Replace("/", "\\");
                        iconList.Add(rurl + "\\" + iconU);

                        var tr = rows[row];
                        var cols = tr.ChildNodes.Where(m => m.OriginalName.ToLower() == "td").ToList();
                        for (var column = 0; column < cols.Count; column++)
                        {
                            var cAttr = cols[column].Attributes["colspan"];
                            var colspan = (cAttr != null) ? int.Parse(cAttr.Value) : 1;
                            var rAttr = cols[column].Attributes["rowspan"];
                            var rowspan = (rAttr != null) ? int.Parse(rAttr.Value) : 1;
                            var text = string.IsNullOrEmpty(cols[column].InnerText) ? nulltxt : cols[column].InnerText;
                            var startColumn = 0;
                            for (var i = 0; i < rowspan; i++)
                            {
                                for (var j = 0; j < colspan; j++)
                                {
                                    var d = startColumn == 0 ? column : startColumn;
                                    if (string.IsNullOrEmpty(arr[row + i][d + j]))
                                        arr[row + i][d + j] = text;
                                    else
                                    {
                                        var t = column + j + 1;
                                        startColumn = t;
                                        while (true)
                                        {
                                            if (string.IsNullOrEmpty(arr[row][t]))
                                            {
                                                arr[row][t] = text;
                                                break;
                                            }
                                            t++;
                                        }
                                    }
                                }
                            }
                        }
                    }
                    foreach (var th in Columns)
                    {
                        dt.Columns.Add(th.InnerText);
                    }
                    for (var i = 0; i < arr.Length; i++)
                    {
                        //if (i == 0)
                        //{
                        //    for (var j = 0; j < arr[i].Length; j++)
                        //    {
                        //        var columnTxt = arr[i][j] == nulltxt ? "Column" + j : arr[i][j];
                        //        dt.Columns.Add(columnTxt);
                        //    }
                        //}

                        var row = dt.NewRow();
                        for (var k = 0; k < arr[i].Length; k++)
                        {
                            var columnTxt = arr[i][k] == nulltxt ? "" : arr[i][k];
                            row[k] = columnTxt;
                        }
                        dt.Rows.Add(row);
                    }
                }
            }
            return dt;
        }

        /// <summary>
        /// 使用npoi创建一个excel文件
        /// </summary>
        public Item NPOICreateExcel(DataTable dt, string filePath)
        {
            //声明一个工作簿
            XSSFWorkbook workBook = new XSSFWorkbook();
            //创建一个sheet页
            ISheet sheet = workBook.CreateSheet("MySheet");
            //向第一行第一列的单元格添加文本

            //样式
            ICellStyle headCellStyle = workBook.CreateCellStyle();
            headCellStyle.Alignment = HorizontalAlignment.Center;
            headCellStyle.FillForegroundColor = IndexedColors.Grey40Percent.Index;
            headCellStyle.FillPattern = FillPattern.SolidForeground;


            ICellStyle cellStyle1 = workBook.CreateCellStyle();
            cellStyle1.Alignment = HorizontalAlignment.Center;
            cellStyle1.FillForegroundColor = IndexedColors.Yellow.Index;
            cellStyle1.FillPattern = FillPattern.SolidForeground;

            ICellStyle cellStyle2 = workBook.CreateCellStyle();
            cellStyle2.Alignment = HorizontalAlignment.Center;

            //导入excel标题
            IRow rowHeader = sheet.CreateRow(0);
            if (rowHeader == null)//workbook 创建的sheet里是获取不到对应的excel行和列的单元格对象
            {
                rowHeader = sheet.CreateRow(0);
            }

            ICell tcell = rowHeader.CreateCell(0);
            tcell.CellStyle = headCellStyle;
            tcell.SetCellValue("图示");

            for (int i = 0; i < dt.Columns.Count + 1; i++)
            {
                ICell cell = rowHeader.CreateCell(i);
                cell.CellStyle = headCellStyle;
                cell.SetCellValue(dt.Columns[i-1].ColumnName);
            }

            sheet.SetColumnWidth(0, 4 * 256);

            //导入excel内容
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                IRow row = sheet.GetRow(i + 1);//获取第二行
                if (row == null)//workbook 创建的sheet里是获取不到对应的excel行和列的单元格对象
                {
                    row = sheet.CreateRow(i + 1);
                }
                for (int j = 0; j < dt.Columns.Count + 1; j++)
                {
                    if (j == 0)
                    {
                        Picture(workBook, sheet, i + 1, iconList[i]);
                        continue;
                    }
                    ICell cell = row.GetCell(j);//获取第一列
                    if (cell == null)
                    {
                        cell = row.CreateCell(j);
                    }
                    cell.CellStyle = cellStyle2;

                    if (j == dt.Columns.Count)
                    {
                        cell.CellStyle = cellStyle1;
                    }
                    cell.SetCellValue(dt.Rows[i][j-1].ToString());
                }
            }
            for (int i = 1; i < dt.Columns.Count + 1; i++)
            {
                sheet.AutoSizeColumn(i);
            }

            //清空图标list
            iconList.Clear();

            //输出excel文件
            using (FileStream fs = File.OpenWrite(filePath))
            {
                workBook.Write(fs);//向打开的这个xls文件中写入并保存。
                //上传到aras
                Item file = Cinn.newItem("File", "add");
                file.setProperty("filename", "需求追踪报表导出.xlsx");
                file.attachPhysicalFile(filePath);
                file = file.apply();

                //删除文件
                File.Delete(filePath);

                //返回File
                return file;
            }
        }

        //excel插入图片
        public static void Picture(IWorkbook workbook, ISheet sheet, int row, string dPath)
        {
            //本地版:无法使用相对路径                               
            System.Drawing.Image imgOutput = System.Drawing.Bitmap.FromFile(dPath);
            //设置大小
            //System.Drawing.Image img = imgOutput.GetThumbnailImage(20, 20, null, IntPtr.Zero);
            //图片转换为文件流
            MemoryStream ms = new MemoryStream();
            imgOutput.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
            BinaryReader br = new BinaryReader(ms);
            var picBytes = ms.ToArray();
            ms.Close();
            int pictureIdx = workbook.AddPicture(picBytes, NPOI.SS.UserModel.PictureType.PNG);  //添加图片
            XSSFDrawing drawing = (XSSFDrawing)sheet.CreateDrawingPatriarch();
            XSSFClientAnchor anchor = new XSSFClientAnchor(0, 0, 0, 0, 0, row, 1, row + 1);

            XSSFPicture picture = (XSSFPicture)drawing.CreatePicture(anchor, pictureIdx);
            //picture.Resize();

        }
        public Item generateXML4Item(Item req)
        {
            if (!CheckLicense())
            {
                return Cinn.newError(CstrErrMessage);
            }
            string re_direction = req.getProperty("re_direction");
            string reqdoc_id = req.getProperty("itemid");
            string tablexml = req.node.SelectSingleNode("tablexml").InnerXml;
            gridStyle = new StringBuilder();
            gridStyle.Append(tablexml.Substring(0, tablexml.Length - 8));
            //inputRowData(req);
            if (re_direction == "0")
            {
                Item reqdoc_relation3 = Cinn.newItem("re_Requirement", "get");
                reqdoc_relation3.setID(reqdoc_id);
                reqdoc_relation3 = reqdoc_relation3.apply();
                reqdoc_relation3.fetchRelationships("re_Requirement_UseCase", "related_id(id,title,major_rev,owned_by_id,current_state,cn_estimate_duration,description,cn_notes)");
                reqdoc_relation3.fetchRelationships("re_Requirement_TestCase", "related_id(id,title,major_rev,owned_by_id,current_state,cn_estimate_duration,description,cn_notes)");
            }
            else
            {
                bool isReqXml = false;
                Item actReqs = Cinn.newItem("Activity2", "get");
                Item actReqRd = Cinn.newItem("Activity2 Requirement", "get");
                actReqRd.setProperty("related_id", reqdoc_id);
                actReqs.addRelationship(actReqRd);
                actReqs = getSearch(req, actReqs, req.getProperty("re_direction"));
                if (actReqs.getItemCount() > 0)
                {
                    for (int i = 0; i < actReqs.getItemCount(); i++)
                    {
                        setActItemXmlSearch(req, actReqs.getItemByIndex(i), req.getProperty("re_direction"));
                    }
                    if (remarkBool)
                    {
                        Item reqdoc_relation2 = Cinn.newItem("re_Requirement", "get");
                        reqdoc_relation2.setID(reqdoc_id);
                        reqdoc_relation2 = getSearch(req, reqdoc_relation2, "66");
                        if (reqdoc_relation2.getItemCount() > 0)
                        {
                            setActItemXmlSearch(req, reqdoc_relation2, "66");
                            isReqXml = true;
                        }
                    }
                    else
                    {
                        isReqXml = true;
                    }
                }
                else
                {
                    Item reqdoc_relation = Cinn.newItem("re_Requirement", "get");
                    reqdoc_relation.setID(reqdoc_id);
                    reqdoc_relation = getSearch(req, reqdoc_relation, "66");
                    if (reqdoc_relation.getItemCount() > 0)
                    {
                        setReqItemXml(reqdoc_relation);
                        isReqXml = true;
                    }
                }
                if (isReqXml)
                {
                    gridStyle.Append("</tr>");
                }
            }
            return Cinn.newResult(gridStyle.ToString());
        }

        private Item getSearch(Item inputRows, Item reqRd, string re_direction)
        {
            string theadname, theadmajor, theaddays, theaddescription, theadremark;
            theadname = inputRows.getProperty("theadname");
            theadmajor = inputRows.getProperty("theadmajor");
            theaddays = inputRows.getProperty("theaddays");
            theaddescription = inputRows.getProperty("theaddescription");
            theadremark = inputRows.getProperty("theadremark");
            if (!String.IsNullOrEmpty(theadname) || !String.IsNullOrEmpty(theadmajor) || !String.IsNullOrEmpty(theaddays) || !String.IsNullOrEmpty(theaddescription) || !String.IsNullOrEmpty(theadremark))
            {
                if (!String.IsNullOrEmpty(theadname))
                {
                    switch (re_direction)
                    {
                        case "0":
                            break;
                        case "1":
                            reqRd.setProperty("name", theadname);
                            if (wildCard(theadname))
                            {
                                reqRd.setPropertyAttribute("name", "condition", "like");
                            }
                            break;
                        case "66":
                            reqRd.setProperty("req_title", theadname);
                            if (wildCard(theadname))
                            {
                                reqRd.setPropertyAttribute("req_title", "condition", "like");
                            }
                            break;
                    }
                }
                if (!String.IsNullOrEmpty(theadmajor))
                {
                    switch (re_direction)
                    {
                        case "0":
                            break;
                        case "1":
                            return Cinn.newError("项目任务无版本");
                        case "66":
                            reqRd.setProperty("major_rev", theadmajor);
                            if (wildCard(theadmajor))
                            {
                                reqRd.setPropertyAttribute("major_rev", "condition", "like");
                            }
                            break;
                    }
                }
                if (!String.IsNullOrEmpty(theaddays))
                {
                    switch (re_direction)
                    {
                        case "0":
                            break;
                        case "1":
                            reqRd.setProperty("expected_duration", theaddays);
                            if (wildCard(theaddays))
                            {
                                reqRd.setPropertyAttribute("expected_duration", "condition", "like");
                            }
                            break;
                        case "66":
                            return Cinn.newError("需求无预计工时");
                    }
                }
                if (!String.IsNullOrEmpty(theaddescription))
                {
                    switch (re_direction)
                    {
                        case "0":
                            break;
                        case "1":
                            return Cinn.newError("项目任务无案例说明");
                        case "66":
                            return Cinn.newError("需求无案例说明");
                    }
                }
                if (!String.IsNullOrEmpty(theadremark))
                {
                    switch (re_direction)
                    {
                        case "0":
                            break;
                        case "1":
                            break;
                        case "66":
                            return Cinn.newError("需求无备注说明");
                    }
                }
            }
            return reqRd.apply();
        }

        private bool wildCard(string thead)
        {
            if (thead.Contains("%") || thead.Contains("*"))
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        private string wildcard(string thead)
        {
            if (thead.Contains("%"))
            {
                thead = thead.Replace("%", "_");
                if (thead.Contains("*"))
                {
                    thead = thead.Replace("*", "%");
                    return thead;
                }
                return thead;
            }
            if (thead.Contains("*"))
            {
                thead = thead.Replace("*", "%");
            }
            return thead;
        }

        private bool wildItemKeyedName(string theadA, string type, string id)
        {
            Item wildItem;
            if (id == "")
            {
                return true;
            }
            switch (type)
            {
                case "identity":
                    wildItem = Cinn.applySQL("select id from innovator.[IDENTITY] where id='" + id + "' and KEYED_NAME like N'" + wildcard(theadA) + "'");
                    if (wildItem.getItemCount() < 0)
                    {
                        return true;
                    }
                    return false;
                case "lcs":
                    string liclang = Cinn.getI18NSessionContext().GetLanguageSuffix();
                    wildItem = Cinn.applySQL("select id from innovator.[LIFE_CYCLE_STATE] where id='" + id + "' and [LABEL" + liclang + "] like N'" + wildcard(theadA) + "'");
                    if (wildItem.getItemCount() < 0)
                    {
                        return true;
                    }
                    return false;
                case "prj":
                    wildItem = Cinn.applySQL("select name from innovator.PROJECT where wbs_id='" + id + "' and NAME like N'" + wildcard(theadA) + "'");
                    if (wildItem.getItemCount() < 0)
                    {
                        return true;
                    }
                    return false;
            }
            return true;
        }

        private void setReqItemXml(Item req)
        {
            string state = req.getProperty("current_state", "");
            string stateName = "";
            string owner = req.getProperty("managed_by_id", "");
            if (owner != "")
            {
                owner = Cinn.getItemById("Identity", owner).getProperty("keyed_name");
            }
            if (state != "")
            {
                Item stateItem = Cinn.getItemById("Life Cycle State", state);
                stateName = stateItem.getProperty("name");
                state = stateItem.getProperty("label", stateName);
            }
            gridStyle.Append("<tr level=\"");
            gridStyle.Append("0");
            gridStyle.Append("\" icon0=\"../Solutions/RE/Images/Requirement.svg\" icon1=\"../Solutions/RE/Images/Requirement.svg\" id=\"" + req.getNewID().ToString() + "\"><userdata key=\"gridData_rowItemID\" value=\"");
            gridStyle.Append(req.getID());
            gridStyle.Append("\" />");
            gridStyle.Append("<td>" + Escape(req.getProperty("req_title", "")) + "</td>");
            gridStyle.Append("<td>" + Escape(req.getProperty("major_rev", "")) + "</td>");
            gridStyle.Append("<td>" + Escape(owner) + "</td>");
            gridStyle.Append("<td></td><td></td><td></td>");
            gridStyle.Append("<td bgColor='" + getStateColor(stateName) + "'>" + Escape(state) + "</td>");
        }

        private void setActItemXml(Item actReq)
        {
            string state = actReq.getProperty("current_state", "");
            string stateName = "";
            string actProjectName = "";
            string owner = actReq.getProperty("managed_by_id", "");
            if (owner != "")
            {
                owner = Cinn.getItemById("Identity", owner).getProperty("keyed_name");
            }
            if (state != "")
            {
                Item stateItem = Cinn.getItemById("Life Cycle State", state);
                stateName = stateItem.getProperty("name");
                state = stateItem.getProperty("label", stateName);
            }
            Item wbsAct = Cinn.newItem("WBS Activity2", "get");
            wbsAct.setProperty("related_id", actReq.getID());
            wbsAct = wbsAct.apply();
            Item actProject = Cinn.newItem("Project", "get");
            if (!wbsAct.isEmpty())
            {
                string wbsActID = wbsAct.getProperty("source_id");
                string wbs_id = getWbsID(wbsActID);
                actProject.setProperty("wbs_id", wbs_id);
                actProject = actProject.apply();
                actProjectName = actProject.getProperty("name", "");
            }
            gridStyle.Append("<tr level=\"");
            gridStyle.Append("1");
            gridStyle.Append("\" icon0=\"../Images/Activity2.svg\" icon1=\"../Images/Activity2.svg\" id=\"" + actReq.getNewID().ToString() + "\"><userdata key=\"gridData_rowItemID\" value=\"");
            gridStyle.Append(actReq.getID());
            gridStyle.Append("\" />");
            gridStyle.Append("<td>" + Escape(actReq.getProperty("name", "")) + "</td>");
            gridStyle.Append("<td></td>");
            gridStyle.Append("<td>" + Escape(owner) + "</td>");
            gridStyle.Append("<td>" + actReq.getProperty("expected_duration", "") + "</td>");
            gridStyle.Append("<td></td>");
            gridStyle.Append("<td>" + Escape(actProjectName) + "</td>");
            gridStyle.Append("<td bgColor='" + getStateColor(stateName) + "'>" + Escape(state) + "</td>");
            gridStyle.Append("</tr>");
        }

        private void setUseCaseItemXml(Item req)
        {
            Item useCaseItem = req.getRelatedItem();
            string state = useCaseItem.getProperty("current_state", "");
            string stateName = "";
            string owner = useCaseItem.getProperty("owned_by_id", "");
            if (owner != "")
            {
                owner = Cinn.getItemById("Identity", owner).getProperty("keyed_name");
            }
            if (state != "")
            {
                Item stateItem = Cinn.getItemById("Life Cycle State", state);
                stateName = stateItem.getProperty("name");
                state = stateItem.getProperty("label", stateName);
            }
            gridStyle.Append("<tr level=\"");
            gridStyle.Append("1");
            gridStyle.Append("\" icon0=\"../Images/List_2.svg\" icon1=\"../Images/List_2.svg\" id=\"" + useCaseItem.getNewID().ToString() + "\"><userdata key=\"gridData_rowItemID\" value=\"");
            gridStyle.Append(useCaseItem.getID());
            gridStyle.Append("\" />");
            gridStyle.Append("<td>" + Escape(useCaseItem.getProperty("title", "")) + "</td>");
            gridStyle.Append("<td>" + Escape(useCaseItem.getProperty("major_rev", "")) + "</td>");
            gridStyle.Append("<td>" + Escape(owner) + "</td>");
            gridStyle.Append("<td>" + useCaseItem.getProperty("cn_estimate_duration", "") + "</td>");
            gridStyle.Append("<td>" + Escape(useCaseItem.getProperty("description", "")) + "</td>");
            gridStyle.Append("<td>" + Escape(useCaseItem.getProperty("cn_notes", "")) + "</td>");
            gridStyle.Append("<td bgColor='" + getStateColor(stateName) + "'>" + Escape(state) + "</td>");
            useCaseItem.fetchRelationships("re_UseCase_TestCase", "related_id(id,title,major_rev,owned_by_id,current_state,cn_estimate_duration,description,cn_notes)");
            GetUseCaseStructure(useCaseItem.getRelationships("re_UseCase_TestCase"), 2);
            gridStyle.Append("</tr>");
        }

        private bool remarkBool = true;
        private void setActItemXmlSearch(Item inputRows, Item reqRd, string re_direction)
        {
            Item reqdoc_relation = Cinn.newItem("re_Requirement", "get");
            reqdoc_relation.setID(inputRows.getProperty("itemid"));
            string theadleader, theadstate, theadremark;
            theadleader = inputRows.getProperty("theadleader");
            theadstate = inputRows.getProperty("theadstate");
            theadremark = inputRows.getProperty("theadremark");
            if (!String.IsNullOrEmpty(theadleader) || !String.IsNullOrEmpty(theadstate) || !String.IsNullOrEmpty(theadremark))
            {

                if (!String.IsNullOrEmpty(theadleader))
                {
                    switch (re_direction)
                    {
                        case "0":
                        case "1":
                        case "66":
                            string owner = reqRd.getProperty("managed_by_id", "");
                            if (wildItemKeyedName(theadleader, "identity", owner))
                            {
                                return;
                            }
                            break;
                    }
                }
                if (!String.IsNullOrEmpty(theadstate))
                {
                    switch (re_direction)
                    {
                        case "0":
                        case "1":
                        case "66":
                            string stateID = reqRd.getProperty("current_state", "");
                            if (wildItemKeyedName(theadstate, "lcs", stateID))
                            {
                                return;
                            }
                            break;
                    }
                }
                if (!String.IsNullOrEmpty(theadremark))
                {
                    switch (re_direction)
                    {
                        case "0":
                        case "1":
                            Item wbsAct = Cinn.newItem("WBS Activity2", "get");
                            wbsAct.setProperty("related_id", reqRd.getID());
                            wbsAct = wbsAct.apply();
                            if (wbsAct.getItemCount() > 0)
                            {
                                string wbsActID = wbsAct.getProperty("source_id");
                                string wbs_id = getWbsID(wbsActID);
                                if (wildItemKeyedName(theadremark, "prj", wbs_id))
                                {
                                    return;
                                }
                            }
                            else
                            {
                                return;
                            }
                            break;
                        case "66":
                            return;
                    }
                }
                switch (re_direction)
                {
                    case "0":
                        break;
                    case "1":
                        if (remarkBool)
                        {
                            reqdoc_relation = reqdoc_relation.apply();
                            setReqItemXml(reqdoc_relation);
                            setActItemXml(reqRd);
                            remarkBool = false;
                            return;
                        }
                        setActItemXml(reqRd);
                        return;
                    case "66":
                        reqdoc_relation = reqdoc_relation.apply();
                        setReqItemXml(reqdoc_relation);
                        return;
                }
            }
            else
            {
                if (remarkBool)
                {
                    reqdoc_relation = reqdoc_relation.apply();
                    setReqItemXml(reqdoc_relation);
                    setActItemXml(reqRd);
                    remarkBool = false;
                }
                else
                {
                    setActItemXml(reqRd);
                }

            }
        }

        /// <summary>
        /// 搜索列
        /// </summary>
        private void GetTrackReportBaseTamplates()
        {
            //网格列
            gridStyle.Append("<inputrow>");
            gridStyle.Append("  <td bgColor='#BDDEF7' />");
            gridStyle.Append("  <td bgColor='#BDDEF7' />");
            gridStyle.Append("  <td bgColor='#BDDEF7' />");
            gridStyle.Append("  <td bgColor='#BDDEF7' />");
            gridStyle.Append("  <td bgColor='#BDDEF7' />");
            gridStyle.Append("  <td bgColor='#BDDEF7' />");
            gridStyle.Append("  <td bgColor='#BDDEF7' />");
            gridStyle.Append("</inputrow>");
        }

        //private void inputRowData(Item data)
        //{
        //    inputRow = new List<inputRowClass>();
        //    inputRow.Add(new inputRowClass
        //    {
        //        inputRowName = "1",
        //        inputRowValue = "2"
        //    });
        //}

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
