using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aras.IOM;
using Aras.Server.Core;
using bwInnovatorCore;

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
        //git测试上传

        private string Escape(string data)
        {
            return System.Security.SecurityElement.Escape(data);
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
