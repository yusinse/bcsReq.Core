using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Aras.IOM;
using Aras.Server.Core;
using NPOI.XWPF.UserModel;
using NPOI.OpenXmlFormats.Wordprocessing;
using Newtonsoft.Json;
using bwInnovatorCore;
using System.IO;
using System.Xml;

namespace bcsReq.Core
{
    public class ExportWordFactory
    {
        #region "                   宣告區"

        protected Innovator Cinn;
        protected CInnovator.bwGeneric CbwGeneric = new CInnovator.bwGeneric(); //Modify by kenny 2019/04/11
        protected CGeneric.Common CoCommon;//Modify by kenny 2019/04/11
        private string CstrErrMessage = "";

        protected Innovator innovator { get; private set; }
        protected IServerConnection conn { get; private set; }
        protected string LangCode { get; private set; }

        private RequirementStyle ReqStyle;
        private XWPFDocument doc;
        //文档内表格数量
        private int TableCount = 0;
        //文档内列表数量
        private int abstractNumIdCount = 0;
        //临时文件路径
        private string downloadDir;
        //模板文件路径
        private string templateDir;

        #endregion

        #region "                   進入點"

        public ExportWordFactory()
        {
            Cinn = new Innovator(null);
        }


        public ExportWordFactory(Innovator getInnovator,CallContext CCO)
        {
            //System.Diagnostics.Debugger.Break();

            innovator = getInnovator;
            conn = innovator.getConnection();
            downloadDir= CCO.Server.MapPath("../Client/Solutions/RE/TemporaryFiles/"+innovator.getNewID());
            //templateDir= CCO.Server.MapPath("../Client/Solutions/RE/TemplateFiles/模板.docx");
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
        
        public Item Export2Word(Item exportItem)
        {
            if (CheckLicense() == false)
            {
                return Cinn.newError(CstrErrMessage);
            }

            try
            {
                //System.Diagnostics.Debugger.Break();
                string itemId= exportItem.getID();
                string itemtypeName= exportItem.getType();
                string templateId = exportItem.getProperty("bcs_export_template");
                Item fileItem =innovator.newError("对象类名称错误");
                switch (itemtypeName.ToLower())
                {
                    case "re_requirement_document":
                        fileItem = ExportReqDoc2Word(itemId, itemtypeName, templateId);
                        break;
                    case "re_requirement":
                        fileItem = ExportReq2Word(itemId, itemtypeName, templateId);
                        break;
                }

                return fileItem;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private Item ExportReqDoc2Word(string itemId, string itemtypeName,string templateId)
        {
            try
            {
                //读取需求文档数据
                Item reqDoc = innovator.getItemById(itemtypeName, itemId);
                string keyed_name = reqDoc.getProperty("keyed_name");

                //读取需求文档数据关系数据
                Item reqDocRels = innovator.newItem("re_Req_Doc_Content", "get");
                reqDocRels.setAttribute("select", "chapter,related_id(req_title,content)");
                reqDocRels.setProperty("source_id", itemId);
                Item reqItem = innovator.newItem("re_Requirement", "get");
                reqItem.setAttribute("language", "en");
                reqItem.setAttribute("select", "req_title,content");
                reqDocRels.setRelatedItem(reqItem);
                //reqDocRels.setAttribute("language", "en");
                reqDocRels = reqDocRels.apply();

                //读取Aras样式
                ReqStyle = new RequirementStyle();

                //创建临时目录
                Directory.CreateDirectory(downloadDir);

                //下载模板文件
                templateDir = downloadTemplateFile(templateId);

                // 读取模板文档
                //XWPFDocument template = new XWPFDocument(new FileStream(templateDir, FileMode.Open));
                doc = new XWPFDocument(new FileStream(templateDir, FileMode.Open));
                // 获得模板文档的整体样式
                //var wordStyles = template.GetCTStyle();

                //创建新的word文档
                //doc = new XWPFDocument();

                //设置文档样式
                //XWPFStyles newStyles = doc.CreateStyles();
                //newStyles.SetStyles(wordStyles);
                //创建列表
                doc.CreateNumbering();

                //设置Word文档标题
                XWPFParagraph titlePara = doc.CreateParagraph();
                titlePara.Alignment = ParagraphAlignment.CENTER;
                XWPFRun titleRun = titlePara.CreateRun();
                titleRun.SetText(reqDoc.getProperty("reqdoc_title", ""));
                titleRun.IsBold = true;
                titleRun.FontSize = 18;

                //创建Word目录
                TOC tOC = new TOC(doc.Document.body.AddNewSdt());

                //生成Word内容
                generateWordContent(reqDocRels, "", 1);

                //生成Word目录
                CreateTOC(tOC);

                //生成Word文件
                string fileName = ReplaceBadCharOfFileName(keyed_name) + ".docx";
                string filePath = downloadDir + "\\" + fileName;
                FileStream sw = File.Create(filePath);
                doc.Write(sw);
                sw.Close();
                //将文件上传至Aras
                Item fileItem = innovator.newItem("File", "add");
                fileItem.setProperty("filename", fileName);
                fileItem.attachPhysicalFile(filePath);
                fileItem = fileItem.apply();
                //删除临时文件与临时文件夹
                File.Delete(filePath);
                File.Delete(templateDir);
                Directory.Delete(downloadDir);

                return fileItem;
            }
            catch (Exception ex)
            {
                if (Directory.Exists(downloadDir))
                {
                    foreach (string f in Directory.GetFiles(downloadDir))
                    {
                        if (File.Exists(f))
                        {
                            File.Delete(f);
                        }
                    }
                    Directory.Delete(downloadDir);
                }
                throw ex;
            }
        }

        private Item ExportReq2Word(string itemId, string itemtypeName,string templateId)
        {
            try
            {
                //读取需求数据
                //Item req = innovator.getItemById(itemtypeName, itemId);
                Item req = innovator.newItem(itemtypeName, "get");
                req.setID(itemId);
                req.setAttribute("select", "keyed_name,req_title,content");
                req.setAttribute("language", "en");
                req = req.apply();
                string keyed_name = req.getProperty("keyed_name");

                //读取Aras样式
                ReqStyle = new RequirementStyle();

                //创建临时目录
                Directory.CreateDirectory(downloadDir);

                //下载模板文件
                templateDir = downloadTemplateFile(templateId);

                // 读取模板文档
                XWPFDocument template = new XWPFDocument(new FileStream(templateDir, FileMode.Open));
                // 获得模板文档的整体样式
                var wordStyles = template.GetCTStyle();

                //创建新的word文档
                doc = new XWPFDocument();

                //设置文档样式
                XWPFStyles newStyles = doc.CreateStyles();
                newStyles.SetStyles(wordStyles);
                //创建列表
                doc.CreateNumbering();

                //设置Word文档标题
                XWPFParagraph titlePara = doc.CreateParagraph();
                titlePara.Alignment = ParagraphAlignment.CENTER;
                XWPFRun titleRun = titlePara.CreateRun();
                titleRun.SetText(req.getProperty("req_title", ""));
                titleRun.IsBold = true;
                titleRun.FontSize = 18;

                //生成Word内容
                createContent(req);

                //生成Word文件
                string fileName = ReplaceBadCharOfFileName(keyed_name) + ".docx";
                string filePath = downloadDir + "\\" + fileName;
                FileStream sw = File.Create(filePath);
                doc.Write(sw);
                sw.Close();
                //将文件上传至Aras
                Item fileItem = innovator.newItem("File", "add");
                fileItem.setProperty("filename", fileName);
                fileItem.attachPhysicalFile(filePath);
                fileItem = fileItem.apply();
                //删除临时文件与临时文件夹
                File.Delete(filePath);
                File.Delete(templateDir);
                Directory.Delete(downloadDir);

                return fileItem;
            }
            catch (Exception ex)
            {
                if (Directory.Exists(downloadDir))
                {
                    foreach (string f in Directory.GetFiles(downloadDir))
                    {
                        if (File.Exists(f))
                        {
                            File.Delete(f);
                        }
                    }
                    Directory.Delete(downloadDir);
                }
                throw ex;
            }
        }

        private void generateWordContent(Item reqDocRels, string chapter, int level)
        {
            for (int i = 0; i < reqDocRels.getItemCount(); i++)
            {
                string subChapter = chapter + (i + 1).ToString();
                Item reqDocRel = reqDocRels.getItemsByXPath("//Item[chapter='" + (subChapter) + "']");
                if (reqDocRel.getItemCount() < 1)
                {
                    break;
                }

                Item requirement = reqDocRel.getRelatedItem();
                //生成标题
                createHeading(requirement, level, subChapter);
                //生成内容
                createContent(requirement);

                generateWordContent(reqDocRels, subChapter + ".", level + 1);
            }
        }

        private void createHeading(Item requirement, int styleId, string chapter)
        {
            //向新文档中添加段落
            XWPFParagraph para = doc.CreateParagraph();
            //设置样式ID
            para.Style = styleId.ToString();

            //书签开始
            var m_p = doc.Document.body.GetPArray(doc.Paragraphs.Count - 1);
            CT_Bookmark m_ctbook1 = new CT_Bookmark();
            var m_bookId = doc.Paragraphs.Count + 1;
            m_ctbook1.id = m_bookId.ToString();//"0";
            m_ctbook1.name = "_Toc" + requirement.getID();//书签名，超链接用
            m_p.Items.Add(m_ctbook1);
            m_p.ItemsElementName.Add(ParagraphItemsChoiceType.bookmarkStart);

            //向该段落中添加文字
            XWPFRun r1 = para.CreateRun();
            r1.SetText(chapter + requirement.getProperty("req_title", ""));

            //书签结束
            CT_Bookmark m_ctbook2 = new CT_Bookmark();
            m_ctbook2.id = m_bookId.ToString();//"0";
            m_p.Items.Add(m_ctbook2);
            m_p.ItemsElementName.Add(ParagraphItemsChoiceType.bookmarkEnd);
        }

        private void createContent(Item requirement)
        {
            string content = requirement.getProperty("content","","en");
            if (string.IsNullOrEmpty(content))
            {
                return;
            }
            XmlDocument contentDom = new XmlDocument();
            contentDom.LoadXml(content);
            foreach (XmlNode field in contentDom.FirstChild.ChildNodes)
            {
                string fieldName = field.Name;
                if (fieldName == "Requirement-Info")
                {
                    continue;
                }

                switch (fieldName)
                {
                    case "Graphic"://图片
                        //向新文档中添加段落
                        XWPFParagraph para = doc.CreateParagraph();
                        generateReqGraphic(para, field);
                        break;
                    case "Label"://标签
                        //向新文档中添加段落
                        XWPFParagraph para1 = doc.CreateParagraph();
                        generateReqLabel(para1, field);
                        break;
                    case "Text"://文本
                        //向新文档中添加段落
                        XWPFParagraph para2 = doc.CreateParagraph();
                        generateReqText(para2, field);
                        break;
                    case "Subtitle"://副标题
                        //向新文档中添加段落
                        XWPFParagraph para3 = doc.CreateParagraph();
                        generateReqSubTitle(para3, field);
                        break;
                    case "Title"://标题
                        //向新文档中添加段落
                        XWPFParagraph para4 = doc.CreateParagraph();
                        generateReqTitle(para4, field);
                        break;
                    case "Entry"://当作是空行
                        //向新文档中添加段落
                        XWPFParagraph para5 = doc.CreateParagraph();
                        generateReqEntry(para5);
                        break;
                    case "List"://列表
                        generateReqList(field, 1);
                        break;
                    case "Row"://行(表格的行)
                        break;
                    case "Table"://表格
                        generateTable(field);
                        //如果下一个控件也是Table的话,必须在他们直接插入一条空行,不然两个Table会合并
                        if (field.NextSibling != null && field.NextSibling.Name == "Table")
                        {
                            doc.CreateParagraph();
                        }
                        TableCount++;
                        break;
                }
            }

        }

        private void generateReqTitle(XWPFParagraph para, XmlNode field)
        {
            foreach (XmlNode runField in field.ChildNodes)
            {
                //向该段落中添加文字
                XWPFRun r1 = para.CreateRun();
                r1.SetText(runField.InnerText);
                //载入基本样式
                r1.FontSize = ReqStyle.Title.FontSize;
                r1.IsBold = ReqStyle.Title.IsBold;
                //加载用户自定义样式
                if (runField.Attributes["under"] != null && runField.Attributes["under"].Value == "true")
                {
                    r1.SetUnderline(UnderlinePatterns.Single);//下划线
                }
                if (runField.Attributes["italic"] != null && runField.Attributes["italic"].Value == "true")
                {
                    r1.IsItalic = true;//斜体
                }
                if (runField.Attributes["strike"] != null && runField.Attributes["strike"].Value == "true")
                {
                    r1.IsStrikeThrough = true;//删除线
                }
                if (runField.Attributes["sub"] != null && runField.Attributes["sub"].Value == "true")
                {
                    r1.Subscript = VerticalAlign.SUBSCRIPT;//下标
                }
                if (runField.Attributes["sup"] != null && runField.Attributes["sup"].Value == "true")
                {
                    r1.Subscript = VerticalAlign.SUPERSCRIPT;//上标
                }
            }
        }

        private void generateReqSubTitle(XWPFParagraph para, XmlNode field)
        {
            //向该段落中添加文字
            XWPFRun r1 = para.CreateRun();
            r1.SetText(field.InnerText);
            //载入基本样式
            r1.FontSize = ReqStyle.Subtitle.FontSize;
            r1.IsBold = ReqStyle.Subtitle.IsBold;
        }

        private void generateReqText(XWPFParagraph para, XmlNode field)
        {
            foreach (XmlNode runField in field.ChildNodes)
            {
                //向该段落中添加文字
                XWPFRun r1 = para.CreateRun();
                r1.SetText(runField.InnerText);
                //载入基本样式
                r1.FontSize = ReqStyle.Text.FontSize;
                r1.IsBold = ReqStyle.Text.IsBold;
                //加载用户自定义样式
                if (runField.Attributes["under"] != null && runField.Attributes["under"].Value == "true")
                {
                    r1.SetUnderline(UnderlinePatterns.Single);//下划线
                }
                if (runField.Attributes["italic"] != null && runField.Attributes["italic"].Value == "true")
                {
                    r1.IsItalic = true;//斜体
                }
                if (runField.Attributes["strike"] != null && runField.Attributes["strike"].Value == "true")
                {
                    r1.IsStrikeThrough = true;//删除线
                }
                if (runField.Attributes["sub"] != null && runField.Attributes["sub"].Value == "true")
                {
                    r1.Subscript = VerticalAlign.SUBSCRIPT;//下标
                }
                if (runField.Attributes["sup"] != null && runField.Attributes["sup"].Value == "true")
                {
                    r1.Subscript = VerticalAlign.SUPERSCRIPT;//上标
                }
                if (runField.Attributes["bold"] != null && runField.Attributes["bold"].Value == "true")
                {
                    r1.IsBold = true;//加粗
                }

                //处理对象类型文本
                if (runField.Attributes["link"] != null && runField.Attributes["link"].Value == "true")
                {
                    XmlAttribute itemtypeAtt = runField.Attributes["itemtype"];
                    XmlAttribute itemidAtt = runField.Attributes["itemid"];
                    XmlAttribute propertyAtt = runField.Attributes["property"];
                    if (itemtypeAtt != null && itemidAtt != null&& propertyAtt!=null)
                    {
                        Item linkItem=innovator.getItemById(itemtypeAtt.Value, itemidAtt.Value);
                        if (linkItem != null&& linkItem.getItemCount()==1)
                        {
                            r1.SetText(linkItem.getProperty(propertyAtt.Value,"undefined"));
                        }
                    }
                }
            }
        }

        private void generateReqLabel(XWPFParagraph para, XmlNode field)
        {
            foreach (XmlNode runField in field.ChildNodes)
            {
                //向该段落中添加文字
                XWPFRun r1 = para.CreateRun();
                r1.SetText(runField.InnerText);
                //载入基本样式
                r1.FontSize = ReqStyle.Label.FontSize;
                r1.IsBold = ReqStyle.Label.IsBold;
                //加载用户自定义样式
                if (runField.Attributes["under"] != null && runField.Attributes["under"].Value == "true")
                {
                    r1.SetUnderline(UnderlinePatterns.Single);//下划线
                }
                if (runField.Attributes["italic"] != null && runField.Attributes["italic"].Value == "true")
                {
                    r1.IsItalic = true;//斜体
                }
                if (runField.Attributes["strike"] != null && runField.Attributes["strike"].Value == "true")
                {
                    r1.IsStrikeThrough = true;//删除线
                }
                if (runField.Attributes["sub"] != null && runField.Attributes["sub"].Value == "true")
                {
                    r1.Subscript = VerticalAlign.SUBSCRIPT;//下标
                }
                if (runField.Attributes["sup"] != null && runField.Attributes["sup"].Value == "true")
                {
                    r1.Subscript = VerticalAlign.SUPERSCRIPT;//上标
                }
            }
        }

        private void generateReqGraphic(XWPFParagraph para, XmlNode field)
        {
            Item Graphic = innovator.getItemById("tp_Image", field.Attributes["imageId"].Value);
            string file_id = Graphic.getProperty("src", "").Replace("vault:///?fileId=", "");
            Item imageFile = innovator.getItemById("File", file_id);
            string filePath = downloadDir+"\\" + file_id + imageFile.getProperty("filename");
            conn.DownloadFile(imageFile, filePath, true);
            if (!File.Exists(filePath))
            {
                return;
            }

            //向该段落中添加图片
            XWPFRun r1 = para.CreateRun();
            using (FileStream fsImg = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                r1.AddPicture(fsImg, (int)PictureType.PNG, "Image", (int)(400.0 * 9525), (int)(300.0 * 9525));
            }

            File.Delete(filePath);
        }

        private void generateReqEntry(XWPFParagraph para)
        {
            //载入基本样式
            para.Alignment = ParagraphAlignment.LEFT;
            para.VerticalAlignment = TextAlignment.TOP;
        }

        private void generateReqList(XmlNode field, int level)
        {
            #region 创建列表样式

            CT_AbstractNum cTAbstractNum = new CT_AbstractNum();
            cTAbstractNum.abstractNumId = abstractNumIdCount.ToString();
            CT_LongHexNumber cTLong = new CT_LongHexNumber();
            cTLong.val = Encoding.UTF8.GetBytes(abstractNumIdCount.ToString());
            cTAbstractNum.nsid = cTLong;

            CT_Lvl cTLvl = new CT_Lvl();
            cTLvl.start.val = "1";
            cTLvl.pPr.AddNewInd().left = (420 * level).ToString();

            string listType = field.Attributes["type"].Value;
            //列表符号
            switch (listType)
            {
                case "numeric":
                    cTLvl.numFmt.val = ST_NumberFormat.@decimal;//数字类型
                    cTLvl.lvlText.val = "%1.";
                    break;
                case "bullet":
                    cTLvl.numFmt.val = ST_NumberFormat.bullet;//圆点类型
                    cTLvl.lvlText.val = "";
                    CT_Fonts cTFont = cTLvl.rPr.AddNewRFonts();
                    cTFont.ascii = "Wingdings";
                    cTFont.hAnsi = "Wingdings";
                    break;
                case "alpha":
                    cTLvl.numFmt.val = ST_NumberFormat.lowerLetter;//字母类型
                    cTLvl.lvlText.val = "%1.";
                    break;
            }

            cTAbstractNum.lvl = new List<CT_Lvl>();
            cTAbstractNum.lvl.Add(cTLvl);

            XWPFAbstractNum abstractNum = new XWPFAbstractNum(cTAbstractNum);
            XWPFNumbering numbering = doc.GetNumbering();
            string abstractNumID = numbering.AddAbstractNum(abstractNum);
            string numId = numbering.AddNum(abstractNumID);

            abstractNumIdCount++;

            #endregion

            foreach (XmlNode listItemFields in field.ChildNodes)
            {
                //只允许一个listItem下有一个图片/文本或多个List
                bool is_haveField = false;
                foreach (XmlNode listItemField in listItemFields.ChildNodes)
                {
                    string fieldName = listItemField.Name;
                    switch (fieldName)
                    {
                        case "Graphic"://图片
                            if (is_haveField) { continue; }
                            //向新文档中添加段落
                            XWPFParagraph para = doc.CreateParagraph();
                            //para.Style = "a8";
                            para.SetNumID(numId);
                            generateReqGraphic(para, listItemField);
                            is_haveField = true;
                            break;
                        case "Text"://文本
                            if (is_haveField) { continue; }
                            //向新文档中添加段落
                            XWPFParagraph para2 = doc.CreateParagraph();
                            //para.Style = "a8";
                            para2.SetNumID(numId);
                            generateReqText(para2, listItemField);
                            is_haveField = true;
                            break;
                        case "List"://列表
                            generateReqList(listItemField, level + 1);
                            break;
                    }
                }
            }
        }

        private void generateTable(XmlNode field)
        {
            string MergeMatrixString = field.Attributes["MergeMatrix"].Value;
            dynamic MergeConfig = JsonConvert.DeserializeObject(MergeMatrixString);
            string[] columnsWidth = field.Attributes["ColWidth"].Value.Split('|');

            int Rows = (int)MergeConfig["length"].Value;
            int Columns = (int)MergeConfig["count"].Value;

            //创建表格 
            XWPFTable table = doc.CreateTable(Rows, Columns);
            CT_Tbl cTTbl = doc.Document.body.GetTblArray(TableCount);
            CT_TblLayoutType cTtblLay = cTTbl.tblPr.AddNewTblLayout();
            cTtblLay.type = ST_TblLayoutType.@fixed;

            int rowIndex = 0;
            int columnIndex = 0;
            bool isFirstField = true;
            //string columnWidth = (8522 / Columns).ToString();

            foreach (XmlNode rowNode in field.ChildNodes)
            {
                columnIndex = 0;
                foreach (XmlNode columnNode in rowNode.ChildNodes)
                {
                    XWPFTableCell cell = table.GetRow(rowIndex).GetCell(columnIndex);
                    CT_Tc cTCell = cell.GetCTTc();
                    CT_TcPr cTTcpr = cTCell.AddNewTcPr();

                    CT_VerticalJc cTVer = cTTcpr.AddNewVAlign();
                    switch (columnNode.Attributes["valign"].Value)
                    {
                        case "top":
                            cTVer.val = ST_VerticalJc.top;
                            break;
                        case "middle":
                            cTVer.val = ST_VerticalJc.center;
                            break;
                        case "bottom":
                            cTVer.val = ST_VerticalJc.bottom;
                            break;
                    }

                    ParagraphAlignment paraAlign = ParagraphAlignment.LEFT;
                    switch (columnNode.Attributes["align"].Value)
                    {
                        case "left":
                            paraAlign = ParagraphAlignment.LEFT;
                            break;
                        case "right":
                            paraAlign = ParagraphAlignment.RIGHT;
                            break;
                        case "center":
                            paraAlign = ParagraphAlignment.CENTER;
                            break;
                        case "justify":
                            paraAlign = ParagraphAlignment.LEFT;
                            break;
                    }

                    CT_TblWidth cTTcw = cTTcpr.AddNewTcW();
                    cTTcw.type = ST_TblWidth.dxa;
                    cTTcw.w = (85.22*int.Parse(columnsWidth[columnIndex])).ToString();

                    isFirstField = true;
                    foreach (XmlNode fileNode in columnNode.ChildNodes)
                    {
                        string fieldName = fileNode.Name;
                        switch (fieldName)
                        {
                            case "Graphic"://图片
                                //向新文档中添加段落
                                XWPFParagraph para;
                                if (isFirstField)
                                {
                                    para = cell.Paragraphs[0];
                                }
                                else
                                {
                                    para = cell.AddParagraph();
                                }
                                para.Alignment = paraAlign;
                                generateReqGraphic(para, fileNode);
                                break;
                            case "Text"://文本
                                //向新文档中添加段落
                                XWPFParagraph para2;
                                if (isFirstField)
                                {
                                    para2 = cell.Paragraphs[0];
                                }
                                else
                                {
                                    para2 = cell.AddParagraph();
                                }
                                para2.Alignment = paraAlign;

                                generateReqText(para2, fileNode);
                                break;
                            case "List"://列表
                                generateReqTableList(fileNode, cell, 0, isFirstField, paraAlign);
                                break;
                        }
                        isFirstField = false;
                    }
                    columnIndex++;
                }
                rowIndex++;
            }

            #region 处理单元格合并

            List<List<string>> rowLists = new List<List<string>>();
            foreach (var rowJobject in MergeConfig)
            {
                if (rowJobject.Name == "length" || rowJobject.Name == "count" || rowJobject.Name == "array")
                {
                    continue;
                }
                rowLists.Add(rowJobject.Value.ToObject<List<string>>());
            }

            for(rowIndex = 0;rowIndex<rowLists.Count;rowIndex++)
            {
                List<string> rowList = rowLists[rowIndex];
                columnIndex = 0;
                int toRowIndex=rowIndex;
                for (columnIndex=0; columnIndex<rowList.Count; columnIndex++)
                {
                    string cellId = rowList[columnIndex];
                    int toColumnIndex = 0;
                    toRowIndex = rowLists.FindLastIndex(delegate (List<string> cIndex)
                    {
                        toColumnIndex = cIndex.FindLastIndex(delegate (string fieldId) {
                            return cellId == fieldId;
                        });

                        return toColumnIndex >= 0;
                    });

                    //合并列单元格
                    if (rowIndex == toRowIndex && columnIndex != toColumnIndex)
                    {
                        mergeCellsHorizontal(table, rowIndex, columnIndex, toColumnIndex);
                    }
                    //合并行单元格
                    if (rowIndex != toRowIndex && columnIndex == toColumnIndex)
                    {
                        mergeCellsVertically(table, columnIndex, rowIndex, toRowIndex);
                    }

                    columnIndex = toColumnIndex;
                }
                rowIndex = toRowIndex;
            }

            #endregion
        }

        // word跨列合并单元格
        private void mergeCellsHorizontal(XWPFTable table, int row, int fromCell, int toCell)
        {
            for (int cellIndex = fromCell; cellIndex <= toCell; cellIndex++)
            {
                XWPFTableCell cell = table.GetRow(row).GetCell(cellIndex);
                if (cellIndex == fromCell)
                {
                    // The first merged cell is set with RESTART merge value    
                    cell.GetCTTc().tcPr.AddNewHMerge().val = ST_Merge.restart;
                }
                else
                {
                    // Cells which join (merge) the first one, are set with CONTINUE    
                    cell.GetCTTc().tcPr.AddNewHMerge().val = ST_Merge.@continue;
                }
            }
        }

        // word跨行并单元格  
        private void mergeCellsVertically(XWPFTable table, int col, int fromRow, int toRow)
        {
            for (int rowIndex = fromRow; rowIndex <= toRow; rowIndex++)
            {
                XWPFTableCell cell = table.GetRow(rowIndex).GetCell(col);
                if (rowIndex == fromRow)
                {
                    // The first merged cell is set with RESTART merge value    
                    cell.GetCTTc().tcPr.AddNewVMerge().val = ST_Merge.restart;
                }
                else
                {
                    // Cells which join (merge) the first one, are set with CONTINUE    
                    cell.GetCTTc().tcPr.AddNewVMerge().val = ST_Merge.@continue;
                }
            }
        }

        private void generateReqTableList(XmlNode field, XWPFTableCell cell, int level, bool isFirstField, ParagraphAlignment paraAlign)
        {

            #region 创建列表样式

            CT_AbstractNum cTAbstractNum = new CT_AbstractNum();
            cTAbstractNum.abstractNumId = abstractNumIdCount.ToString();
            CT_LongHexNumber cTLong = new CT_LongHexNumber();
            cTLong.val = Encoding.UTF8.GetBytes(abstractNumIdCount.ToString());
            cTAbstractNum.nsid = cTLong;

            CT_Lvl cTLvl = new CT_Lvl();
            cTLvl.start.val = "1";
            cTLvl.pPr.AddNewInd().left = (420 * level).ToString();

            string listType = field.Attributes["type"].Value;
            //列表符号
            switch (listType)
            {
                case "numeric":
                    cTLvl.numFmt.val = ST_NumberFormat.@decimal;//数字类型
                    cTLvl.lvlText.val = "%1.";
                    break;
                case "bullet":
                    cTLvl.numFmt.val = ST_NumberFormat.bullet;//圆点类型
                    cTLvl.lvlText.val = "";
                    CT_Fonts cTFont = cTLvl.rPr.AddNewRFonts();
                    cTFont.ascii = "Wingdings";
                    cTFont.hAnsi = "Wingdings";
                    break;
                case "alpha":
                    cTLvl.numFmt.val = ST_NumberFormat.lowerLetter;//字母类型
                    cTLvl.lvlText.val = "%1.";
                    break;
            }

            cTAbstractNum.lvl = new List<CT_Lvl>();
            cTAbstractNum.lvl.Add(cTLvl);

            XWPFAbstractNum abstractNum = new XWPFAbstractNum(cTAbstractNum);
            XWPFNumbering numbering = doc.GetNumbering();
            string abstractNumID = numbering.AddAbstractNum(abstractNum);
            string numId = numbering.AddNum(abstractNumID);

            abstractNumIdCount++;

            #endregion

            foreach (XmlNode listItemFields in field.ChildNodes)
            {
                //只允许一个listItem下有一个图片/文本或多个List
                bool is_haveField = false;
                foreach (XmlNode listItemField in listItemFields.ChildNodes)
                {
                    string fieldName = listItemField.Name;
                    switch (fieldName)
                    {
                        case "Graphic"://图片
                            if (is_haveField) { continue; }
                            //向新文档中添加段落
                            XWPFParagraph para;
                            if (isFirstField)
                            {
                                para = cell.Paragraphs[0];
                            }
                            else
                            {
                                para = cell.AddParagraph();
                            }
                            //para.Style = "a8";
                            para.Alignment = paraAlign;
                            para.SetNumID(numId);
                            generateReqGraphic(para, listItemField);
                            is_haveField = true;
                            break;
                        case "Text"://文本
                            if (is_haveField) { continue; }
                            //向新文档中添加段落
                            XWPFParagraph para2;
                            if (isFirstField)
                            {
                                para2 = cell.Paragraphs[0];
                            }
                            else
                            {
                                para2 = cell.AddParagraph();
                            }
                            //para.Style = "a8";
                            para2.Alignment = paraAlign;
                            para2.SetNumID(numId);
                            generateReqText(para2, listItemField);
                            is_haveField = true;
                            break;
                        case "List"://列表
                            generateReqTableList(listItemField, cell, level + 1, isFirstField, paraAlign);
                            break;
                    }
                    isFirstField = false;
                }
            }
        }

        private string downloadTemplateFile(string templateId)
        {
            string errorMessage = "获取模板文件失败";
            Item docFile = innovator.newItem("Document File", "get");
            docFile.setProperty("source_id", templateId);
            docFile.setAttribute("maxRecords", "1");
            docFile.setAttribute("select", "related_id(id,filename)");
            docFile = docFile.apply();
            if (docFile.isError() || docFile.getItemCount() < 1)
            {
                throw new Exception(errorMessage);
            }

            Item file = docFile.getRelatedItem();
            string filePath = downloadDir+"\\" + file.getID() + file.getProperty("filename");
            conn.DownloadFile(file, filePath, true);
            if (!File.Exists(filePath))
            {
                throw new Exception(errorMessage);
            }
            return filePath;
        }

        private void CreateTOC(TOC tOC)
        {
            //tengz
            XWPFStyles styles = doc.GetStyles();
            string tocStartOrEnd = "start";
            int paraIndex = 0;

            //TOC tOC = new TOC(doc.Document.body.AddNewSdt());
            foreach (XWPFParagraph current in doc.Paragraphs)
            {
                //tengz
                string style = current.Style;
                XWPFStyle xwpf_style = styles.GetStyle(style);
                if (style != null && xwpf_style != null && xwpf_style.Name.StartsWith("heading"))
                //if (style != null && style.StartsWith("Heading"))
                {
                    try
                    {
                        //tengz
                        int level = int.Parse(xwpf_style.Name.Substring("heading".Length));
                        //目录只显示3层
                        if (level > 3)
                        {
                            continue;
                        }

                        //int level = int.Parse(style.Substring("Heading".Length));
                        var mark = doc.Document.body.GetPArray(paraIndex).GetBookmarkStartList();

                        //tengz
                        tOC.AddRow(level, current.Text, 1, mark.ElementAt(0).name, tocStartOrEnd);
                        //tOC.AddRow(level, current.Text, 1, "112723803");

                        tocStartOrEnd = "";
                    }
                    catch (FormatException)
                    {
                    }
                }

                paraIndex++;
            }

            tOC.AddRow(0, "", 0, "", "end");
        }

        private class TOC
        {
            // Token: 0x0600011C RID: 284 RVA: 0x0000449F File Offset: 0x0000269F
            public TOC() : this(new CT_SdtBlock())
            {
            }

            // Token: 0x0600011D RID: 285 RVA: 0x000044AC File Offset: 0x000026AC
            public TOC(CT_SdtBlock block)
            {
                this.block = block;
                CT_SdtPr expr_13 = block.AddNewSdtPr();
                expr_13.AddNewId().val = "4844945";
                expr_13.AddNewDocPartObj().AddNewDocPartGallery().val = "Table of Contents";
                CT_RPr expr_42 = block.AddNewSdtEndPr().AddNewRPr();
                CT_Fonts expr_48 = expr_42.AddNewRFonts();
                expr_48.asciiTheme = ST_Theme.minorHAnsi;
                expr_48.eastAsiaTheme = ST_Theme.minorHAnsi;
                expr_48.hAnsiTheme = ST_Theme.minorHAnsi;
                expr_48.cstheme = ST_Theme.minorBidi;
                expr_42.AddNewB().val = false;
                expr_42.AddNewBCs().val = false;
                expr_42.AddNewColor().val = "auto";
                expr_42.AddNewSz().val = 24uL;
                expr_42.AddNewSzCs().val = 24uL;
                CT_P arg_C1_0 = block.AddNewSdtContent().AddNewP();
                byte[] bytes = Encoding.Unicode.GetBytes("00EF7E24");
                arg_C1_0.rsidR = bytes;
                arg_C1_0.rsidRDefault = bytes;
                //tengz
                arg_C1_0.AddNewPPr().AddNewPStyle().val = "TOC";
                arg_C1_0.AddNewR().AddNewT().Value = "目录";
            }

            // Token: 0x0600011E RID: 286 RVA: 0x000045B1 File Offset: 0x000027B1
            public CT_SdtBlock GetBlock()
            {
                return this.block;
            }

            // Token: 0x0600011F RID: 287 RVA: 0x000045BC File Offset: 0x000027BC
            public void AddRow(int level, string title, int page, string bookmarkRef, string tocStartOrEnd)
            {
                CT_P arg_20_0 = this.block.sdtContent.AddNewP();
                byte[] bytes = Encoding.Unicode.GetBytes("00EF7E24");
                arg_20_0.rsidR = bytes;
                arg_20_0.rsidRDefault = bytes;

                //处理目录无法更新的问题
                if (tocStartOrEnd == "end")
                {
                    CT_R expr_B4 = arg_20_0.AddNewR();
                    expr_B4.AddNewFldChar().fldCharType = ST_FldCharType.end;
                    return;
                }

                CT_PPr expr_34 = arg_20_0.AddNewPPr();
                //tengz
                //expr_34.AddNewPStyle().val = level.ToString()+1;
                //expr_34.AddNewPStyle().val = "toc " + level;
                expr_34.AddNewPStyle().val = "TOC" + level.ToString();
                CT_TabStop expr_5A = expr_34.AddNewTabs().AddNewTab();
                expr_5A.val = ST_TabJc.right;
                expr_5A.leader = ST_TabTlc.dot;
                expr_5A.pos = "8290";
                expr_34.AddNewRPr().AddNewNoProof();

                //处理目录无法更新的问题
                if (tocStartOrEnd == "start")
                {
                    CT_R expr_B5 = arg_20_0.AddNewR();
                    expr_B5.AddNewFldChar().fldCharType = ST_FldCharType.begin;
                    CT_R expr_D1 = arg_20_0.AddNewR();
                    CT_Text expr_E1 = expr_D1.AddNewInstrText();
                    expr_E1.space = "preserve";
                    expr_E1.Value = "TOC \\o \"1 - 3\" \\h \\z \\u";
                    CT_R expr_118 = arg_20_0.AddNewR();
                    expr_118.AddNewFldChar().fldCharType = ST_FldCharType.separate;
                }
                CT_R expr_B6 = arg_20_0.AddNewR();
                expr_B6.AddNewRPr().AddNewNoProof();
                expr_B6.AddNewFldChar().fldCharType = ST_FldCharType.begin;
                CT_R expr_D2 = arg_20_0.AddNewR();
                expr_D2.AddNewRPr().AddNewNoProof();
                CT_Text expr_E2 = expr_D2.AddNewInstrText();
                expr_E2.space = "preserve";
                expr_E2.Value = " PAGEREF " + bookmarkRef + " \\h ";
                arg_20_0.AddNewR().AddNewRPr().AddNewNoProof();
                CT_R expr_119 = arg_20_0.AddNewR();
                expr_119.AddNewRPr().AddNewNoProof();
                expr_119.AddNewFldChar().fldCharType = ST_FldCharType.separate;


                CT_R expr_83 = arg_20_0.AddNewR();
                expr_83.AddNewRPr().AddNewNoProof();
                expr_83.AddNewT().Value = title;
                CT_R expr_9F = arg_20_0.AddNewR();
                expr_9F.AddNewRPr().AddNewNoProof();
                expr_9F.AddNewTab();



                CT_R expr_135 = arg_20_0.AddNewR();
                expr_135.AddNewRPr().AddNewNoProof();
                expr_135.AddNewT().Value = page.ToString();
                CT_R expr_156 = arg_20_0.AddNewR();
                expr_156.AddNewRPr().AddNewNoProof();
                expr_156.AddNewFldChar().fldCharType = ST_FldCharType.end;
            }

            // Token: 0x04000135 RID: 309
            private CT_SdtBlock block;
        }

        private class RequirementStyle
        {
            public RequirementStyle()
            {
                //Aras内样式是CSS格式不知道如何进行格式化,暂时先不处理
                //Item tp_Stylesheet = inn.getItemByKeyedName("tp_Stylesheet", "RE PDF Style Settings");
                //string style_content = tp_Stylesheet.getProperty("style_content");

                Title.FontSize = 18;
                Title.IsBold = true;

                Text.FontSize = 12;
                Text.IsBold = false;

                Label.FontSize = 13;
                Label.IsBold = true;

                Subtitle.FontSize = 16;
                Subtitle.IsBold = false;

                Entry.Align = ParagraphAlignment.LEFT;
                Entry.VAlign = TextAlignment.TOP;
            }

            public FieldStyle Title = new FieldStyle();
            public FieldStyle Text = new FieldStyle();
            public FieldStyle Label = new FieldStyle();
            public FieldStyle Subtitle = new FieldStyle();
            public FieldStyle Entry = new FieldStyle();
        }

        private class FieldStyle
        {
            public int FontSize;
            public bool IsBold;

            public ParagraphAlignment Align;
            public TextAlignment VAlign;
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

        /// <summary>
        /// 去掉文件名中的无效字符,如 \ / : * ? " < > | 
        /// </summary>
        /// <param name="fileName">待处理的文件名</param>
        /// <returns>处理后的文件名</returns>
        private string ReplaceBadCharOfFileName(string fileName)
        {
            string str = fileName;
            str = str.Replace("\\", "");
            str = str.Replace("/", "");
            str = str.Replace(":", "");
            str = str.Replace("*", "");
            str = str.Replace("?", "");
            str = str.Replace("\"", "");
            str = str.Replace("<", "");
            str = str.Replace(">", "");
            str = str.Replace("|", "");
            return str;
        }

        #endregion
    }
}
