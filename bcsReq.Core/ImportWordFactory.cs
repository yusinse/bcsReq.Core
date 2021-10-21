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
    public class ImportWordFactory
    {
        #region "                   宣告區"

        protected Innovator Cinn;
        protected CInnovator.bwGeneric CbwGeneric = new CInnovator.bwGeneric(); //Modify by kenny 2019/04/11
        protected CGeneric.Common CoCommon;//Modify by kenny 2019/04/11
        private string CstrErrMessage = "";

        protected Innovator innovator { get; private set; }
        protected IServerConnection conn { get; private set; }
        protected string LangCode { get; private set; }

        //临时文件路径
        private string downloadDir;

        private XWPFDocument doc;
        private XWPFNumbering xwpfNum;
        private XWPFStyles styles;
        private int TableCount = 0;
        private Item XmlSchema;

        #endregion

        #region "                   進入點"

        public ImportWordFactory()
        {
            Cinn = new Innovator(null);
        }


        public ImportWordFactory(Innovator getInnovator, CallContext CCO)
        {
            //System.Diagnostics.Debugger.Break();

            innovator = getInnovator;
            conn = innovator.getConnection();
            downloadDir = CCO.Server.MapPath("../Client/Solutions/RE/TemporaryFiles/")+innovator.getNewID();
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

        public Item ImportFromWord(string fileId, string itemtypeName)
        {
            //System.Diagnostics.Debugger.Break();
            if (CheckLicense() == false)
            {
                return Cinn.newError(CstrErrMessage);
            }

            try
            {
                Item result = innovator.newError("对象类名称错误");
                switch (itemtypeName.ToLower())
                {
                    case "re_requirement_document":
                        result = ImportReqDocFromWord(fileId);
                        break;
                    case "re_requirement":
                        result = ImportReqFromWord(fileId);
                        break;
                }

                return result;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private Item ImportReqDocFromWord(string fileId)
        {
            try
            {
                XmlSchema = innovator.getItemById("tp_XmlSchema", "8DF7037346A64816B8BBD8700AFCFE15");

                //创建临时目录
                Directory.CreateDirectory(downloadDir);

                //下载Word文件到临时目录
                Item file = innovator.getItemById("File", fileId);
                string filePath = downloadDir+"//" + file.getProperty("filename");
                conn.DownloadFile(file, filePath, true);
                if (!File.Exists(filePath))
                {
                    throw new Exception("下载Word文件到临时目录失败!");
                }

                // 读取Word文档
                doc = new XWPFDocument(new FileStream(filePath, FileMode.Open));

                //读取文档列表样式
                xwpfNum = doc.GetNumbering();

                //读取文档样式
                styles = doc.GetStyles();

                Item reqs = innovator.newItem();
                Item req = null;
                XmlDocument RequirementDom = null;
                int rootChapter = 0;
                for (int i = 0; i < doc.BodyElements.Count; i++)
                {
                    IBodyElement xwpfItem = doc.BodyElements[i];
                    switch (xwpfItem.ElementType)
                    {
                        case BodyElementType.PARAGRAPH:
                            XWPFParagraph para = (XWPFParagraph)xwpfItem;
                            bool isHeading = false;
                            if (para.StyleID != null)
                            {
                                var paraStyle = styles.GetStyle(para.StyleID);
                                if (paraStyle.Name.Contains("heading"))
                                {
                                    if (!checkHeadingIsEmpty(para))
                                    {
                                        isHeading = true;

                                        if (req != null)
                                        {
                                            //req.setProperty("content", RequirementDom.OuterXml, LangCode);
                                            req.setProperty("content", RequirementDom.OuterXml, "en");
                                            reqs.appendItem(req);
                                        }

                                        req = createNewRequirementItem(paraStyle, req, ref rootChapter);
                                        RequirementDom = generateRequirementDom(req);
                                    }
                                }
                            }
                            if (req == null && !isHeading)
                            {
                                continue;
                            }
                            generateParagraph(RequirementDom, RequirementDom.DocumentElement, para, isHeading, ref i, doc.BodyElements, req);
                            break;
                        case BodyElementType.TABLE:
                            if (req == null)
                            {
                                TableCount++;
                                continue;
                            }

                            XWPFTable table = (XWPFTable)xwpfItem;
                            generateTable(table, RequirementDom, req);
                            TableCount++;
                            break;
                    }
                }

                //文档内没有标题的情况下req就会为null
                if (req != null)
                {
                    //req.setProperty("content", RequirementDom.OuterXml, LangCode);
                    req.setProperty("content", RequirementDom.OuterXml, "en");
                    reqs.appendItem(req);
                }


                if (reqs.getItemCount() > 1)
                {
                    reqs.removeItem(reqs.getItemByIndex(0));
                }
                else
                {
                    reqs = innovator.newError("文件内无有效数据!");
                }

                //删除临时文件与临时文件夹
                File.Delete(filePath);
                Directory.Delete(downloadDir);

                return reqs;
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

        private Item ImportReqFromWord(string fileId)
        {
            try
            {
                //创建临时目录
                Directory.CreateDirectory(downloadDir);

                //下载Word文件到临时目录
                Item file = innovator.getItemById("File", fileId);
                string filePath = downloadDir+"//" + file.getProperty("filename");
                conn.DownloadFile(file, filePath, true);
                if (!File.Exists(filePath))
                {
                    throw new Exception("下载Word文件到临时目录失败!");
                }

                // 读取Word文档
                doc = new XWPFDocument(new FileStream(filePath, FileMode.Open));

                //读取文档列表样式
                xwpfNum = doc.GetNumbering();

                //读取文档样式
                styles = doc.GetStyles();

                XmlDocument RequirementDom = new XmlDocument();
                RequirementDom.AppendChild(RequirementDom.CreateElement("Result"));
                for (int i = 0; i < doc.BodyElements.Count; i++)
                {
                    IBodyElement xwpfItem = doc.BodyElements[i];
                    switch (xwpfItem.ElementType)
                    {
                        case BodyElementType.PARAGRAPH:
                            XWPFParagraph para = (XWPFParagraph)xwpfItem;
                            generateParagraph(RequirementDom, RequirementDom.DocumentElement, para, ref i, doc.BodyElements, null);
                            break;
                        case BodyElementType.TABLE:
                            XWPFTable table = (XWPFTable)xwpfItem;
                            generateTable(table, RequirementDom, null);
                            TableCount++;
                            break;
                    }
                }

                //删除临时文件与临时文件夹
                File.Delete(filePath);
                Directory.Delete(downloadDir);

                return innovator.newResult(RequirementDom.OuterXml);
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

        private void generateParagraph(XmlDocument RequirementDom, XmlElement parentNode, XWPFParagraph para, bool isHeading, ref int index, IList<IBodyElement> bodyElements, Item req)
        {
            XmlNamespaceManager xmlNameSpace = new XmlNamespaceManager(RequirementDom.NameTable);
            xmlNameSpace.AddNamespace("re", "http://www.aras.com/REStandard");
            xmlNameSpace.AddNamespace("aras", "http://aras.com/ArasTechDoc");

            if (isHeading)
            {
                string title = "";
                foreach (var run in para.Runs)
                {
                    title += run.Text;
                }
                title = removeOrder(title);
                RequirementDom.SelectSingleNode("//re:Requirement-Title/aras:emph", xmlNameSpace).InnerText = title;
                req.setProperty("req_title", title);
            }
            else
            {
                string numId = para.GetNumID();
                if (numId != null)
                {
                    int numLvl = 0;

                    XmlElement listNode = RequirementDom.CreateElement("List");
                    SetArasIdAttribute(RequirementDom, listNode, innovator.getNewID());
                    parentNode.AppendChild(listNode);
                    XmlElement listItemNode = RequirementDom.CreateElement("List-Item");
                    SetArasIdAttribute(RequirementDom, listItemNode, innovator.getNewID());
                    listNode.AppendChild(listItemNode);
                    XmlElement textNode = RequirementDom.CreateElement("Text");
                    SetArasIdAttribute(RequirementDom, textNode, innovator.getNewID());
                    listItemNode.AppendChild(textNode);
                    generateParagraphRuns(RequirementDom, para, textNode, listItemNode, req);

                    numLvl = int.Parse(para.GetNumIlvl());
                    //设置列表符号
                    listNode.SetAttribute("type", getListFormat(numId, numLvl));

                    //继续检查下一元素,看看是否还是列表
                    int preNumlvl = numLvl;
                    XmlElement preListItemNode = listItemNode;
                    for (int ii = index + 1; ii < bodyElements.Count; ii++)
                    {
                        IBodyElement xwpfItem = bodyElements[ii];
                        //当前元素不是段落则代表不是接续的列表,结束
                        if (xwpfItem.ElementType != BodyElementType.PARAGRAPH)
                        {
                            break;
                        }
                        XWPFParagraph nextPara = (XWPFParagraph)xwpfItem;
                        string nextNumId = nextPara.GetNumID();
                        //当前段落是标题或不是列表,结束
                        if ((nextPara.StyleID != null && styles.GetStyle(nextPara.StyleID).Name.Contains("heading")) || nextNumId == null)
                        {
                            break;
                        }
                        //当前段落列表ID与主列表ID不一样,代表是新列表,结束
                        if (numId != nextNumId)
                        {
                            break;
                        }
                        int nextNumLvl = int.Parse(nextPara.GetNumIlvl());
                        //主列表层阶大于当前段落列表层阶,当作是新列表,结束
                        if (numLvl > nextNumLvl)
                        {
                            break;
                        }
                        if (preNumlvl == nextNumLvl)
                        {
                            XmlElement nextListItemNode = RequirementDom.CreateElement("List-Item");
                            SetArasIdAttribute(RequirementDom, nextListItemNode, innovator.getNewID());
                            preListItemNode.SelectSingleNode("..").AppendChild(nextListItemNode);
                            XmlElement nextTextNode = RequirementDom.CreateElement("Text");
                            SetArasIdAttribute(RequirementDom, nextTextNode, innovator.getNewID());
                            nextListItemNode.AppendChild(nextTextNode);
                            generateParagraphRuns(RequirementDom, nextPara, nextTextNode, nextListItemNode, req);

                            preListItemNode = nextListItemNode;
                        }
                        else if (preNumlvl < nextNumLvl)
                        {

                            XmlElement nextListNode = RequirementDom.CreateElement("List");
                            SetArasIdAttribute(RequirementDom, nextListNode, innovator.getNewID());
                            //设置列表符号
                            nextListNode.SetAttribute("type", getListFormat(nextNumId, nextNumLvl));
                            preListItemNode.AppendChild(nextListNode);

                            XmlElement nextListItemNode = RequirementDom.CreateElement("List-Item");
                            SetArasIdAttribute(RequirementDom, nextListItemNode, innovator.getNewID());
                            nextListNode.AppendChild(nextListItemNode);

                            XmlElement nextTextNode = RequirementDom.CreateElement("Text");
                            SetArasIdAttribute(RequirementDom, nextTextNode, innovator.getNewID());
                            nextListItemNode.AppendChild(nextTextNode);
                            generateParagraphRuns(RequirementDom, nextPara, nextTextNode, nextListItemNode, req);

                            preListItemNode = nextListItemNode;
                        }
                        else
                        {
                            string parentXPath = "";
                            for (int j = 0; j < preNumlvl - nextNumLvl; j++)
                            {
                                parentXPath += "../../";
                            }
                            parentXPath += "..";

                            XmlElement nextListItemNode = RequirementDom.CreateElement("List-Item");
                            SetArasIdAttribute(RequirementDom, nextListItemNode, innovator.getNewID());
                            preListItemNode.SelectSingleNode(parentXPath).AppendChild(nextListItemNode);
                            XmlElement nextTextNode = RequirementDom.CreateElement("Text");
                            SetArasIdAttribute(RequirementDom, nextTextNode, innovator.getNewID());
                            nextListItemNode.AppendChild(nextTextNode);
                            generateParagraphRuns(RequirementDom, nextPara, nextTextNode, nextListItemNode, req);

                            preListItemNode = nextListItemNode;
                        }
                        preNumlvl = nextNumLvl;
                        index++;
                    }
                }
                else
                {
                    XmlElement textNode = RequirementDom.CreateElement("Text");
                    SetArasIdAttribute(RequirementDom, textNode, innovator.getNewID());
                    parentNode.AppendChild(textNode);
                    generateParagraphRuns(RequirementDom, para, textNode, parentNode, req);
                }
            }
        }

        private void generateParagraph(XmlDocument RequirementDom, XmlElement parentNode, XWPFParagraph para, ref int index, IList<IBodyElement> bodyElements, Item req)
        {
            string numId = para.GetNumID();
            //string styleId = para.StyleID;

            if (numId != null)
            {
                int numLvl = 0;

                XmlElement listNode = RequirementDom.CreateElement("List");
                SetArasIdAttribute(RequirementDom, listNode, innovator.getNewID());
                parentNode.AppendChild(listNode);
                XmlElement listItemNode = RequirementDom.CreateElement("List-Item");
                SetArasIdAttribute(RequirementDom, listItemNode, innovator.getNewID());
                listNode.AppendChild(listItemNode);
                XmlElement textNode = RequirementDom.CreateElement("Text");
                SetArasIdAttribute(RequirementDom, textNode, innovator.getNewID());
                listItemNode.AppendChild(textNode);
                generateParagraphRuns(RequirementDom, para, textNode, listItemNode, req);

                numLvl = int.Parse(para.GetNumIlvl());
                //设置列表符号
                listNode.SetAttribute("type", getListFormat(numId, numLvl));

                //继续检查下一元素,看看是否还是列表
                int preNumlvl = numLvl;
                XmlElement preListItemNode = listItemNode;
                for (int ii = index + 1; ii < bodyElements.Count; ii++)
                {
                    IBodyElement xwpfItem = bodyElements[ii];
                    //当前元素不是段落则代表不是接续的列表,结束
                    if (xwpfItem.ElementType != BodyElementType.PARAGRAPH)
                    {
                        break;
                    }
                    XWPFParagraph nextPara = (XWPFParagraph)xwpfItem;
                    string nextNumId = nextPara.GetNumID();
                    //当前段落是标题或不是列表,结束
                    if ((nextPara.StyleID != null && styles.GetStyle(nextPara.StyleID).Name.Contains("heading")) || nextNumId == null)
                    {
                        break;
                    }
                    //当前段落列表ID与主列表ID不一样,代表是新列表,结束
                    if (numId != nextNumId)
                    {
                        break;
                    }
                    int nextNumLvl = int.Parse(nextPara.GetNumIlvl());
                    //主列表层阶大于当前段落列表层阶,当作是新列表,结束
                    if (numLvl > nextNumLvl)
                    {
                        break;
                    }
                    if (preNumlvl == nextNumLvl)
                    {
                        XmlElement nextListItemNode = RequirementDom.CreateElement("List-Item");
                        SetArasIdAttribute(RequirementDom, nextListItemNode, innovator.getNewID());
                        preListItemNode.SelectSingleNode("..").AppendChild(nextListItemNode);
                        XmlElement nextTextNode = RequirementDom.CreateElement("Text");
                        SetArasIdAttribute(RequirementDom, nextTextNode, innovator.getNewID());
                        nextListItemNode.AppendChild(nextTextNode);
                        generateParagraphRuns(RequirementDom, nextPara, nextTextNode, nextListItemNode, req);

                        preListItemNode = nextListItemNode;
                    }
                    else if (preNumlvl < nextNumLvl)
                    {

                        XmlElement nextListNode = RequirementDom.CreateElement("List");
                        SetArasIdAttribute(RequirementDom, nextListNode, innovator.getNewID());
                        //设置列表符号
                        nextListNode.SetAttribute("type", getListFormat(nextNumId, nextNumLvl));
                        preListItemNode.AppendChild(nextListNode);

                        XmlElement nextListItemNode = RequirementDom.CreateElement("List-Item");
                        SetArasIdAttribute(RequirementDom, nextListItemNode, innovator.getNewID());
                        nextListNode.AppendChild(nextListItemNode);

                        XmlElement nextTextNode = RequirementDom.CreateElement("Text");
                        SetArasIdAttribute(RequirementDom, nextTextNode, innovator.getNewID());
                        nextListItemNode.AppendChild(nextTextNode);
                        generateParagraphRuns(RequirementDom, nextPara, nextTextNode, nextListItemNode, req);

                        preListItemNode = nextListItemNode;
                    }
                    else
                    {
                        string parentXPath = "";
                        for (int j = 0; j < preNumlvl - nextNumLvl; j++)
                        {
                            parentXPath += "../../";
                        }
                        parentXPath += "..";

                        XmlElement nextListItemNode = RequirementDom.CreateElement("List-Item");
                        SetArasIdAttribute(RequirementDom, nextListItemNode, innovator.getNewID());
                        preListItemNode.SelectSingleNode(parentXPath).AppendChild(nextListItemNode);
                        XmlElement nextTextNode = RequirementDom.CreateElement("Text");
                        SetArasIdAttribute(RequirementDom, nextTextNode, innovator.getNewID());
                        nextListItemNode.AppendChild(nextTextNode);
                        generateParagraphRuns(RequirementDom, nextPara, nextTextNode, nextListItemNode, req);

                        preListItemNode = nextListItemNode;
                    }
                    preNumlvl = nextNumLvl;
                    index++;
                }
            }
            else
            {
                XmlElement textNode = RequirementDom.CreateElement("Text");
                SetArasIdAttribute(RequirementDom, textNode, innovator.getNewID());
                parentNode.AppendChild(textNode);
                generateParagraphRuns(RequirementDom, para, textNode, parentNode, req);
            }

        }

        private void generateParagraphRuns(XmlDocument RequirementDom, XWPFParagraph para, XmlElement textNode, XmlElement parentNode, Item req)
        {
            foreach (var run in para.Runs)
            {
                //图片
                List<XWPFPicture> pics = run.GetEmbeddedPictures(); 
                if (pics.Count > 0)
                {
                    foreach (XWPFPicture pic in pics)
                    {
                        XmlElement picNode = RequirementDom.CreateElement("Graphic");
                        SetArasIdAttribute(RequirementDom, picNode, innovator.getNewID());
                        string ref_id = innovator.getNewID();

                        XWPFPictureData picData = pic.GetPictureData();
                        string imagePath = Base64StringToImage(picData);
                        Item imageFile = innovator.newItem("File", "add");
                        imageFile.setProperty("filename", picData.FileName);
                        imageFile.attachPhysicalFile(imagePath);
                        imageFile = imageFile.apply();
                        File.Delete(imagePath);
                        if (imageFile.isError())
                        {
                            throw new Exception("上传图片至Aras失败");
                        }
                        Item Graphic = innovator.newItem("tp_Image", "add");
                        Graphic.setNewID();
                        Graphic.setProperty("item_number", Graphic.getID());
                        Graphic.setProperty("name", picData.FileName);
                        Graphic.setProperty("src", "vault:///?fileId=" + imageFile.getID());
                        Graphic = Graphic.apply();
                        if (Graphic.isError())
                        {
                            throw new Exception("建立Graphic失败!");
                        }

                        picNode.SetAttribute("imageId", Graphic.getID());
                        parentNode.AppendChild(picNode);

                        //导入需求时不需要进行如下设置
                        if (req != null)
                        {
                            picNode.SetAttribute("ref-id", ref_id);
                            Item tp_ref = req.createRelationship("re_ImageReference", "add");
                            tp_ref.setProperty("reference_id", ref_id);
                            tp_ref.setRelatedItem(Graphic);
                        }
                    }
                }
                else
                {
                    //文字
                    XmlElement emphNode = RequirementDom.CreateElement("aras:emph");
                    emphNode.SetAttribute("xmlns", "");
                    emphNode.InnerText = run.Text;
                    //下划线
                    if (run.Underline != UnderlinePatterns.None)
                    {
                        emphNode.SetAttribute("under", "true");
                    }
                    //斜体
                    if (run.IsItalic)
                    {
                        emphNode.SetAttribute("italic", "true");
                    }
                    //删除线
                    if (run.IsStrikeThrough || run.IsDoubleStrikeThrough)
                    {
                        emphNode.SetAttribute("strike", "true");
                    }
                    //下标
                    if (run.Subscript == VerticalAlign.SUBSCRIPT)
                    {
                        emphNode.SetAttribute("sub", "true");
                    }
                    //上标
                    if (run.Subscript == VerticalAlign.SUPERSCRIPT)
                    {
                        emphNode.SetAttribute("sup", "true");
                    }
                    //加粗
                    if (run.IsBold)
                    {
                        emphNode.SetAttribute("bold", "true");
                    }
                    textNode.AppendChild(emphNode);

                }
            }
            //if (string.IsNullOrWhiteSpace(textNode.InnerText) && havePic)
            if (string.IsNullOrWhiteSpace(textNode.InnerText))
            {
                parentNode.RemoveChild(textNode);
            }
        }

        private void generateTable(XWPFTable table, XmlDocument RequirementDom, Item req)
        {
            #region 生成Aras表格数据
            Dictionary<String, dynamic> MergeMatrix = new Dictionary<String, dynamic>();
            MergeMatrix.Add("array", new List<string>());
            string rowId;

            XmlElement tableNode = RequirementDom.CreateElement("Table");
            SetArasIdAttribute(RequirementDom, tableNode, innovator.getNewID());
            
            RequirementDom.DocumentElement.AppendChild(tableNode);
            
            for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++)
            {
                XmlElement rowNode = RequirementDom.CreateElement("Row");
                rowId = innovator.getNewID();
                SetArasIdAttribute(RequirementDom, rowNode, rowId);
                tableNode.AppendChild(rowNode);
                MergeMatrix.ElementAt(0).Value.Add(rowId);

                List<string> cellGUIDs = new List<string>();
                XWPFTableRow tableRow = table.Rows[rowIndex];
                List<XWPFTableCell> cells = tableRow.GetTableCells();
                for (int cellIndex = 0; cellIndex < cells.Count; cellIndex++)
                {
                    XmlElement cellNode = RequirementDom.CreateElement("Entry");
                    SetArasIdAttribute(RequirementDom, cellNode, innovator.getNewID());
                    cellNode.SetAttribute("valign", "");
                    cellNode.SetAttribute("align", "");
                    rowNode.AppendChild(cellNode);

                    XWPFTableCell cell = cells[cellIndex];
                    CT_Tc cTCell = cell.GetCTTc();
                    if (cTCell.tcPr.vMerge != null && cTCell.tcPr.vMerge.val == ST_Merge.@continue)
                    {
                        cellGUIDs.Add(MergeMatrix.ElementAt(rowIndex).Value[cellIndex]);
                    }
                    else
                    {
                        cellGUIDs.Add(Guid.NewGuid().ToString());
                    }

                    if (cTCell.tcPr.vAlign != null)
                    {
                        switch (cTCell.tcPr.vAlign.val)
                        {
                            case ST_VerticalJc.top:
                                cellNode.SetAttribute("valign", "top");
                                break;
                            case ST_VerticalJc.center:
                                cellNode.SetAttribute("valign", "middle");
                                break;
                            case ST_VerticalJc.bottom:
                                cellNode.SetAttribute("valign", "bottom");
                                break;
                            default:
                                cellNode.SetAttribute("valign", "top");
                                break;
                        }
                    }
                    else
                    {
                        cellNode.SetAttribute("valign", "top");
                    }

                    for (int paraIndex = 0; paraIndex < cell.Paragraphs.Count; paraIndex++)
                    {
                        XWPFParagraph para = cell.Paragraphs[paraIndex];
                        switch (para.Alignment)
                        {
                            case ParagraphAlignment.LEFT:
                                cellNode.SetAttribute("align", "left");
                                break;
                            case ParagraphAlignment.RIGHT:
                                cellNode.SetAttribute("align", "right");
                                break;
                            case ParagraphAlignment.CENTER:
                                cellNode.SetAttribute("align", "center");
                                break;
                            default:
                                cellNode.SetAttribute("align", "left");
                                break;
                        }

                        generateParagraph(RequirementDom, cellNode, para, false, ref paraIndex, cell.BodyElements, req);
                    }

                    #region 处理合并列情况
                    //如果表格列有合并那Cells就只会有一个合并的列,但Aras中需要有多个列,所以需要手动添加空列进去
                    if (cTCell.tcPr.gridSpan != null)
                    {
                        int gridSpan = int.Parse(cTCell.tcPr.gridSpan.val);
                        for (int span = 1; span < gridSpan; span++)
                        {
                            cellGUIDs.Add(cellGUIDs[cellGUIDs.Count-1]);
                            XmlElement mergeCellNode = RequirementDom.CreateElement("Entry");
                            SetArasIdAttribute(RequirementDom, mergeCellNode, innovator.getNewID());
                            mergeCellNode.SetAttribute("valign", "top");
                            mergeCellNode.SetAttribute("align", "left");
                            rowNode.AppendChild(mergeCellNode);
                        }
                    }
                    #endregion
                }

                MergeMatrix.Add(rowId, cellGUIDs);
            }

            #endregion

            #region 设置Aras表格列宽(百分比)

            CT_Tbl cTTbl = doc.Document.body.GetTblArray(TableCount);
            int maxCellCount = cTTbl.tblGrid.gridCol.Count;
            List<int> cellWidths = new List<int>();
            int tableWidth = 0;
            //计算出表格总宽
            for (int i = 0; i < maxCellCount; i++)
            {
                tableWidth += int.Parse(cTTbl.tblGrid.gridCol[i].w.ToString());
            }

            //计算每列宽度百分比
            for (int i = 0; i < maxCellCount; i++)
            {
                cellWidths.Add((int)Math.Round((Double)cTTbl.tblGrid.gridCol[i].w * 100 / tableWidth));
            }
            tableNode.SetAttribute("ColWidth", string.Join("|", cellWidths.ToArray()));

            #endregion

            #region 设置Aras表格单元格合并数据
            //总行数
            MergeMatrix.Add("length", table.Rows.Count);
            //总列数
            MergeMatrix.Add("count", maxCellCount);

            string json = JsonConvert.SerializeObject(MergeMatrix);
            tableNode.SetAttribute("MergeMatrix", json);

            #endregion
        }

        //判断标题是否为空(或空格)
        private bool checkHeadingIsEmpty(XWPFParagraph para)
        {
            string title = "";
            foreach (var run in para.Runs)
            {
                title += run.Text;
            }
            return string.IsNullOrWhiteSpace(title);
        }

        //删除标题前的序号(举例: 1.1引言-->引言)
        private string removeOrder(string title)
        {
            char[] titleChars = title.ToCharArray();
            int startIndex = 0;
            for (int i = 0; i < titleChars.Length; i++)
            {
                char titleChar = titleChars[i];
                if (!char.IsNumber(titleChar) && titleChar != '.')
                {
                    startIndex = i;
                    break;
                }
            }
            return title.Substring(startIndex);
        }

        //创建RequirementItem
        private Item createNewRequirementItem(XWPFStyle paraStyle,Item req,ref int rootChapter)
        {
            int level = int.Parse(paraStyle.Name.Replace("heading ", ""));
            string chapter = generateChapter(req, level, ref rootChapter);
            req = innovator.newItem("re_Requirement");
            req.setProperty("chapter", chapter);
            req.setNewID();

            req.setProperty("classification", "Requirement");
            req.setProperty("req_priority", "Low");
            req.setProperty("req_risk", "Low");
            req.setProperty("managed_by_id",innovator.getUserAliases());
            req.setProperty("req_complexity", "Low");
            req.setPropertyItem("req_document_type", XmlSchema);

            return req;
        }

        //生成需求结构编号
        private string generateChapter(Item preReq, int level, ref int rootChapter)
        {
            string chapter = "";
            if (level == 1)
            {
                rootChapter++;
                chapter = rootChapter.ToString();
            }
            else
            {
                //当前标题层阶不是1且前序标题为空,结构错误
                if (preReq == null)
                {
                    throw new Exception("标题结构错误,必须从标题1开始");
                }
                else
                {
                    chapter = preReq.getProperty("chapter");
                    string[] chapterArray = chapter.Split('.');
                    if (level > chapterArray.Length)
                    {
                        //当前标题层阶大于前序标题层阶1阶以上,跳级了!结构错误
                        if ((level - chapterArray.Length) > 1)
                        {
                            throw new Exception("标题结构错误,不能跳级!");
                        }
                        chapter += ".1";
                    }
                    else
                    {
                        chapter = "";
                        for (int i = 0; i < level - 1; i++)
                        {
                            chapter += chapterArray[i] + ".";
                        }
                        chapter += (int.Parse(chapterArray[level - 1]) + 1).ToString();
                    }
                }
            }
            return chapter;
        }

        //生成需求内容XML
        private XmlDocument generateRequirementDom(Item req)
        {
            XmlDocument RequirementDom = new XmlDocument();
            string xml = "<Requirement aras:id='" + Guid.NewGuid().ToString("N") + "' reqId='" + req.getID() + "' " +
                        "xmlns='http://www.aras.com/REStandard' xmlns:aras='http://aras.com/ArasTechDoc' >" +
                            "<Requirement-Info aras:id='" + Guid.NewGuid().ToString("N") + "' >" +
                                "<Requirement-Chapter aras:id='" + Guid.NewGuid().ToString("N") + "' >" +
                                    "<aras:emph xmlns='' ></aras:emph>" +
                                "</Requirement-Chapter>" +
                                "<Requirement-Title aras:id='" + Guid.NewGuid().ToString("N") + "' >" +
                                    "<aras:emph></aras:emph>" +
                                "</Requirement-Title>" +
                                "<Requirement-Number aras:id='" + Guid.NewGuid().ToString("N") + "' >" +
                                    "<aras:emph></aras:emph>" +
                                "</Requirement-Number>" +
                            "</Requirement-Info> " +
                        "</Requirement>";
            RequirementDom.LoadXml(xml);

            return RequirementDom;
        }

        //设置节点aras:id属性
        private void SetArasIdAttribute(XmlDocument RequirementDom, XmlElement node, string id)
        {
            XmlAttribute attribute = RequirementDom.CreateAttribute("aras", "id", "http://www.aras.com/REStandard");
            attribute.InnerText = id;
            node.Attributes.Append(attribute);
        }

        //Base64转图片
        private string Base64StringToImage(XWPFPictureData picData)
        {
            try
            {
                string imagePath = downloadDir+"\\" + picData.FileName;
                MemoryStream ms = new MemoryStream(picData.Data);
                System.Drawing.Bitmap bmp = new System.Drawing.Bitmap(ms);

                //图片后缀格式
                string suffix = picData.SuggestFileExtension();
                var suffixName = suffix == "png"
                        ? System.Drawing.Imaging.ImageFormat.Png
                        : suffix == "jpg" || suffix == "jpeg"
                            ? System.Drawing.Imaging.ImageFormat.Jpeg
                            : suffix == "bmp"
                                ? System.Drawing.Imaging.ImageFormat.Bmp
                                : suffix == "gif"
                                    ? System.Drawing.Imaging.ImageFormat.Gif
                                    : System.Drawing.Imaging.ImageFormat.Jpeg;
                bmp.Save(imagePath, suffixName);
                ms.Close();
                if (!File.Exists(imagePath))
                {
                    throw new Exception("生成图片失败");
                }
                return imagePath;
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        //获取列表符号
        private string getListFormat(string numId, int numLvl)
        {
            try
            {
                string absId = xwpfNum.GetNum(numId).GetCTNum().abstractNumId.val;
                ST_NumberFormat numFmt = xwpfNum.GetAbstractNum(absId).GetAbstractNum().GetLvlArray(numLvl).numFmt.val;
                switch (numFmt)
                {
                    case ST_NumberFormat.@decimal:
                        return "numeric";
                    case ST_NumberFormat.lowerLetter:
                    case ST_NumberFormat.upperLetter:
                        return "alpha";
                    case ST_NumberFormat.bullet:
                        return "bullet";
                }
                return "bullet";
            }
            catch
            {
                throw new Exception("读取列表样式失败!");
            }
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
