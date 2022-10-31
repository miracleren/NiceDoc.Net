
using NPOI.OpenXmlFormats.Wordprocessing;
using NPOI.XWPF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;

/**
 * 基于模板快速生成word文档
 * 目前只支持docx文件
 * <p>
 * by miracleren@gmail.com
 */

namespace NiceDoc.Net
{
    public class NiceDoc
    {
        //private HWPFDocument doc;
        private XWPFDocument docx;
        private int status = 0;
        private List<XWPFTable> allTables = new List<XWPFTable>();

        /**
         * 根据路径初始化word模板
         *
         * @param path
         */
        public NiceDoc(string path)
        {
            if (!path.EndsWith(".docx"))
                System.Console.WriteLine("无效文档后缀，当前只支持docx格式word文档模板。");

            FileStream inFile;
            try
            {
                inFile = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                docx = new XWPFDocument(inFile);

                //遍历段落生加载表格列表
                allTables.AddRange(new List<XWPFTable>(docx.Tables));
                pushLabels(new Dictionary<string, object>());
                status = 1;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
            finally
            {
                if (docx == null)
                    docx = new XWPFDocument();
            }
        }


        /**
         * 往模板填充标签值
         * {{labelName}}
         *
         * @param labels 标签值
         * @return
         */
        public void pushLabels(Dictionary<string, object> labels)
        {
            //遍历普通段落内容对像，填充标签值
            List<XWPFParagraph> paragraphs = new List<XWPFParagraph>(docx.Paragraphs);
            replaceLabelsInParagraphs(paragraphs, labels);

            //遍历表格内容，并填充标签值
            List<XWPFTable> tables = status == 0 ? new List<XWPFTable>(docx.Tables) : allTables;
            foreach (XWPFTable table in tables)
            {
                //表格行
                List<XWPFTableRow> rows = table.Rows;
                foreach (XWPFTableRow row in rows)
                {
                    //表格单元格
                    List<XWPFTableCell> cells = row.GetTableCells();
                    foreach (XWPFTableCell cell in cells)
                    {
                        //表格段落
                        List<XWPFParagraph> cellParagraphs = new List<XWPFParagraph>(cell.Paragraphs);
                        replaceLabelsInParagraphs(cellParagraphs, labels);
                    }
                }
            }

            //页眉标签值填充
            List<XWPFHeader> headers = new List<XWPFHeader>(docx.HeaderList);
            foreach (XWPFHeader header in headers)
            {
                List<XWPFParagraph> headerParagraphs = new List<XWPFParagraph>(header.Paragraphs);
                replaceLabelsInParagraphs(headerParagraphs, labels);
            }

            //页脚填充
            List<XWPFFooter> footers = new List<XWPFFooter>(docx.FooterList);
            foreach (XWPFFooter footer in footers)
            {
                List<XWPFParagraph> footerParagraphs = new List<XWPFParagraph>(footer.Paragraphs);
                replaceLabelsInParagraphs(footerParagraphs, labels);
            }
        }

        /**
         * 往模板填充标签值实体类
         *
         * @param entity
         */
        public void pushLabels(Object entity)
        {
            pushLabels(NiceUtils.entityToDictionary(entity));
        }


        /**
         * 填充表格内容到文档
         * {{tableName:colName}}
         *
         * @param tableName
         * @param list
         */
        public void pushTable(string tableName, List<Dictionary<string, object>> list)
        {
            //List<XWPFTable> tables = new List<XWPFTable>(docx.Tables);
            List<XWPFTable> tables = allTables;
            int tableIndex = 0;
            foreach (XWPFTable table in tables)
            {
                bool isFind = false;
                XWPFTableRow baseRow = null;

                List<XWPFTableRow> rows = table.Rows;
                int rowCount = rows.Count;
                for (int i = 0; i < rowCount; i++)
                {
                    List<XWPFTableCell> cells = rows[i].GetTableCells();
                    foreach (XWPFTableCell cell in cells)
                    {
                        List<XWPFParagraph> cellParagraphs = new List<XWPFParagraph>(cell.Paragraphs);
                        foreach (XWPFParagraph cellParagraph in cellParagraphs)
                        {
                            //查找表格标识名称
                            if (!isFind)
                            {
                                if (cellParagraph.Text.Contains(NiceUtils.labelFormat("table#" + tableName)))
                                {
                                    isFind = true;
                                }
                                else
                                {
                                    isFind = false;
                                    break;
                                }
                            }

                            //记录开始数据行
                            if (cellParagraph.Text.Contains("{{col#"))
                            {
                                baseRow = rows[i];
                                break;
                            }
                        }
                        if (!isFind)
                            break;
                    }
                    if (!isFind)
                        break;

                    //已知数据行，开始填充数据
                    if (baseRow != null)
                    {
                        //int addRowIndex = 1;
                        //foreach (Dictionary<string, object> listRow in list)
                        //{
                        //    CT_Tbl m_CTTbl = docx.Document.body.GetTblArray()[1];
                        //    CT_Row ctRow = table.getCTTbl().insertNewTr(i + addRowIndex);

                        //    //table.(i + addRowIndex-1);
                        //    //CT_Row ctRow = table.GetRow(1).GetCTRow();
                        //    //XWPFTableRow newRow = new XWPFTableRow(ctRow, table);
                        //    //copyRowAndPushLabels(newRow, baseRow, listRow);

                        //    CT_Row targetRow = table.CreateRow().GetCTRow();
                        //    targetRow.trPr = baseRow.GetCTRow().trPr;
                        //    targetRow.rsidR = baseRow.GetCTRow().rsidR;
                        //    targetRow.rsidTr = baseRow.GetCTRow().rsidTr;
                        //    XWPFTableRow newRow = new XWPFTableRow(targetRow, table);

                        //    table.Rows.Add(newRow);
                        //    addRowIndex++;
                        //}

                        //baseRow = null;
                        //table.RemoveRow(i);
                        int addRowIndex = 1;
                        //XWPFTableRow row = table.CreateRow();
                        foreach (Dictionary<string, object> listRow in list)
                        {
                            //CT_Row ctRow2 = docx.Document.body.GetTblArray()[tableIndex].InsertNewTr(i + addRowIndex - 1);
                            CT_Row ctRow = table.GetCTTbl().InsertNewTr(i + addRowIndex - 1);
                            XWPFTableRow newRow = new XWPFTableRow(ctRow, table);
                            copyRowAndPushLabels(newRow, baseRow, listRow);
                            //table.addRow(newRow, i + addRowIndex);
                            addRowIndex++;
                        }
                        //docx.Document.body.GetTblArray()[tableIndex].RemoveTr(i + addRowIndex - 1);
                        table.GetCTTbl().RemoveTr(i + addRowIndex - 1);
                        baseRow = null;

                    }
                }
                //删除table标识行
                if (isFind)
                    table.RemoveRow(0);

                tableIndex++;
            }
        }

        /**
         * 拷贝行，并填充相关值
         *
         * @param newRow
         * @param baseRow
         * @param params
         */
        private void copyRowAndPushLabels(XWPFTableRow newRow, XWPFTableRow baseRow, Dictionary<string, object> pars)
        {
            newRow.GetCTRow().trPr = baseRow.GetCTRow().trPr;

            foreach (XWPFTableCell cell in baseRow.GetTableCells())
            {
                XWPFTableCell newCell = newRow.AddNewTableCell();
                newCell.GetCTTc().tcPr = cell.GetCTTc().tcPr;
                bool isFirst = true;
                //newCell.setParagraph(cell.Paragraphs.get(0));
                foreach (XWPFParagraph paragraph in new List<XWPFParagraph>(cell.Paragraphs))
                {
                    XWPFParagraph newParagraph = isFirst ? newCell.Paragraphs[0] : newCell.AddParagraph();
                    isFirst = false;
                    //newParagraph.GetCTP().pPr = paragraph.GetCTP().pPr;
                    foreach (XWPFRun run in paragraph.Runs)
                    {
                        XWPFRun newRun = newParagraph.CreateRun();
                        newRun.GetCTR().rPr = run.GetCTR().rPr;

                        string text = run.GetText(0);
                        if (text == null)
                            continue;
                        else
                            newRun.SetText(text);

                        MatchCollection labels = NiceUtils.getMatchingLabels(text);
                        foreach (Match m in labels)
                        {
                            string label = m.Value;
                            string[] key = label.Split('#');
                            if (pars.ContainsKey(key[key.Length - 1]))
                            {
                                newRun.SetText(text.Replace(NiceUtils.labelFormat(label), pars[key[key.Length - 1]].ToString()), 0);
                            }
                        }
                    }
                }

            }
        }

        /**
         * 段落列表填充标签
         *
         * @param paragraphs
         * @param params
         */
        private void replaceLabelsInParagraphs(List<XWPFParagraph> paragraphs, Dictionary<string, object> pars)
        {
            for (int i = 0; i < paragraphs.Count; i++)
            {
                XWPFParagraph paragraph = paragraphs[i];

                //获取doc表格，包括子表格，docx.getTables()无法获取子表格
                if (status == 0 )
                {
                    foreach(XWPFTable table in new List<XWPFTable>(paragraph.Body.Tables))
                    {
                        if(!allTables.Contains(table))
                        {
                            allTables.Add(table);
                        }
                    }
                    return;
                }

                string text = paragraph.Text;
                if (text == null || text == "" || !text.Contains("{{"))
                    continue;
                else if (text.Contains("{{v-"))
                    logicLabelsInParagraph(paragraphs, i, pars);
                replaceLabelsInParagraph(paragraph, pars);
            }

        }

        /**
         * 清空标签被分割的其它文本
         *
         * @param runs
         */
        private void removeRun(List<XWPFRun> runs)
        {
            //runs.RemoveAt(runs.Count - 1);
            //foreach (XWPFRun run in runs)
            //{
            //    run.SetText("", 0);
            //}
            for (int i = 0; i < runs.Count - 1; i++)
            {
                runs[i].SetText("", 0);
            }
        }

        /**
        * 逻辑语句处理
        */
        private void logicLabelsInParagraph(List<XWPFParagraph> paragraphs, int index, Dictionary<string, object> pars)
        {
            string nowText = "";
            int runCount = 0;
            List<XWPFRun> labelRuns = new List<XWPFRun>();
            bool isShow = true;

            for (int i = index; i < paragraphs.Count; i++)
            {
                XWPFParagraph paragraph = paragraphs[i];
                List<XWPFRun> runs = new List<XWPFRun>(paragraph.Runs);

                foreach (XWPFRun run in runs)
                {
                    if (run.GetText(0) != null && (run.GetText(0).Contains("{{") || runCount > 0))
                    {
                        nowText += run.GetText(0);
                        runCount++;
                        labelRuns.Add(run);

                        MatchCollection labels = NiceUtils.getMatchingLabels(nowText);
                        int labelFindCount = 0;
                        foreach (Match m in labels)
                        {
                            labelFindCount++;
                            string label = m.Value;
                            string[] key = label.Split('#');

                            if (key.Length == 2)
                            {
                                int indexName = key[1].IndexOf("=") + 1 + key[1].IndexOf("&") + 1;
                                string keyName = indexName > 0 ? key[1].Substring(0, indexName - 1) : key[1];
                                if (pars.ContainsKey(keyName))
                                {
                                    string val = pars[keyName] == null ? "" : pars[keyName].ToString();
                                    //条件判断语句
                                    if (key[0] == "v-if")
                                    {
                                        if (key[1].Contains("="))
                                        {
                                            isShow = (val == key[1].Substring(indexName));
                                        }
                                        else if (key[1].Contains("&"))
                                        {
                                            int curVal = Convert.ToInt32(key[1].Substring(indexName));
                                            isShow = (Convert.ToInt32(val) & curVal) == curVal;
                                        }
                                        else
                                        {
                                            isShow = (val == "true");
                                        }

                                        if (isShow == false)
                                        {
                                            if (nowText.IndexOf("{{end-if}}") > nowText.IndexOf(NiceUtils.labelFormat(label)))
                                                nowText = nowText.Replace(nowText.Substring(nowText.IndexOf(NiceUtils.labelFormat(label)), nowText.IndexOf("{{end-if}}")), "");
                                            else
                                                nowText = nowText.Replace(nowText.Substring(nowText.IndexOf(NiceUtils.labelFormat(label))), "");
                                        }
                                        else
                                            nowText = nowText.Replace(NiceUtils.labelFormat(label), "");

                                        run.SetText(nowText.Replace(NiceUtils.labelFormat(label), ""), 0);
                                        removeRun(labelRuns);
                                    }
                                }
                            }
                            else if (label == "end-if")
                            {
                                run.SetText(nowText.Replace(NiceUtils.labelFormat(label), ""), 0);
                                removeRun(labelRuns);
                                isShow = true;
                            }


                        }
                        if (labelFindCount > 0)
                        {
                            nowText = "";
                            runCount = 0;
                            labelRuns = new List<XWPFRun>();
                        }
                    }

                    if (isShow != true)
                    {
                        run.SetText("", 0);
                    }

                }
            }
        }


        /**
         * 段落填充标签
         *
         * @param paragraph
         * @param params
         */
        private void replaceLabelsInParagraph(XWPFParagraph paragraph, Dictionary<string, object> pars)
        {
            //遍历文本对象，查找标识标签
            List<XWPFRun> runs = new List<XWPFRun>(paragraph.Runs);
            string nowText = "";
            int runCount = 0;
            List<XWPFRun> labelRuns = new List<XWPFRun>();

            //常规标签
            foreach (XWPFRun run in runs)
            {
                //防止文本对象标签被分割
                if (run.GetText(0) != null && (run.GetText(0).Contains("{{") || runCount > 0))
                {
                    nowText += run.GetText(0);
                    runCount++;
                    labelRuns.Add(run);

                    //System.out.println(nowText);
                    MatchCollection labels = NiceUtils.getMatchingLabels(nowText);
                    int labelFindCount = 0;
                    foreach (Match m in labels)
                    {
                        labelFindCount++;
                        string label = m.Value;

                        string[] key = label.Split('#');
                        int indexName = key[0].IndexOf("=") + 1 + key[0].IndexOf("&") + 1;
                        string keyName = indexName > 0 ? key[0].Substring(0, indexName - 1) : key[0];
                        //标签书签
                        if (pars.ContainsKey(keyName))
                        {
                            object val = pars[keyName] == null ? "" : pars[keyName];
                            //普通文本标签
                            if (key.Length == 1)
                            {
                                nowText = nowText.Replace(NiceUtils.labelFormat(label), val.ToString());
                                run.SetText(nowText, 0);
                                continue;
                            }

                            if (key.Length == 2)
                            {
                                //日期类型填充
                                if (key[1].StartsWith("Date:"))
                                {
                                    string textVal = val.ToString() == "" ? val.ToString() : Convert.ToDateTime(val).ToString(key[1].Replace("Date:", ""));
                                    nowText = nowText.Replace(NiceUtils.labelFormat(label), textVal);
                                    run.SetText(nowText, 0);
                                    continue;
                                }

                                //枚举数组标签
                                if (key[1].StartsWith("[") && key[1].EndsWith("]"))
                                {
                                    string group = key[1].Substring(1, key[1].Length - 1);
                                    foreach (string keyVal in group.Split(','))
                                    {
                                        if (keyVal.IndexOf(val + ":") == 0)
                                        {
                                            nowText = nowText.Replace(NiceUtils.labelFormat(label), keyVal.Replace(val + ":", ""));
                                            run.SetText(nowText, 0);
                                            removeRun(labelRuns);
                                        }
                                    }
                                    continue;
                                }

                                //值判定类型标签
                                string[] boolVal = key[1].Split(':');
                                string trueVal = boolVal[0];
                                string falseVal = boolVal.Length == 1 ? "" : boolVal[1];
                                if (boolVal.Length >= 1)
                                {
                                    string textVal = "";
                                    if (key[0].Contains("="))
                                    {
                                        textVal = (val.ToString() == key[0].Substring(indexName) ? trueVal : falseVal);
                                    }
                                    else if (key[0].Contains("&"))
                                    {
                                        int curVal = Convert.ToInt32(key[0].Substring(indexName));
                                        textVal = (Convert.ToInt32(val) & curVal) == curVal ? trueVal : falseVal;
                                    }
                                    else
                                    {
                                        textVal = val.ToString() == "true" ? trueVal : falseVal;
                                    }
                                    nowText = nowText.Replace(NiceUtils.labelFormat(label), textVal);
                                    run.SetText(nowText, 0);
                                    removeRun(labelRuns);
                                    continue;
                                }

                            }
                        }
                    }

                    if (labelFindCount > 0)
                    {
                        nowText = "";
                        runCount = 0;
                        labelRuns = new List<XWPFRun>();
                    }
                }

            }
        }

        /**
         * 清除条件语句产生的空段落
         */
        public void removeNullParagraphs()
        {
            List<XWPFParagraph> paragraphs = new List<XWPFParagraph>(docx.Paragraphs);
            List<IBodyElement> listBe = new List<IBodyElement>(docx.BodyElements);

            for (int i = 0; i < listBe.Count; i++)
            {
                if (listBe[i].ElementType == BodyElementType.PARAGRAPH)
                {
                    if (paragraphs[docx.GetParagraphPos(i)].Text.Contains("R"))
                    {
                        docx.RemoveBodyElement(i);
                        i--;
                        continue;
                    }
                }

            }
        }

        /**
         * 段落条件标签处理
         *
         * @param paragraph
         * @param params
         */
        private void syntaxLabelsInParagraph(XWPFParagraph paragraph, Dictionary<string, object> pars)
        {

        }

        /**
         * 设置word只读
         */
        public void setReadOnly()
        {
            docx.EnforceFillingFormsProtection();
        }

        /**
         * 保存word文件到目录下
         *
         * @param path
         * @param name
         */
        public void save(string path, string name)
        {
            try
            {
                FileStream outStream = new FileStream(path + name, FileMode.CreateNew);
                docx.Write(outStream);
                outStream.Close();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }
        }
    }
}
