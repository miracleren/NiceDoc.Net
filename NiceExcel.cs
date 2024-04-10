using NPOI.SS.Formula.Eval;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.Util;
using NPOI.XSSF.UserModel;
using Org.BouncyCastle.Utilities;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace NiceDoc.Net
{
    public class NiceExcel
    {
        private XSSFWorkbook xlsx;

        /**
         * 根据路径初始化word模板
         *
         * @param path
         */
        public NiceExcel(string path)
        {
            if (!path.EndsWith(".xlsx"))
                System.Console.WriteLine("无效文档后缀，当前只支持xlsx格式Excel文档模板。");

            FileStream inFile;
            try
            {
                inFile = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                xlsx = new XSSFWorkbook(inFile);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
            finally
            {
                if (xlsx == null) xlsx = new XSSFWorkbook();
            }
        }

        /**
         * 往模板填充标签值
         * {{labelName}}
         *
         * @param labels 标签值
         */
        public void pushLabels(Dictionary<string, object> labels)
        {
            //遍历excel所有sheet
            for (int i = 0; i < xlsx.NumberOfSheets; i++)
            {
                ISheet sheet = xlsx.GetSheetAt(i);
                //表格遍历行
                for (int rowNum = 0; rowNum <= sheet.LastRowNum; rowNum++)
                {
                    IRow row = sheet.GetRow(rowNum);
                    if (row == null)
                    {
                        continue;
                    }
                    for (int cellNum = 0; cellNum <= row.LastCellNum; cellNum++)
                    {
                        ICell cell = row.GetCell(cellNum);
                        if (cell == null)
                        {
                            continue;
                        }
                        replaceLabelsInCell(cell, labels);
                    }
                }
            }
        }

        /**
         * 段落填充标签
         *
         * @param cell
         * @param pars
         */
        private void replaceLabelsInCell(ICell cell, Dictionary<string, object> pars)
        {
            string cellValue = cell.ToString();
            if (String.IsNullOrEmpty(cellValue) || cellValue.Contains("col#"))
                return;

            MatchCollection labels = NiceUtils.getMatchingLabels(cellValue);
            foreach (Match m in labels)
            {
                string label = m.Value;

                string[] key = label.Split('#');
                int indexName = key[0].IndexOf("=") + 1 + key[0].IndexOf("&") + 1;
                string keyName = indexName > 0 ? key[0].Substring(0, indexName - 1) : key[0];
                //标签书签
                if (pars.ContainsKey(keyName))
                {
                    //普通文本标签
                    Object val = pars[keyName] == null ? "" : pars[keyName];
                    if (key.Length == 1)
                    {
                        cellValue = cellValue.Replace(NiceUtils.labelFormat(label), val.ToString());
                        cell.SetCellValue(cellValue);
                        continue;
                    }

                    if (key.Length == 2)
                    {
                        //日期类型填充
                        if (key[1].StartsWith("Date:"))
                        {
                            string textVal = val.Equals("") ? val.ToString() : new SimpleDateFormat(key[1].Replace("Date:", "")).Format(val);
                            cellValue = cellValue.Replace(NiceUtils.labelFormat(label), textVal);
                            cell.SetCellValue(cellValue);
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
                                    cellValue = cellValue.Replace(NiceUtils.labelFormat(label), keyVal.Replace(val + ":", ""));
                                    cell.SetCellValue(cellValue);
                                }
                            }
                            continue;
                        }

                        //值判定类型标签
                        string[] boolLabel = key[1].Split(':');
                        string trueVal = boolLabel[0];
                        string falseVal = boolLabel.Length == 1 ? "" : boolLabel[1];
                        if (boolLabel.Length >= 1)
                        {
                            string textVal = "";
                            if (key[0].Contains("="))
                            {
                                textVal = val.ToString().Equals(key[0].Substring(indexName)) ? trueVal : falseVal;
                            }
                            else if (key[0].Contains("&"))
                            {
                                int curVal = Convert.ToInt32(key[0].Substring(indexName));
                                textVal = (Convert.ToInt32(val.ToString()) & curVal) == curVal ? trueVal : falseVal;
                            }
                            else
                            {
                                textVal = val.ToString().Equals("true") ? trueVal : falseVal;
                            }
                            cellValue = cellValue.Replace(NiceUtils.labelFormat(label), textVal);
                            cell.SetCellValue(cellValue);
                        }

                    }
                }
                else if (keyName.Equals("v-if"))
                {
                    logicLabelsInParagraph(cell, pars);
                }
            }

        }


        /**
         * 逻辑语句处理，同一cell内有效
         */
        private void logicLabelsInParagraph(ICell cell, Dictionary<string, object> pars)
        {
            string cellValue = cell.ToString();

            bool isShow = true;
            MatchCollection labels = NiceUtils.getMatchingLabels(cellValue);
            foreach (Match m in labels)
            {
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
                        if (key[0].Equals("v-if"))
                        {
                            if (key[1].Contains("="))
                            {
                                isShow = val.Equals(key[1].Substring(indexName));
                            }
                            else if (key[1].Contains("&"))
                            {
                                int curVal = Convert.ToInt32(key[1].Substring(indexName));
                                isShow = (Convert.ToInt32(val) & curVal) == curVal;
                            }
                            else
                            {
                                isShow = val.Equals("true");
                            }

                            if (isShow == false)
                            {
                                if (cellValue.IndexOf("{{end-if}}") > cellValue.IndexOf(NiceUtils.labelFormat(label)))
                                    cellValue = cellValue.Replace(cellValue.Substring(cellValue.IndexOf(NiceUtils.labelFormat(label)), cellValue.IndexOf("{{end-if}}")), "");
                                else
                                    cellValue = cellValue.Replace(cellValue.Substring(cellValue.IndexOf(NiceUtils.labelFormat(label))), "");
                            }
                            else cellValue = cellValue.Replace(NiceUtils.labelFormat(label), "");

                            cell.SetCellValue(cellValue);
                        }
                    }
                }
                else if (label.Equals("end-if"))
                {
                    cell.SetCellValue(cellValue.Replace(NiceUtils.labelFormat(label), ""));
                }
            }
        }

        /**
         * 填充表格内容到excel
         * {{tableName:colName}}
         *
         * @param tableName
         * @param list
         */
        public void pushTable(string tableName, List<Dictionary<string, object>> list)
        {
            //遍历excel所有sheet
            for (int i = 0; i < xlsx.NumberOfSheets; i++)
            {
                ISheet sheet = xlsx.GetSheetAt(i);
                //表格遍历行
                for (int rowNum = 0; rowNum <= sheet.LastRowNum; rowNum++)
                {
                    IRow row = sheet.GetRow(rowNum);
                    if (row == null)
                    {
                        continue;
                    }
                    for (int cellNum = 0; cellNum <= row.LastCellNum; cellNum++)
                    {
                        ICell cell = row.GetCell(cellNum);
                        if (cell != null && cell.ToString().Contains(tableName + "/col#"))
                        {
                            // 插入数据空白行，数据往后移
                            sheet.ShiftRows(rowNum + 1, sheet.LastRowNum, list.Count - 1);

                            //插入表格数据
                            int addNum = 0;
                            foreach (Dictionary<string, object> rowData in list)
                            {
                                //拷贝当前行
                                IRow setRow = sheet.GetRow(rowNum + addNum);
                                if (list.Count > addNum + 1)
                                {
                                    IRow newRow = sheet.CreateRow(rowNum + addNum + 1);
                                    copyRow(setRow, newRow);
                                }

                                //填充当前行内容数据
                                for (int setCellNum = 0; setCellNum <= row.LastCellNum; setCellNum++)
                                {
                                    ICell setCell = setRow.GetCell(setCellNum);
                                    if (setCell != null)
                                    {
                                        string text = setCell.ToString();
                                        MatchCollection labels = NiceUtils.getMatchingLabels(text);
                                        foreach (Match m in labels)
                                        {
                                            string label = m.Value;
                                            string[] key = label.Split('#');
                                            if (rowData.ContainsKey(key[key.Length - 1]))
                                            {
                                                string val = text.Replace(NiceUtils.labelFormat(label), rowData[key[key.Length - 1]].ToString());
                                                if (NiceUtils.isNumber(rowData[key[key.Length - 1]]))
                                                    setCell.SetCellValue(Convert.ToDouble(val));
                                                else
                                                    setCell.SetCellValue(val);
                                            }
                                        }
                                    }
                                }
                                addNum++;
                            }
                            return;
                        }
                    }
                }
            }
        }

        /**
         * 拷贝行数据
         *
         * @param currentRow
         * @param newRow
         */
        private static void copyRow(IRow currentRow, IRow newRow)
        {
            newRow.Height = currentRow.Height;
            for (int i = 0; i < currentRow.LastCellNum; i++)
            {
                ICell oldCell = currentRow.GetCell(i);
                ICell newCell = newRow.CreateCell(i);
                if (oldCell != null)
                {
                    // 复制样式和值
                    newCell.CellStyle = oldCell.CellStyle;
                    switch (oldCell.CellType)
                    {
                        case CellType.String:
                            newCell.SetCellValue(oldCell.StringCellValue);
                            break;
                        case CellType.Numeric:
                            newCell.SetCellValue(oldCell.NumericCellValue);
                            break;
                        case CellType.Boolean:
                            newCell.SetCellValue(oldCell.BooleanCellValue);
                            break;
                        // ...其他类型
                        default:
                            newCell.SetCellType(oldCell.CellType);
                            break;
                    }
                }
            }
        }

        /**
         * 保存excel文件到目录下
         *
         * @param path
         * @param name
         */
        public void save(string path, string name)
        {
            try
            {
                FileStream outStream = new FileStream(path + name, FileMode.CreateNew);
                xlsx.Write(outStream);
                outStream.Close();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }
        }
    }

}
