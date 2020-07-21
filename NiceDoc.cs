using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Words.Saving;
using Aspose.Words.Tables;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Range = Aspose.Words.Range;

namespace aspword_cli
{
    public class NiceDoc
    {

        /**
    * 20200720 基于aspose模板生成word，
    * by miracleren
    */

        private static String ASPOSE_VERSION = "18.7.0";
        Document doc;

        /**
         * 初始化模板
         *
         * @param tempPath
         */
        public NiceDoc(String tempPath)
        {
            try
            {
                doc = new Document(tempPath);
                Console.WriteLine("create docx successully");
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
        }

        /**
         * 标签数据替换
         *
         * @param values map值列表
         */
        public void setLabel(Dictionary<string, object> values)
        {
            Range range = doc.Range;
            string wordText = range.Text;
            MatchCollection pars = matcher(wordText);
            foreach (Match par in pars)
            {
                //Console.WriteLine(pars.group().toString());
                String con = par.ToString();
                String[] cons = con.Split(':');
                //纯内容标签替换
                try
                {
                    if (cons.Length == 1)
                    {
                        String labVal = StringOf(values[con]);
                        rangeReplace(con, labVal);
                    }
                    else
                    {
                        if (cons.Length == 3)
                        {
                            //类型标签
                            String typeName = cons[0];
                            String typePar = cons[1];
                            String typeVal = cons[2];
                            if ("SC" == typeName)
                            {
                                //单选
                                if (StringOf(values[typePar]) == typeVal)
                                    rangeReplace(con, "√");
                                else
                                    rangeReplace(con, "□");
                            }
                            else if ("MC" == typeName)
                            {
                                //多选
                                //String value = StringOf(values.get(typePar));
                                int parval = values[typePar] == null ? 0 : Convert.ToInt16(values[typePar]);
                                int val = Convert.ToInt16(typeVal);
                                if ((parval & val) == val)
                                    rangeReplace(con, "√");
                                else
                                    rangeReplace(con, "□");
                            }
                        }
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine(con + "::" + e);
                }

            }

            //标签更新完成，处理表达式
            setSyntax(values);
        }

        /**
         * 表格循环数据填充
         *
         * @param list
         * @param tableName
         */
        public void setTable(List<Dictionary<string, object>> list, string tableName)
        {
            NodeCollection bookTables = doc.GetChildNodes(NodeType.Table, true);
            foreach (Object table in bookTables)
            {
                Table tb = (Table)table;
                //判断是否循环列表
                String rowFistText = tb.Rows[0].GetText();
                String tableConfig = getFirstParName(rowFistText);
                if (tableConfig != "")
                {
                    //第一行为表格配置信息
                    String[] cons = tableConfig.Split(':');
                    if (cons[0] != "TABLE" && cons[1] != tableName)
                        break;
                }
                else
                    break;

                //查找配置循环列
                int i = 0, tempIndex = -1;
                Row tempRow = null;
                foreach (Row trow in ((Table)table).Rows)
                {
                    if (tempRow != null)
                        break;
                    foreach (Cell tcell in trow.Cells)
                    {
                        if (getFirstParName(tcell.GetText()).Contains("COL"))
                        {
                            tempRow = trow;
                            tempIndex = i;
                            break;
                        }
                    }
                    i++;
                }
                if (tempRow == null)
                    return;

                //克隆行，并赋值
                foreach (Dictionary<string, object> rowData in list)
                {
                    Row newRow = (Row)tempRow.Clone(true);
                    foreach (Cell newRowCell in newRow.Cells)
                    {
                        String cellPars = getFirstParName(newRowCell.Range.Text);
                        if (cellPars != "")
                        {
                            String[] pars = cellPars.Split(':');
                            if (pars[0] == "COL")
                            {
                                rangeReplace(newRowCell.Range, cellPars, StringOf(rowData[pars[1]]));
                            }
                        }
                    }
                    ((Table)table).AppendChild(newRow);
                }

                //清除配置行
                ((Table)table).RemoveChild(((Table)table).Rows[0]);
                ((Table)table).RemoveChild(tempRow);
            }
        }

        /**
         * 表达式判断
         * <p>
         * 目前支持
         * {{V-IF:par}}{{END:par}}  显示隐藏数据,等号目前支持 ==，！=
         */
        public void setSyntax(Dictionary<string, object> values)
        {
            Range range = doc.Range;
            String wordText = range.Text;
            MatchCollection pars = matcher(wordText);
            foreach (Match par in pars)
            {
                String con = par.ToString();
                //if显示隐藏表达式
                if (con.Contains("V-IF:"))
                {
                    String[] cons = con.Split(':');
                    String syn = cons[1];
                    if (syn.Contains("=="))
                    {
                        String[] tem = syn.Replace("==", "@").Split('@');
                        if (StringOf(values[tem[0]]) == tem[1].ToString())
                        {
                            rangeReplace(con, "");
                            rangeReplace("END:" + tem[0], "");
                        }
                        else
                        {
                            Regex pattern = new Regex("(?=\\{\\{" + con + "\\}\\})(.+?)(?<=\\{\\{END:" + tem[0] + "\\}\\})");
                            rangeReplace(pattern, "");

                        }
                    }
                    else if (syn.Contains("!="))
                    {
                        String[] tem = syn.Replace("!=", "@").Split('@');
                        if (StringOf(values[tem[0]]) != tem[1].ToString())
                        {
                            rangeReplace(con, "");
                            rangeReplace("END:" + tem[0], "");
                        }
                        else
                        {
                            Regex pattern = new Regex("(?=\\{\\{" + con + "\\}\\})(.+?)(?<=\\{\\{END:" + tem[0] + "\\}\\})");
                            rangeReplace(pattern, "");

                        }
                    }
                    else
                    {
                        if (StringOf(values[con]) == "true")
                        {
                            rangeReplace(con, "");
                            rangeReplace("END:" + cons, "");
                        }
                        else
                        {
                            Regex pattern = new Regex("(?=\\{\\{" + con + "\\}\\})(.+?)(?<=\\{\\{END:" + cons + "\\}\\})");
                            rangeReplace(pattern, "");

                        }
                    }
                    Console.WriteLine("不支持当前表达式：" + syn);
                }
            }
        }

        /**
         * 实体类转map
         *
         * @param object
         * @return
         */
        public static Dictionary<string, object> entityToMap(object obj)
        {
            Dictionary<string, object> map = new Dictionary<string, object>();
            System.Reflection.PropertyInfo[] properties = obj.GetType().GetProperties(System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.Public);
            foreach (System.Reflection.PropertyInfo item in properties)
            {
                try
                {
                    string name = item.Name;
                    object value = item.GetValue(obj, null);
                    map.Add(name, value);
                }
                catch (Exception e)
                {
                    Console.WriteLine("实体类转换：" + e);
                }
            }
            return map;
        }

        /**
         * 文本替换
         *
         * @param oldStr
         * @param newStr
         */
        private void rangeReplace(String oldStr, String newStr)
        {
            Range range = doc.Range;
            try
            {
                range.Replace("{{" + oldStr + "}}", newStr, new FindReplaceOptions());
            }
            catch (Exception e)
            {
                Console.WriteLine(oldStr + ">>>>>>" + e);
            }
        }

        /**
         * 文本替换
         *
         * @param pattern
         * @param newStr
         */
        private void rangeReplace(Regex pattern, String newStr)
        {
            Range range = doc.Range;
            try
            {
                range.Replace(pattern, newStr, new FindReplaceOptions());
            }
            catch (Exception e)
            {
                Console.WriteLine("pattern >>>>>>" + e);
            }
        }

        /**
         * 文本替换
         *
         * @param range
         * @param oldStr
         * @param newStr
         */
        private void rangeReplace(Range range, String oldStr, String newStr)
        {
            try
            {
                range.Replace("{{" + oldStr + "}}", newStr, new FindReplaceOptions());
                //range.replace("{{" + oldStr + "}}", newStr,true,false);
            }
            catch (Exception e)
            {
                Console.WriteLine(oldStr + ">>>>>>" + e);
            }
        }


        /**
         * {{par}} 参数查找正则
         *
         * @param str 查找串
         * @return 返结果
         */
        private MatchCollection matcher(String str)
        {
            Regex pattern = new Regex("(?<=\\{\\{)(.+?)(?=\\}\\})");
            MatchCollection matcher = pattern.Matches(str);
            return matcher;
        }

        /**
         * 获取数据里第一个标签名称
         *
         * @param str
         * @return
         */
        private String getFirstParName(String str)
        {
            Regex pattern = new Regex("(?<=\\{\\{)(.+?)(?=\\}\\})");
            Match matcher = pattern.Match(str);
            if (matcher.Success)
                return matcher.ToString();
            else
                return "";
        }


        /**
         * 空字符转占位空格
         */
        private String StringOf(Object val)
        {
            return val == null ? "        " : val.ToString();
        }

        public bool save(String ptch)
        {
            try
            {
                doc.Save(ptch);
                return true;
            }
            catch (Exception e)
            {
                Console.WriteLine("保存失败：" + e);
                return false;
            }
        }

        public bool saveOnlyComments(String ptch)
        {
            try
            {
                doc.Protect(ProtectionType.AllowOnlyComments, "teamoneit");
                doc.Save(ptch);
                return true;
            }
            catch (Exception e)
            {
                Console.WriteLine("保存失败：" + e);
                return false;
            }
        }

        public bool savePdf(String ptch)
        {
            try
            {
                PdfSaveOptions op = new PdfSaveOptions();
                op.SaveFormat = SaveFormat.Pdf;
                doc.Save(ptch, op);
                return true;
            }
            catch (Exception e)
            {
                Console.WriteLine("保存失败：" + e);
                return false;
            }
        }

        public MemoryStream saveStream()
        {
            MemoryStream ms = null;
            try
            {
                doc.Save(ms, new OoxmlSaveOptions(SaveFormat.Doc));
            }
            catch (Exception e)
            {
                Console.WriteLine("saveStream 保存失败：" + e);
            }
            return ms;
        }

        protected void finalize()
        {
            doc.Remove();
        }

    }
}
