using System;
using System.IO;
using System.Threading;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using System.Collections;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;

namespace ReportDesign
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        Report report = new Report();
        Dictionary<Company, int> companyDic = new Dictionary<Company, int>();
        Dictionary<string, int> authorDic = new Dictionary<string, int>();
        private void button1_Click(object sender, EventArgs e)
        {
            if (!comboBox1.Text.Equals("按照总体统计") && !comboBox1.Text.Equals("按照年份统计"))
            {
                AssistUtils.showMessage("请选中下拉列表中的内容");
                return;
            }
            companyDic.Clear();
            authorDic.Clear();
            //打开文件
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.InitialDirectory = System.Windows.Forms.Application.StartupPath;
            openFileDialog.Filter = "word文档(2003)|*.doc|word(2007/1010)|*.docx";
            openFileDialog.RestoreDirectory = true;
            openFileDialog.FilterIndex = 1;
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    report.CreateNewDocument(openFileDialog.FileName);
                }
                catch
                {
                    AssistUtils.showMessage("不能打开");
                    return;
                }
                //读取一行
                int index = 1;
                string year = "";
                while (true)
                {
                    try
                    {
                        string paragraph = report.Document.Paragraphs[index++].Range.Text.Trim();
                        //对一行操作。提取作者，（段落第一个and之前，或者从开始到第一个逗号结束提取作者）查看集合是否存在该作者，存在则对应数量＋1；提取机构需要按照每一年份提取（分号之前逗号之后，或者最后一个逗号之后）,存在则+1
                        string paragraphAuthor = paragraph;
                        if (paragraph.Contains("年"))
                        {
                            year = paragraph.Substring(0, paragraph.IndexOf("年"));
                            continue;
                        }
                        try
                        {
                            //分两种情况，一种是有分号，则取数组之后对每个字符串取到最后一个逗号之前，再将每个数组按照逗号拆分，即可得到每个作者，此时的作者需要删掉and
                            //另外是没有分号，取最后一个逗号之前的字符串，再将每个数组按照逗号拆分，即可得到每个作者，此时的作者需要删掉and
                            string[] authorArr = {""};
                            //string 
                            int idx1 = paragraphAuthor.IndexOf(";");
                            #region 没有分号
                            if (idx1 == -1)
                            {
                                idx1 = paragraphAuthor.LastIndexOf(",");
                                if (idx1 != -1)
                                {
                                    string t = paragraphAuthor.Substring(0, idx1);
                                    paragraphAuthor = t;
                                }
                                authorArr = Regex.Split(paragraphAuthor, ",");
                                foreach (string s in authorArr)
                                {
                                    if (s.Contains("and "))
                                    {
                                        if (s.IndexOf("and ") < 3)
                                        {
                                            string author = s.Substring(s.IndexOf("and ") + 3, s.Length - s.IndexOf("and ") - 3);
                                            int idx = author.IndexOf("、");
                                            if (idx != -1)
                                            {
                                                string strT = author.Substring(idx + 1, author.Length - idx - 1);
                                                author = strT.Trim();
                                            }
                                            if (authorDic.ContainsKey(author))
                                            {
                                                authorDic[author]++;
                                            }
                                            else
                                            {
                                                authorDic.Add(author, 1);
                                            }
                                        }
                                        else
                                        {
                                            string[] strArrTemp = Regex.Split(s, "and ");
                                            for (int i = 0; i < strArrTemp.Length; i++)
                                            {
                                                string author = strArrTemp[i].Trim();
                                                int idx = author.IndexOf("、");
                                                if (idx != -1)
                                                {
                                                    string strT = author.Substring(idx + 1, author.Length - idx - 1);
                                                    author = strT.Trim();
                                                }
                                                if (authorDic.ContainsKey(author))
                                                {
                                                    authorDic[author]++;
                                                }
                                                else
                                                {
                                                    authorDic.Add(author, 1);
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                            #endregion
                            #region 有分号
                            else
                            {
                                string[] auArr = Regex.Split(paragraphAuthor, ";");
                                foreach (string str in auArr)
                                {
                                    idx1 = str.LastIndexOf(",");
                                    if (idx1 != -1)
                                    {
                                        string t = str.Substring(0, idx1);
                                        paragraphAuthor = t;
                                    }
                                    authorArr = Regex.Split(paragraphAuthor, ",");
                                    foreach (string s in authorArr)
                                    {
                                        if (s.Contains("and "))
                                        {
                                            if (s.IndexOf("and ") < 3)
                                            {
                                                string author = s.Substring(s.IndexOf("and ") + 3, s.Length - s.IndexOf("and ") - 3);
                                                int idx = author.IndexOf("、");
                                                if (idx != -1)
                                                {
                                                    string strT = author.Substring(idx + 1, author.Length - idx - 1);
                                                    author = strT.Trim();
                                                }
                                                if (authorDic.ContainsKey(author))
                                                {
                                                    authorDic[author]++;
                                                }
                                                else
                                                {
                                                    authorDic.Add(author, 1);
                                                }
                                            }
                                            else
                                            {
                                                string[] strArrTemp = Regex.Split(s, "and ");
                                                for (int i = 0; i < strArrTemp.Length; i++)
                                                {
                                                    string author = strArrTemp[i];
                                                    int idx = author.IndexOf("、");
                                                    if (idx != -1)
                                                    {
                                                        string strT = author.Substring(idx + 1, author.Length - idx - 1);
                                                        author = strT.Trim();
                                                    }
                                                    if (authorDic.ContainsKey(author))
                                                    {
                                                        authorDic[author]++;
                                                    }
                                                    else
                                                    {
                                                        authorDic.Add(author, 1);
                                                    }
                                                }
                                            }
                                        }
                                        else
                                        {
                                            string sAuthor = s;
                                            int idxT = sAuthor.IndexOf("、");
                                            if (idxT != -1)
                                            {
                                                string strT = sAuthor.Substring(idxT + 1, sAuthor.Length - idxT - 1);
                                                sAuthor = strT.Trim();
                                            }
                                            if (authorDic.ContainsKey(sAuthor))
                                            {
                                                authorDic[sAuthor]++;
                                            }
                                            else
                                            {
                                                authorDic.Add(sAuthor, 1);
                                            }
                                        }
                                    }
                                }
                            }
                            #endregion

                            /*
                            int idx1 = paragraph.IndexOf(",");
                            int idx2 = paragraph.IndexOf(" and");
                            int index_ = 0;
                            
                            if (idx1 != -1 && idx2 != -1)
                                author = paragraph.Substring(0, (idx1 > idx2) ? idx2 : idx1).Trim();
                            else if (idx2 == -1 && idx1 == -1)
                            {
                                continue;
                            }
                            else
                            {
                                index_ = idx1 == -1 ? idx2 : idx1;
                                author = paragraph.Substring(0, index_).Trim();                                
                            }
                            int idx = author.IndexOf("、");
                            if (idx != -1)
                            {
                                string strT = author.Substring(idx + 1, author.Length - idx - 1);
                                author = strT;
                            }
                            if (authorDic.ContainsKey(author))
                            {
                                authorDic[author]++;
                            }
                            else
                            {
                                authorDic.Add(author, 1);
                            }*/
                        }
                        catch (Exception ex) { AssistUtils.showMessage(ex.ToString()); }
                        try
                        {//分两种情况：1有分号，则将字符串先分成几个Alex字符传输组根据分号，之后提取最后逗号之后的部分作为机构名称；2无分号，则
                            
                            int idx1 = paragraph.IndexOf(";");
                            if (idx1 == -1)
                            {
                                idx1 = paragraph.LastIndexOf(",");//Inc.
                                Company c = new Company();
                                c.year = year;
                                c.name = paragraph.Substring(idx1 + 1, paragraph.Length - idx1 - 1).Trim();
                                if (!c.name.Contains(" and"))
                                {
                                    if (companyDic.ContainsKey(c))
                                    {
                                        companyDic[c]++;
                                    }
                                    else
                                    {
                                        companyDic.Add(c, 1);
                                    }
                                }
                                else
                                {
                                    Company c2 = new Company();
                                    Company c1 = new Company();
                                    string[] strArr = Regex.Split(c.name, " and");
                                    c2.name = strArr[1].Trim();
                                    c1.name = strArr[0].Trim();
                                    c1.year = year;
                                    c2.year = year;
                                    if (companyDic.ContainsKey(c1))
                                    {
                                        companyDic[c1]++;
                                    }
                                    else
                                    {
                                        companyDic.Add(c1, 1);
                                    }
                                    if (companyDic.ContainsKey(c2))
                                    {
                                        companyDic[c2]++;
                                    }
                                    else
                                    {
                                        companyDic.Add(c2, 1);
                                    }
                                }
                            }
                            else
                            {
                                List<string> list = new List<string>();   
                                string[] arr = Regex.Split(paragraph, ";");
                                //每一个字符串内部提取公司名称
                                for (int i = 0; i < arr.Length;i ++ )
                                {
                                    int idx2 = arr[i].LastIndexOf(",");
                                    if (idx2 == -1)
                                        continue;
                                    string t = arr[i].Substring(idx2 + 1, arr[i].Length - idx2 - 1).Trim();
                                    if (!list.Contains(t))
                                        list.Add(t);
                                }
                                
                                for (int i = 0; i < list.Count; i++)
                                { 
                                    Company company = new Company();
                                    company.year = year;
                                    company.name = list[i];
                                    if(companyDic.ContainsKey(company))
                                        companyDic[company]++;
                                    else
                                        companyDic.Add(company, 1);
                                }
                            }
                        }
                        catch (Exception ex) { AssistUtils.showMessage(ex.ToString()); }
                    }
                    catch
                    {
                        report.CloseDocument();
                        MessageBox.Show("success");
                        break;
                    }

                }
                //结果呈现，写入到word中
                object missing = System.Reflection.Missing.Value;
                object start;
                object end;
                //打开word模板文档
                try
                {
                    report.CreateNewDocument(System.Windows.Forms.Application.StartupPath + "\\model.dot");
                }
                catch
                {
                    AssistUtils.showMessage("模板文件损坏,无法保存！");
                    return;
                }

                object bk_start = "start";
                try
                {
                    start = report.Document.Bookmarks.get_Item(ref bk_start).Start;
                }
                catch
                {
                    AssistUtils.showMessage("模板文件不存在或者被损坏");
                    report.CloseDocument();
                    return;
                }

                Word.Range rangeBasicSetting = null;
                end = start;
                rangeBasicSetting = report.Document.Range(ref start, ref end);
                rangeBasicSetting.Select();

                List<KeyValuePair<string, int>> list_author = new List<KeyValuePair<string, int>>(authorDic);
                list_author.Sort(delegate(KeyValuePair<string, int> v1, KeyValuePair<string, int> v2)
                {
                    return v2.Value.CompareTo(v1.Value);
                });
                foreach (KeyValuePair<string, int> pair in list_author)
                {
                    report.Application.Selection.TypeText("author:" + pair.Key + ",次数:" + pair.Value);
                    report.Application.Selection.TypeParagraph();//Alex
                }

                report.Application.Selection.TypeParagraph();

                List<KeyValuePair<Company, int>> listCom = new List<KeyValuePair<Company, int>>(companyDic);
                listCom.Sort(delegate(KeyValuePair<Company, int> c1, KeyValuePair<Company, int> c2)
                {
                    return String.Compare(c2.Key.year, c1.Key.year);//sort by year.
                    //return c2.Value.CompareTo(c1.Value);
                });
                string y = "";                
                int numOfYear = 0;
                foreach (KeyValuePair<Company, int> pair in listCom)
                {
                    if (pair.Key.year != null && !y.Equals(pair.Key.year))
                    {//
                        numOfYear++;
                        y = pair.Key.year;
                    }
                }
                Dictionary<Company, int>[] dicYear = new Dictionary<Company, int>[numOfYear];
                int idxOfYear = 0;
                if (listCom.Count > 0)
                    y = listCom[0].Key.year;
                else return;
                for (int i = 0; i < numOfYear; i++)
                {
                    dicYear[i] = new Dictionary<Company, int>();
                }
                foreach (KeyValuePair<Company, int> pair in listCom)
                {
                    if (pair.Key.year != null && y != pair.Key.year)
                    {
                        idxOfYear++;                        
                        if (idxOfYear >= numOfYear)
                            break;
                    }                    
                    dicYear[idxOfYear].Add(pair.Key, pair.Value);
                    y = pair.Key.year;
                }
                for (int i = 0; i < dicYear.Length; i++)
                {
                    //sort by count
                    List<KeyValuePair<Company, int>> lTemp = new List<KeyValuePair<Company, int>>(dicYear[i]);
                    lTemp.Sort(delegate(KeyValuePair<Company, int> v1, KeyValuePair<Company, int> v2)
                    {
                        return v2.Value.CompareTo(v1.Value);
                    });
                    if (comboBox1.Text.Trim().Equals("按照年份统计"))
                    {
                        report.Application.Selection.TypeText("year:" + lTemp[0].Key.year);
                        report.Application.Selection.TypeParagraph();
                        foreach (KeyValuePair<Company, int> pair in lTemp)
                        {
                            if (!pair.Key.name.Equals(""))
                            {
                                report.Application.Selection.TypeText("company:" + pair.Key.name + "      出现次数:" + pair.Value);
                                report.Application.Selection.TypeParagraph();
                            }
                        }
                        report.Application.Selection.TypeParagraph();
                    }
                }
                if (comboBox1.Text.Equals("按照总体统计"))
                {
                    Dictionary<string,int> dicTemp = new Dictionary<string,int >();
                    foreach (KeyValuePair<Company, int> pair in companyDic)
                    {
                        string str = pair.Key.name;
                        if (!dicTemp.ContainsKey(str))
                        {
                            dicTemp.Add(str, 1);
                        }
                        else
                        {
                            dicTemp[str]++;
                        }
                    }
                    List<KeyValuePair<string, int>> list_CompanyTotal = new List<KeyValuePair<string, int>>(dicTemp);
                    list_CompanyTotal.Sort(delegate(KeyValuePair<string, int> v1, KeyValuePair<string, int> v2)
                    {
                        return v2.Value.CompareTo(v1.Value);
                    });
                    foreach (KeyValuePair<string, int> pair in list_CompanyTotal)
                    {
                        if (pair.Key.Length > 2)
                        {
                            report.Application.Selection.TypeText("Institude:" + pair.Key + ",             次数:" + pair.Value);
                            report.Application.Selection.TypeParagraph();//Alex
                        }
                    }
                }
                
                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Title = "请输入文件名";
                saveFileDialog.Filter = "word文档(*.doc)|*.doc|word文档(*.docx)|*.docx";
                saveFileDialog.FilterIndex = 1;
                saveFileDialog.RestoreDirectory = true;
                saveFileDialog.FileName = "result";
                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        object localFilePath = saveFileDialog.FileName.ToString();
                        report.Document.SaveAs(ref localFilePath, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
                        report.Document.Close(ref missing, ref missing, ref missing);
                        report.Application.Quit(ref missing, ref missing, ref missing);
                        AssistUtils.showRightMessage("保存成功！");
                    }
                    catch
                    {
                        AssistUtils.showMessage("对不起,目标文件可能正在被使用,无法保存！");
                        object bSaveChanges = false;
                        report.Document.Close(ref bSaveChanges, ref missing, ref missing);
                        report.Application.Quit(ref missing, ref missing, ref missing);
                    }

                }
                else
                {
                    object bSaveChanges = false;
                    report.Document.Close(ref bSaveChanges, ref missing, ref missing);
                    report.Application.Quit(ref missing, ref missing, ref missing);
                }
            }
        }
    }
}
