using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using System.Collections;

namespace ReportDesign
{
    class AssistUtils
    {
        public static bool DicHasKey(Dictionary<string, int> dic,string str)
        {
            foreach (KeyValuePair<string, int> pair in dic)
            {
                if (pair.Key.Equals(str))
                {
                    return true;
                }
            }
            return false;
        }

        public static void showMessage(string message)
        {
            MessageBox.Show(message, "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        public static void showRightMessage(string message)
        {
            MessageBox.Show(message, "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public static bool showMessageOkCancel(string message)
        {
            if (MessageBox.Show(message, "是否确认", MessageBoxButtons.OKCancel) == DialogResult.OK)
            {
                return true;
            }
            else
                return false;
        }

        public static bool showMessageYesNo(string message)
        {
            if (MessageBox.Show(message, "是否保存", MessageBoxButtons.YesNo) == DialogResult.OK)
            {
                return true;
            }
            else
                return false;
        }

        public static bool showMessageRetryCance(string message)
        {
            if (MessageBox.Show(message, "重试", MessageBoxButtons.RetryCancel) == DialogResult.Retry)
            {
                return true;
            }
            else
                return false;
        }

        /// <summary>
        /// 延时函数
        /// </summary>
        /// <param name="secend"></param>
        public static void delayTime(double msecend)
        {
            DateTime tempTime = DateTime.Now;
            while (tempTime.AddMilliseconds(msecend).CompareTo(DateTime.Now) > 0)
                System.Windows.Forms.Application.DoEvents();
        }

        //返回float型数的字符串,只显示float型小数点后四位
        public static string FormatFloattoString(int num,float target)
        {
            return target.ToString("f" + num);//.ToString("f2");
            /*
            string beforeComma;
            string afterComma;
            int temp;

            if (target.ToString().IndexOf(".") == -1)
            {
                return target.ToString();
            }
            else
            {
                beforeComma = target.ToString().Substring(0, target.ToString().IndexOf(".") + 1);
                afterComma = target.ToString().Substring(target.ToString().IndexOf(".") + 1);
                int len = afterComma.Length;
                while (len < num)
                {
                    afterComma += "0";
                    len++;
                }
                if (afterComma.Length <= num)
                {
                    return beforeComma + afterComma;
                }
                else
                {
                    if (int.Parse(target.ToString().Substring(target.ToString().IndexOf(".") + num + 1, 1)) >= 5)
                    {
                        temp = int.Parse(target.ToString().Substring(target.ToString().IndexOf(".") + 1, num));
                        temp = temp + 1;
                        //temp = temp / Math.Pow(10, num);
                        //afterComma = temp.ToString().Substring(temp.ToString().IndexOf(".") + 1, num);
                        afterComma = temp.ToString();
                    }
                    else
                    {
                        afterComma = target.ToString().Substring(target.ToString().IndexOf(".") + 1, num);
                    }
                    
                    return beforeComma + afterComma;
                }
            }
             * */
        }

       public static void InsertToWord(string bookmark)
        {
            Microsoft.Office.Interop.Word.Application appWord = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Document doc = new Microsoft.Office.Interop.Word.Document();
            object oMissing = System.Reflection.Missing.Value;
            object objTemplate = System.Windows.Forms.Application.StartupPath + @"\reportmodel.doc";
            object objDocType = WdDocumentType.wdTypeDocument;
            object objfalse = false;
            object objtrue = true;
            doc = (Document)appWord.Documents.Add(ref objTemplate, ref objfalse, ref objDocType, ref objtrue);
            Bookmarks odf = doc.Bookmarks;
            string testTableremarks = bookmark;
            object obDD_Name = testTableremarks;
            doc.Bookmarks.get_Item(ref obDD_Name).Range.Paste();
            object filename = System.Windows.Forms.Application.StartupPath + @"\test.doc";
            object miss = System.Reflection.Missing.Value;
            object format = WdSaveFormat.wdFormatDocument;
            doc.SaveAs(ref filename, ref format, ref miss, ref miss, ref miss, ref miss,
                ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss);
            object missingValue = Type.Missing;

            object SaveChanges = WdSaveOptions.wdSaveChanges;
            object OriginalFormat = WdOriginalFormat.wdOriginalDocumentFormat;
            object RouteDocument = false;
            doc.Close(ref SaveChanges, ref OriginalFormat, ref RouteDocument);


            object doNotSaveChanges = WdSaveOptions.wdDoNotSaveChanges;
            //doc.Close(ref doNotSaveChanges, ref missingValue, ref missingValue);
            //doc.Close(true, ref missingValue, ref missingValue);
            appWord.Application.Quit(ref miss, ref miss, ref miss);
            doc = null;
            appWord = null;
            AssistUtils.showMessage("生成成功！");
            System.Diagnostics.Process.Start(filename.ToString());//打開文檔
        }

        public static void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                AssistUtils.showMessage("Exception Occured while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}
