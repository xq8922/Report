
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Diagnostics;
using Microsoft.Office.Interop.Word;


namespace ReportDesign
{
    class Report
    {
        private _Application wordApp = null;
        private _Document wordDoc = null;
        public _Application Application
        {
            get
            {
                return wordApp;
            }
            set
            {
                wordApp = value;
            }
        }
        public _Document Document
        {
            get
            {
                return wordDoc;
            }
            set
            {
                wordDoc = value;
            }
        }
        //ͨ��ģ�崴�����ĵ�
        public void CreateNewDocument(string filepath)
        {
            killWinWordProcess();

            wordApp = new ApplicationClass();
            wordApp.DisplayAlerts = WdAlertLevel.wdAlertsNone;
            wordApp.Visible = false;
            object missing = System.Reflection.Missing.Value;
            object templateName = filepath;
            wordDoc = wordApp.Documents.Open(ref templateName, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing);
        }

        //save new file
        public void SaveDocument(string filepath)
        {
            object fileName = filepath;
            object format = WdSaveFormat.wdFormatDocument;
            object miss = System.Reflection.Missing.Value;
            wordDoc.SaveAs(ref fileName, ref format, ref miss,
                ref miss, ref miss, ref miss, ref miss,
                ref miss, ref miss, ref miss, ref miss,
                ref miss, ref miss, ref miss, ref miss,
                ref miss);
            object SaveChanges = WdSaveOptions.wdSaveChanges;
            object OriginalFormat = WdOriginalFormat.wdOriginalDocumentFormat;
            object RouteDocument = false;
            wordDoc.Close(ref SaveChanges, ref OriginalFormat, ref RouteDocument);
            wordApp.Quit(ref SaveChanges, ref OriginalFormat, ref RouteDocument);
        }

        public void CloseDocument()
        {
            try
            {
                object SaveChanges = WdSaveOptions.wdSaveChanges;
                object OriginalFormat = WdOriginalFormat.wdOriginalDocumentFormat;
                object RouteDocument = false;
                wordDoc.Close(ref SaveChanges, ref OriginalFormat, ref RouteDocument);
                wordApp.Quit(ref SaveChanges, ref OriginalFormat, ref RouteDocument);
            }
            catch { }
        }

        public void UpdateContents()
        {
            int count = wordDoc.TablesOfContents.Count;
            for (int i = 0; i < count; i++)
            {
                wordDoc.TablesOfContents[i + 1].UpdatePageNumbers();
            }  
        }

        //insert values where bookmarked
        public bool InsertValue(string bookmark, string value)
        {
            object bkObj = bookmark;
            if (wordApp.ActiveDocument.Bookmarks.Exists(bookmark))
            {
                wordApp.ActiveDocument.Bookmarks.get_Item(ref bkObj).Select();
                wordApp.Selection.TypeText(value);
                return true;
            }
            return false;
        }

        //insert table, bookmark
        public Table InsertTable(string bookmark, int rows, int columns, float width)
        {
            object miss = System.Reflection.Missing.Value;
            object oStart = bookmark;
            Range range = wordDoc.Bookmarks.get_Item(ref oStart).Range; //where to insert
            Table newTable = wordDoc.Tables.Add(range, rows, columns, ref miss, ref miss);
            //set style of table
            newTable.Borders.Enable = 1; //�����б߿�Ĭ��û�б߿�(Ϊ0ʱ����1Ϊʵ�߱߿�2��3Ϊ���߱߿��Ժ������û�Թ�)
            newTable.Borders.OutsideLineWidth = WdLineWidth.wdLineWidth050pt; //width of frame
            if (width != 0)
            {
                newTable.PreferredWidth = width;//width of table
            }
            newTable.AllowPageBreaks = false;
            return newTable;
        }

        //�ϲ���Ԫ�� ����,��ʼ�к�,��ʼ�к�,�����к�,�����к�
        public void MergeCell(Microsoft.Office.Interop.Word.Table table, int row1, int column1, int row2, int column2)
        {
            table.Cell(row1, column1).Merge(table.Cell(row2, column2));
        }

        //���ñ�����ݶ��뷽ʽ Alignˮƽ����Vertical��ֱ����(����룬���ж��룬�Ҷ���ֱ��ӦAlign��Vertical��ֵΪ-1,0,1)
        public void SetParagraph_Table(Microsoft.Office.Interop.Word.Table table, int Align, int Vertical)
        {
            switch (Align)
            {
                case -1: table.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft; break;
                case 0: table.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter; break;//ˮƽ����
                case 1: table.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight; break;//�Ҷ���
            }
            switch (Vertical)
            {
                case -1: table.Range.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalTop; break;//���˶���
                case 0: table.Range.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter; break;//��ֱ����
                case 1: table.Range.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalBottom; break;//�׶˶���
            }
        }

        //���ñ������
        public void SetFont_Table(Microsoft.Office.Interop.Word.Table table, string fontName, double size)
        {
            if (size != 0)
            {
                table.Range.Font.Size = Convert.ToSingle(size);
            }
            if (fontName != "")
            {
                table.Range.Font.Name = fontName;
            }
        }

        //�Ƿ�ʹ�ñ߿�,n�������,use�ǻ��
        public void UseBorder(int n, bool use)
        {
            if (use)
            {
                wordDoc.Content.Tables[n].Borders.Enable = 1;//�����б߿�Ĭ��û�б߿�(Ϊ0ʱ����1Ϊʵ�߱߿�2��3Ϊ���߱߿��Ժ������û�Թ�)
            }
            else
            {
                wordDoc.Content.Tables[n].Borders.Enable = 2;//�����б߿�Ĭ��û�б߿�(Ϊ0ʱ����1Ϊʵ�߱߿�2��3Ϊ���߱߿��Ժ������û�Թ�)
            }
        }

        //��������һ��,n������Ŵ�1��ʼ��
        public void AddRow(int n)
        {
            object miss = System.Reflection.Missing.Value;
            wordDoc.Content.Tables[n].Rows.Add(ref miss);
        }

        //��������һ��
        public void AddRow(Microsoft.Office.Interop.Word.Table table)
        {
            object miss = System.Reflection.Missing.Value;
            table.Rows.Add(ref miss);
        }

        //��������rows��,nΪ�������
        public void AddRow(int n, int rows)
        {
            object miss = System.Reflection.Missing.Value;
            Microsoft.Office.Interop.Word.Table table = wordDoc.Content.Tables[n];
            for (int i = 0; i < rows; i++)
            {
                table.Rows.Add(ref miss);
            }
        }

        public Microsoft.Office.Interop.Word.Table FindTable(string bookmark)
        {
            object oBookmark = bookmark;
            object start = wordDoc.Bookmarks.get_Item(ref oBookmark).Start;
            object end = wordDoc.Content.End;
            Microsoft.Office.Interop.Word.Range rangeTable = wordDoc.Range(ref start, ref end);
            return rangeTable.Tables[1];
        }

        //������е�Ԫ�����Ԫ�أ�table���ڱ��row�кţ�column�кţ�value�����Ԫ��
        public void InsertCell(Microsoft.Office.Interop.Word.Table table, int row, int column, string value)
        {
            table.Cell(row, column).Range.Text = value;
        }

        //������е�Ԫ�����Ԫ�أ�n������Ŵ�1��ʼ�ǣ�row�кţ�column�кţ�value�����Ԫ��
        public void InsertCell(int n, int row, int column, string value)
        {
            wordDoc.Content.Tables[n].Cell(row, column).Range.Text = value;
        }

        //��������һ�����ݣ�nΪ������ţ�row�кţ�columns������values�����ֵ
        public void InsertCell(int n, int row, int columns, string[] values)
        {
            Microsoft.Office.Interop.Word.Table table = wordDoc.Content.Tables[n];
            for (int i = 0; i < columns; i++)
            {
                table.Cell(row, i + 1).Range.Text = values[i];
            }
        }

        public void InsertCell(Microsoft.Office.Interop.Word.Table t, int row, int columns, string[] values)
        {
            Microsoft.Office.Interop.Word.Table table = t;
            for (int i = 0; i < columns; i++)
            {
                table.Cell(row, i + 1).Range.Text = values[i];
            }
        }

        public void InsertCell(Microsoft.Office.Interop.Word.Table t, int row, int columns, List<string[]> list)
        {
            Microsoft.Office.Interop.Word.Table table = t;
            int i = 0;
            foreach (string[] s in list)
            {
                InsertCell(table, row + i, columns, s);
                i++;
            }
        }

        //����ͼƬ
        public void InsertPicture(string bookmark, string picturePath, float width, float hight)
        {
            object miss = System.Reflection.Missing.Value;
            object oStart = bookmark;
            Object linkToFile = false;       //ͼƬ�Ƿ�Ϊ�ⲿ����
            Object saveWithDocument = true;  //ͼƬ�Ƿ����ĵ�һ�𱣴� 
            object range = wordDoc.Bookmarks.get_Item(ref oStart).Range;
            wordDoc.InlineShapes.AddPicture(picturePath, ref linkToFile, ref saveWithDocument, ref range);
            wordDoc.Application.ActiveDocument.InlineShapes[1].Width = width;   //����ͼƬ���
            wordDoc.Application.ActiveDocument.InlineShapes[1].Height = hight;  //����ͼƬ�߶�
        }

        public void InsertGraphToWordFromClipBoard(string bookmark)
        {
            object obDD_Name = bookmark;
            wordDoc.Bookmarks.get_Item(ref obDD_Name).Range.Paste();
        }


        //����һ������,textΪ��������
        public void InsertText(string bookmark, string text)
        {
            object oStart = bookmark;
            object range = wordDoc.Bookmarks.get_Item(ref oStart).Range;
            Paragraph wp = wordDoc.Content.Paragraphs.Add(ref range);
            wp.Format.SpaceBefore = 6;
            wp.Range.Text = text;
            wp.Format.SpaceAfter = 24;
            wp.Range.InsertParagraphAfter();
            wordDoc.Paragraphs.Last.Range.Text = "\n";
        }

        // ɱ��winword.exe����
        public void killWinWordProcess()
        {
            System.Diagnostics.Process[] processes = System.Diagnostics.Process.GetProcessesByName("WINWORD");
            foreach (System.Diagnostics.Process process in processes)
            {
                bool b = process.MainWindowTitle == "";
                if (process.MainWindowTitle == "")
                {
                    process.Kill();
                }
            }
        }
    }
}
