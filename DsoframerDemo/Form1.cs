using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace DsoframerDemo
{
    public partial class Form1 : Form
    {
        Microsoft.Office.Interop.Word.Document document = null;

        public Form1()
        {
            InitializeComponent();

            this.axFramerControl1.Open(Path.Combine(System.Windows.Forms.Application.StartupPath, "AppData", "sample.docx"));
            document = this.axFramerControl1.ActiveDocument as Microsoft.Office.Interop.Word.Document;

            AddSimpleTable(document.Application, document, 10, 3, WdLineStyle.wdLineStyleDashDot, WdLineStyle.wdLineStyleDashDot);

            //在书签处插入文字
            object oStart = "custom";
            Range range = document.Bookmarks.get_Item(ref oStart).Range;
            range.Text = "这里是您要输入的内容";
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            this.axFramerControl1.Close();

            base.OnFormClosing(e);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Range range = document.Application.Selection.Range;
            range.Text = "${userName}";
            document.Bookmarks.Add("userName", range);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            object oStart = "userName";
            Range range = document.Bookmarks.get_Item(ref oStart).Range;
            range.Text = "这里是您要输入的内容";
        }

        public static void AddSimpleTable(Microsoft.Office.Interop.Word.Application WordApp, Document WordDoc, int numrows, int numcolumns, WdLineStyle outStyle, WdLineStyle intStyle)
        {
            Object Nothing = System.Reflection.Missing.Value;
            //文档中创建表格   
            Microsoft.Office.Interop.Word.Table newTable = WordDoc.Tables.Add(WordApp.Selection.Range, numrows, numcolumns, ref Nothing, ref Nothing);
            //设置表格样式   
            newTable.Borders.OutsideLineStyle = outStyle;
            newTable.Borders.InsideLineStyle = intStyle;
            newTable.Columns[1].Width = 100f;
            newTable.Columns[2].Width = 220f;
            newTable.Columns[3].Width = 105f;

            //填充表格内容   
            newTable.Cell(1, 1).Range.Text = "产品详细信息表";
            newTable.Cell(1, 1).Range.Bold = 2;//设置单元格中字体为粗体   
                                               //合并单元格   
            newTable.Cell(1, 1).Merge(newTable.Cell(1, 3));
            WordApp.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;//垂直居中   
            WordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;//水平居中   

            //填充表格内容   
            newTable.Cell(2, 1).Range.Text = "产品基本信息";
            newTable.Cell(2, 1).Range.Font.Color = WdColor.wdColorDarkBlue;//设置单元格内字体颜色   
                                                                           //合并单元格   
            newTable.Cell(2, 1).Merge(newTable.Cell(2, 3));
            WordApp.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;

            //填充表格内容   
            newTable.Cell(3, 1).Range.Text = "品牌名称：";
            newTable.Cell(3, 2).Range.Text = "品牌名称：";
            //纵向合并单元格   
            newTable.Cell(3, 3).Select();//选中一行   
            object moveUnit = WdUnits.wdLine;
            object moveCount = 5;
            object moveExtend = WdMovementType.wdExtend;
            WordApp.Selection.MoveDown(ref moveUnit, ref moveCount, ref moveExtend);
            WordApp.Selection.Cells.Merge();


            newTable.Cell(12, 1).Range.Text = "产品特殊属性";
            newTable.Cell(12, 1).Merge(newTable.Cell(12, 3));

            //在表格中增加行   
            WordDoc.Content.Tables[1].Rows.Add(ref Nothing);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Word.Application objApp = new Microsoft.Office.Interop.Word.Application();
            Document document = objApp.Documents.Open(Path.Combine(System.Windows.Forms.Application.StartupPath, "AppData", "sample1.docx"));

            object missing = System.Reflection.Missing.Value;
            object confirmConversion = false;
            object link = false;
            object attachment = false;
            objApp.Selection.InsertFile(Path.Combine(System.Windows.Forms.Application.StartupPath, "AppData", "sample2.docx"), ref missing, ref confirmConversion, ref link, ref attachment);
            object pBreak = (int)Microsoft.Office.Interop.Word.WdBreakType.wdSectionBreakNextPage;
            objApp.Selection.InsertBreak(ref pBreak);

            objApp.ActiveDocument.Save();
            document.Close();
        }
    }
}
