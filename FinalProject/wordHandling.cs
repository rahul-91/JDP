using System;
using System.Collections.Generic;
using System.Text;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Word;

namespace FinalProject
{
    class wordHandling
    {
        public static Word._Document wd = new Word.Document();
        public static Application wordApp = new Application();

        public static void openDocument()
        {
            wd = wordApp.Documents.Open(System.Windows.Forms.Application.StartupPath.ToString() + "\\Webconreport.docx", ReadOnly: false, Visible: false);
        }

        public static void writetoWord(System.Data.DataTable dt, string ReportHeader, string TableHeader, int inRange, int intTables)
        {


            wd.Activate();
            int ColCount = dt.Columns.Count;

            Object start = Type.Missing;

            Object end = Type.Missing;

            start = Properties.Settings.Default.start;
            end = Properties.Settings.Default.end;

            Word.Range rng = wd.Range(ref start, ref end);

            rng.InsertBefore(TableHeader);
            rng.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            rng.Font.Name = "Arial";
            rng.Font.Size = 10;
            rng.Font.Bold = 1;
           
            rng.InsertParagraphAfter();

            rng.SetRange(rng.End, rng.End);
             
            rng.Text = ReportHeader;
            rng.Font.Name = "Arial";
            rng.Font.Size = 10;
            rng.Font.Bold = 5;
            rng.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

            rng.InsertParagraphAfter();
            
            rng.SetRange(rng.End, rng.End);

            rng.InsertAfter(ReportHeader);
            
            Object defaultTableBehavior = Type.Missing;

            Object autoFitBehavior = Type.Missing;

            object missing = System.Type.Missing;
          
           
            Word.Table tbl = wd.Tables.Add(rng,1,ColCount,defaultTableBehavior,autoFitBehavior);
            
            for (int i = 0; i < ColCount; i++)
            {
                tbl.Cell(1, i + 1).Range.Text = dt.Columns[i].ColumnName;
                tbl.Cell(1, i + 1).WordWrap = true;
                tbl.Cell(1, i + 1).VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalTop;
                tbl.Cell(1, i + 1).Range.Font.Bold = 1;
                tbl.Cell(1, i + 1).SetHeight(14f, WdRowHeightRule.wdRowHeightExactly); 
            }

            for (int i = 0; i < dt.Rows.Count; i++)
            {

                Word.Row newRow = wd.Tables[intTables].Rows.Add(ref missing);

                newRow.Range.Font.Bold = 0;

                newRow.SetHeight(14f, WdRowHeightRule.wdRowHeightExactly);

                newRow.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;

                for (int j = 0; j < ColCount; j++)
                {
                    newRow.Cells[j + 1].Range.Text = dt.Rows[i][j].ToString();
                }

            }
    
            tbl.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle; 
            tbl.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;

            
            rng.SetRange(wd.Tables[intTables].Range.End, wd.Tables[intTables].Range.End);
            object oPageBreak = Word.WdBreakType.wdPageBreak;
            rng.InsertBreak(ref oPageBreak);
            Properties.Settings.Default.start = rng.End;
            Properties.Settings.Default.end = rng.End;
           
            
            
           
            //Properties.Settings.Default.start = rng.End;
            //Properties.Settings.Default.end = rng.End;
             
           
            foreach (Word.Section wordSection in wd.Sections)
            {
                Word.Range footerRange = wordSection.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                footerRange.Font.ColorIndex = Word.WdColorIndex.wdBlack;
                footerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                footerRange.Font.Size = 11;
                footerRange.Font.Name = "Monotype Corsiva";
                footerRange.Text = "WEBCON/SSG/2094/12";
                
                Word.Range headerRange = wordSection.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                headerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                headerRange.Font.Size = 11;
                headerRange.Font.Name = "Monotype Corsiva";
                headerRange.Text = "Physical Survey of Producers & Exporters of JDPs";
               
                object oMissing = System.Reflection.Missing.Value;
                Microsoft.Office.Interop.Word.InlineShape line = headerRange.InlineShapes.AddHorizontalLineStandard();
                line.Height = 1;
                line.Fill.Solid();
                line.HorizontalLineFormat.NoShade = true;
                line.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                line.HorizontalLineFormat.PercentWidth = 100;
             

                Microsoft.Office.Interop.Word.InlineShape fline = wordSection.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.InlineShapes.AddHorizontalLineStandard(ref oMissing);
                fline.Height = 1;
                fline.Fill.Solid();
                fline.HorizontalLineFormat.NoShade = true;
                fline.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                fline.HorizontalLineFormat.PercentWidth = 100;
                
              
            }
            
        }

        public static void CloseandSave()
        {
            
            wd.SaveAs(System.Windows.Forms.Application.StartupPath.ToString() + "\\Webconreport.docx");
            wd.Close();
        }
    }
}
