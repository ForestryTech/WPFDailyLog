using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace WPFDailyLog
{
    class LogWorkSheet
    {

        protected void setUpWorkSheet(List<DailyLogCells> dailyLogCells, Microsoft.Office.Interop.Excel.Worksheet wkSheet) {
            
            foreach (DailyLogCells cells in dailyLogCells) {
                Microsoft.Office.Interop.Excel.Range range;
                range = wkSheet.get_Range(cells.Range);
                range.Font.Name = "Calibri";
                range.Merge();

                if (cells.CellBorder == CellBorder.Yes) {
                    range.Borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
                    range.Borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
                    range.Borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
                    range.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                }

                if (cells.IsFormula) {
                    range.Formula = "=" + cells.CellContents;
                    //range.Hidden = xl;
                    range.FormulaHidden = true;
                } else
                    range.Value = cells.CellContents;

                range.Font.Size = cells.CellText.Size;
                range.VerticalAlignment = XlVAlign.xlVAlignCenter;

                switch (cells.CellText.Position) {
                    case FontPosition.Center:
                        range.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        break;
                    case FontPosition.Left:
                        range.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                        break;
                    case FontPosition.Right:
                        range.HorizontalAlignment = XlHAlign.xlHAlignRight;
                        break;
                    default:
                        range.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        break;
                }

                switch (cells.CellText.Style) {
                    case FontStyle.Bold:
                        range.Font.Bold = true;
                        break;
                    case FontStyle.BoldItalic:
                        range.Font.Bold = true;
                        range.Font.Italic = true;
                        break;
                    case FontStyle.Italic:
                        range.Font.Italic = true;
                        break;
                }

                switch (cells.CellText.Color) {
                    case FontColor.Red:
                        range.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                        
                        break;
                    case FontColor.Green:
                        range.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Green);
                        break;
                }

                if (cells.CellLocked)
                    range.Locked = true;
                else
                    range.Locked = false;
                //range = wkSheet.get_Range("A6:CT66");


            }
        }
    }
}
