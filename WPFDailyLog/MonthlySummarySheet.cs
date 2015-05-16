using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace WPFDailyLog
{
    partial class MonthlySummarySheet : LogWorkSheet
    {
        private string date;
        private Worksheet wkSheet;
        private int daysInMonth;
        private List<string> columnHeads;

        public MonthlySummarySheet(Worksheet wkSheet, int month, int year, int daysInMonth) {
            this.wkSheet = wkSheet;
            this.date = month.ToString() + "/1/" + year.ToString();
            //daysInMonth = DateTime.DaysInMonth(year, month);
            this.daysInMonth = daysInMonth;
            columnHeads = fillColumnHeads();
            fillMonthlySummaryLists();
        }

        public void buildSheet() {
            wkSheet.Name = "Monthly Summary";
            setUpWorkSheet(rowHeadings, wkSheet);
            addWorkSheetFormat();
            addMonthlySummaryFormulas();
            addTotalsFormula();
            addStatsBoxes();
        }

        public void addWorkSheetFormat() {
            wkSheet.get_Range("A:A").EntireColumn.ColumnWidth = 30;
            foreach (string column in columnHeads) {
                if(string.Compare(column, "A") != 0)
                    wkSheet.get_Range(column + ":" + column).EntireColumn.ColumnWidth = 4;
            }
            wkSheet.get_Range(columnHeads[0] + "1:" + columnHeads[columnHeads.Count - 1] + "1").Merge();
            wkSheet.get_Range(columnHeads[0] + "1:" + columnHeads[columnHeads.Count - 1] + "1").Value = date;
            wkSheet.get_Range(columnHeads[0] + "1:" + columnHeads[columnHeads.Count - 1] + "1").NumberFormat = "mmmm yyyy";
            wkSheet.get_Range(columnHeads[0] + "1:" + columnHeads[columnHeads.Count - 1] + "1").Font.Bold = true;
            wkSheet.get_Range(columnHeads[0] + "1:" + columnHeads[columnHeads.Count - 1] + "1").Font.Size = 16;
            wkSheet.get_Range(columnHeads[0] + "1:" + columnHeads[columnHeads.Count - 1] + "1").HorizontalAlignment = XlHAlign.xlHAlignCenter;
            wkSheet.get_Range(columnHeads[columnHeads.Count - 2] + ":" + columnHeads[columnHeads.Count - 2]).EntireColumn.ColumnWidth = 16;
            wkSheet.get_Range(columnHeads[columnHeads.Count - 1] + ":" + columnHeads[columnHeads.Count - 1]).EntireColumn.ColumnWidth = 10;
            wkSheet.get_Range("A28:" + columnHeads[columnHeads.Count - 1] + "28").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
            wkSheet.get_Range("A29:" + columnHeads[columnHeads.Count - 1] + "29").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
            wkSheet.get_Range("A43:" + columnHeads[columnHeads.Count - 1] + "43").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
            wkSheet.get_Range("A44:" + columnHeads[columnHeads.Count - 1] + "44").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
            wkSheet.get_Range("A47:" + columnHeads[columnHeads.Count - 1] + "49").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
            wkSheet.get_Range("A66:" + columnHeads[columnHeads.Count - 1] + "66").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
            wkSheet.get_Range("A67:" + columnHeads[columnHeads.Count - 1] + "67").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
            wkSheet.get_Range("A70:" + columnHeads[columnHeads.Count - 1] + "71").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
           
        }

        private void addMonthlySummaryFormulas() {
             // =SUM(SUMIF('1'!AU7:AU26,A(CURRENT ROW),'1'!BM7:BM26) + SUMIF('1'!BU7:BU26,A(CURRENT ROW),'1'!CM7:CM26))
            Range newRange;
            string formula;
             
            for (int i = 2; i <= 42; i++) {
                if(!(i == 28 || i == 29)){
                    for (int x = 1; x <= daysInMonth+2; x++) {
                        newRange = wkSheet.get_Range(columnHeads[x] + i.ToString());
                        if (i == 2) {
                            if (x == daysInMonth+2) {
                                newRange.Value = "Crew Days";
                                newRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
                            } else if (x == daysInMonth + 1) {
                                newRange.Value = "Total Person Hours";
                                newRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
                            } else
                                newRange.Value = x.ToString();

                            newRange.Font.Bold = true;
                            newRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                            newRange.Borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
                            newRange.Borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
                            newRange.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                        } else {
                            formula = "SUM(SUMIF('" + x.ToString() + "'!AU7:AU26,A" + i.ToString() + ",'" + x.ToString() + "'!BM7:BM26) + SUMIF('" + x.ToString() + "'!BU7:BU26,A" + i.ToString() + ",'" + x.ToString() + "'!CM7:CM26))";
                            if (x == daysInMonth+2) {
                                newRange.Value = "=" + columnHeads[columnHeads.Count - 2] + i.ToString() + "/40";
                                newRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightBlue);
                            } else if (x == daysInMonth + 1) {
                                newRange.Value = "=SUM(" + columnHeads[1] + i.ToString() + ":" + columnHeads[columnHeads.Count - 3] + i.ToString() + ")";
                                newRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightBlue);
                            } else
                                newRange.Value = "=" + formula;

                            newRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                            newRange.Borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
                            newRange.Borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
                            newRange.Borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
                            newRange.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                        }
                    }
                }

            }
        }

        private void addTotalsFormula() {
            string row = "46";
            string formula;
            Range newRange;
            for (int i = 1; i < daysInMonth + 1; i++) {
                // =SUM(B3:B27)+SUM(B30:B41)
                formula = "=SUM(" + columnHeads[i] + "3:" + columnHeads[i] + "27)+SUM(" + columnHeads[i] + "30:" + columnHeads[i] + "41)";
                wkSheet.get_Range(columnHeads[i] + row).Value = formula;
                wkSheet.get_Range(columnHeads[i] + row).Borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
                wkSheet.get_Range(columnHeads[i] + row).Borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
                wkSheet.get_Range(columnHeads[i] + row).Borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
                wkSheet.get_Range(columnHeads[i] + row).Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                wkSheet.get_Range(columnHeads[i] + row).HorizontalAlignment = XlHAlign.xlHAlignCenter;
            }
            newRange = wkSheet.get_Range(columnHeads[daysInMonth + 1] + row + ":" + columnHeads[daysInMonth + 2] + row);
            newRange.Merge();
            newRange.Value = "=SUM(B" + row + ":" + columnHeads[daysInMonth] + row + ")";
            newRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightBlue);
            newRange.Borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
            newRange.Borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
            newRange.Borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
            newRange.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
            newRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
        }

        private void addStatsBoxes() {
            Range newRange;
            string range;
            wkSheet.get_Range(columnHeads[columnHeads.Count - 2] + "49:" + columnHeads[columnHeads.Count - 1] + "49").Merge();
            wkSheet.get_Range(columnHeads[columnHeads.Count - 2] + "49:" + columnHeads[columnHeads.Count - 1] + "49").Value = "Total Responses";
            wkSheet.get_Range(columnHeads[columnHeads.Count - 2] + "49:" + columnHeads[columnHeads.Count - 1] + "49").HorizontalAlignment = XlHAlign.xlHAlignCenter;
            wkSheet.get_Range(columnHeads[columnHeads.Count - 2] + "49:" + columnHeads[columnHeads.Count - 1] + "49").Font.Bold = true;

            for (int i = 51; i <= 73; i++) {
                range = columnHeads[columnHeads.Count - 2] + i.ToString() + ":" + columnHeads[columnHeads.Count - 1] + i.ToString();
                wkSheet.get_Range(range).Merge();
                wkSheet.get_Range(range).HorizontalAlignment = XlHAlign.xlHAlignCenter;

            }
            newRange = wkSheet.get_Range("B51:" + columnHeads[columnHeads.Count - 1] + "65");
            addBorders(newRange);
            newRange = wkSheet.get_Range("B68:" + columnHeads[columnHeads.Count - 1] + "69");
            addBorders(newRange);
            newRange = wkSheet.get_Range("B72:" + columnHeads[columnHeads.Count - 1] + "73");
            addBorders(newRange);
        }

        private void addBorders(Range newRange) {
            newRange.Borders[XlBordersIndex.xlInsideHorizontal].LineStyle = XlLineStyle.xlContinuous;
            newRange.Borders[XlBordersIndex.xlInsideVertical].LineStyle = XlLineStyle.xlContinuous;
            newRange.Borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
            newRange.Borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
            newRange.Borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
            newRange.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
            newRange.BorderAround2(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium);
            newRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
        }

        private List<string> fillColumnHeads() {
            List<string> tempColumns = new List<string>();
            string col;
            char c;
            for (int i = 1; i <= daysInMonth+3; i++) {
                if ((i + 64) > 90) {
                    c = (char)((i - 26) + 64);
                    col = "A" + c.ToString();
                } else {
                    c = (char)(i + 64);
                    col = c.ToString();
                }

                tempColumns.Add(col);
            }
            return tempColumns;
        }
    }
}
