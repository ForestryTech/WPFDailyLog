using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace WPFDailyLog
{
    partial class StationInfoSheet : LogWorkSheet
    {
        Microsoft.Office.Interop.Excel.Worksheet wkSheet;
        Microsoft.Office.Interop.Excel.Range range;
        private string date;
        private string station;
        private string forest;
        private string schedule;
        private List<Employee> employee;
        private Dictionary<string, string> daysOff;

        public StationInfoSheet(Microsoft.Office.Interop.Excel.Worksheet wksht, string date, string station, string forest, string schedule,List<Employee> employee, Dictionary<string, string> daysOff) {
            this.wkSheet = wksht;
            this.date = date;
            this.station = station;
            this.forest = forest;
            this.schedule = schedule;
            this.employee = employee;
            this.daysOff = daysOff;
            this.employee = employee;
            fillStationInfoCells();
        }

        public void buildSheet() {
            string employeeNameCell;
            string employeeDayOffCell;
            string empName;
            resizeColumns();
            base.setUpWorkSheet(headerCells, wkSheet);
            base.setUpWorkSheet(crewInfoCells, wkSheet);
            
            wkSheet.Name = "Station Info";
            range = wkSheet.get_Range("E1");
            range.NumberFormat = "mmmm yyyy";


            for (int i = 5; i < 15; i++) {
                employeeNameCell = "B" + i.ToString();
                employeeDayOffCell = "E" + i.ToString();
                
                if (!(employee[i - 5].Last == "EMPTY")) {
                    range = wkSheet.get_Range(employeeNameCell);
                    empName = employee[i - 5].Last + ", " + employee[i - 5].First;
                    range.Value = empName;
                    range = wkSheet.get_Range(employeeDayOffCell);
                    range.Value = daysOff[empName];
                }
            }

            addHiddenData();
            addInCellDropDowns();
            range = wkSheet.get_Range("G5:M14");

            FormatCondition format = (FormatCondition)(wkSheet.get_Range("G5:M14", Type.Missing).FormatConditions.Add(XlFormatConditionType.xlCellValue, XlFormatConditionOperator.xlEqual,
                "=\"ON\"", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing));
            format.Font.Bold = true;
            format.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            format.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(221, 235, 247));

            FormatCondition offFormat = (FormatCondition)(wkSheet.get_Range("G5:M14", Type.Missing).FormatConditions.Add(XlFormatConditionType.xlCellValue, XlFormatConditionOperator.xlEqual,
                "=\"OFF\"", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing));
            offFormat.Font.Bold = true;
            offFormat.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(156,0,6));
            offFormat.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(169,208,142));
            
        }

        private void resizeColumns() {
            range = wkSheet.get_Range("A:A");
            range.ColumnWidth = 10;
            range = wkSheet.get_Range("E:E");
            range.ColumnWidth = 30;
            range = wkSheet.get_Range("G:G");
            range.ColumnWidth = 9;
            range = wkSheet.get_Range("H:H");
            range.ColumnWidth = 9;
            range = wkSheet.get_Range("I:I");
            range.ColumnWidth = 9;
            range = wkSheet.get_Range("J:J");
            range.ColumnWidth = 11;
            range = wkSheet.get_Range("K:K");
            range.ColumnWidth = 9;
            range = wkSheet.get_Range("L:L");
            range.ColumnWidth = 9;
            range = wkSheet.get_Range("M:M");
            range.ColumnWidth = 9;
        }

        private void addHiddenData() {
            int ctr = 23;
            foreach (string day in daysOffList) {
                range = wkSheet.get_Range("A" + ctr.ToString());
                range.Value = day;
                ctr++;
            }
            range = wkSheet.get_Range("A23:A36");
            range.EntireRow.Hidden = true;
        }

        private void addInCellDropDowns() {
            string newRange;
            for (int i = 5; i < 15; i++) {
                newRange = "E" + i.ToString();  // Status
                range = wkSheet.get_Range(newRange);
                range.Validation.Add(XlDVType.xlValidateList, XlDVAlertStyle.xlValidAlertInformation, XlFormatConditionOperator.xlBetween, "=$A$23:$A$36", Type.Missing);
            }
        }


    }
}
