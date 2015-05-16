using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;


namespace WPFDailyLog
{
    partial class DailyLogWorksheet : LogWorkSheet
    {
        /* This class will build the daily log worksheets that are used to enter daily activities. It will take a 
         * reference to a worksheet in the constructor
         */
        //private Microsoft.Office.Interop.Excel.Worksheet wkSheet;
        private Microsoft.Office.Interop.Excel.Range range;
        private string sheetDate;

        //int day;
        // Constructor
        public DailyLogWorksheet() {
            //this.wkSheet = wkSheet;
            //this.day = day;
            //wkSheet.Name = day.ToString();
            //fillDailyLogCells();
        }

        
        public void AddHiddenData(Worksheet wkSheet) {
            int ctr = 72;
            string newRange;
            string statusString = string.Join(",", Status.ToArray());
            for (double i = 1; i <= 10; i+=.5) { // add hours to worksheet
                newRange = "M" + ctr.ToString();
                range = wkSheet.get_Range(newRange); 
                range.Value = i.ToString("f1");
                ctr++;
            }
            ctr = 72;
            foreach (String status in Status) { // Add status strings in worksheet
                newRange = "BQ" + ctr.ToString();
                range = wkSheet.get_Range(newRange);
                range.Value = status.ToString();
                ctr++;
            }

            ctr = 72;
            foreach (String activity in activityTypes) {
                wkSheet.get_Range("CL" + ctr.ToString()).Value = activity;
                ctr++;
            }
            ctr--;
            for (int i = 28; i <= 64; i+=2) {
                wkSheet.get_Range("AE" + i.ToString()).Validation.Add(XlDVType.xlValidateList, XlDVAlertStyle.xlValidAlertInformation, XlFormatConditionOperator.xlBetween, "=$CL72:CL"+ctr.ToString(), Type.Missing);
            }

            for (int i = 7; i <= 25; i+=2) { // Add in-cell dropdown for Statsu and Hours for crewmembers
                newRange = "P" + i.ToString();  // Status
                range = wkSheet.get_Range(newRange);
                range.Validation.Add(XlDVType.xlValidateList, XlDVAlertStyle.xlValidAlertInformation, XlFormatConditionOperator.xlBetween, "=$BQ$72:$BQ$84", Type.Missing);
                newRange = "AB" + i.ToString();  // Hours
                range = wkSheet.get_Range(newRange);
                range.Validation.Add(XlDVType.xlValidateList, XlDVAlertStyle.xlValidAlertInformation, XlFormatConditionOperator.xlBetween, "=$M$72:$M$90", Type.Missing);

            }

            for (int i = 73; i <= 150; i++) { // Add start date of PayPeriods
                newRange = "U" + i.ToString();
                range = wkSheet.get_Range(newRange);
                if (i == 73)
                    range.Value = "1/11/2015";
                else
                    range.Formula = "=U" + (i - 1).ToString() + "+14";
            }

            for (int i = 73; i <= 150; i++) { // Add end date of Pay Periods
                newRange = "V" + i.ToString();
                range = wkSheet.get_Range(newRange);
                if (i == 73)
                    range.Value = "=U73+13";
                else
                    range.Formula = "=V" + (i - 1).ToString() + "+14";
            }

            ctr = 1;
            for (int i = 73; i <= 150; i++) { // Add pay period numbers
                newRange = "W" + i.ToString();
                range = wkSheet.get_Range(newRange);
                if (ctr == 27)
                    ctr = 1;
                range.Value = ctr++;
                //ctr++;
            }

            ctr = 72;
            //forestList.Sort();
            foreach (Forest forest in forestList) {
                // AC, AB, AE 72-90
                wkSheet.get_Range("AC" + ctr.ToString()).Value = forest.ForestName;
                wkSheet.get_Range("AD" + ctr.ToString()).Value = forest.Abbr;
                wkSheet.get_Range("AE" + ctr.ToString()).Value = forest.ForestNumber;
                ctr++;
            }
            
        }
        // build sheet-outlines, formats, cell sizing
        public void buildSheet(Worksheet wkSheet, int day, int monthNum, int yearNum) {
            wkSheet.Name = day.ToString();
            sheetDate = monthNum.ToString() + "/" + day.ToString() + "/" + yearNum.ToString();
            fillDailyLogCells();
            AddHiddenData(wkSheet);
            wkSheet.Columns.EntireColumn.ColumnWidth = 1;
            wkSheet.Rows.EntireRow.RowHeight = 12;
            range = wkSheet.get_Range("AC1");
            range.EntireColumn.ColumnWidth = 3;
            wkSheet.Names.Add("leaveTypes", Properties.Resources.LeaveType);
            range = wkSheet.get_Range("J4");
            range.NumberFormat = "dddd, mmmm dd";
            base.setUpWorkSheet(dailyLogCells, wkSheet);
            

        }



        // add formulas
        private void setUpWorkSheetFormulas() {
            
        }
        // add text to cells
        private void addTextToCells() {

        }

    }
}
