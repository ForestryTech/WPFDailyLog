using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace WPFDailyLog
{
    class DailyLogBook
    {
        Microsoft.Office.Interop.Excel.Application xlApp;
        Workbook wkBook;
        Worksheet wkSheet;
        // Set in contstructor
        private int daysInMonth;
        private int monthNum;
        private int yearNum;
        // Get and Set methods
        
        private string stationNumber;
        private string forest;
        private string schedule;
        private List<Employee> employee;
        private Dictionary<string, string> daysOff;


        private StationInfoSheet stationInfoSheet;
        private DailyLogWorksheet dailyLogWorkSheet;

        
        public DailyLogBook(string month, string year) {
            if (string.IsNullOrEmpty(month))
                monthNum = 1;
            else
                monthNum = DateTime.Parse("1." + month + "2015").Month;
            if (string.IsNullOrEmpty(year))
                yearNum = DateTime.Now.Year;
            else
                yearNum = Convert.ToInt32(year);
            this.stationNumber = "";
            this.forest = "";
            this.schedule = "";
            /*xlApp = new Application();
            
            //daysInMonth = DateTime.DaysInMonth(yearNum, monthNum);
            daysInMonth = 4;
            xlApp.SheetsInNewWorkbook = daysInMonth + 2;
            wkBook = xlApp.Workbooks.Add();*/
        }

        public void CreateWorkBook(){
            try {
                xlApp = new Application();
                string filename = System.Globalization.CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(monthNum) + "_" + yearNum.ToString() + ".xlsx";
                string path = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                path += @"\Daily_Log_" + yearNum.ToString();
                daysInMonth = DateTime.DaysInMonth(yearNum, monthNum);
                daysInMonth = 4;
                xlApp.SheetsInNewWorkbook = daysInMonth + 2;
                wkBook = xlApp.Workbooks.Add();
                xlApp.Visible = true;
                createStationInfoSheet();
                createDailyLogWorksheets();
                createMonthlySummarySheet();
                createFolderIfNotExist(path);
                wkBook.SaveAs(path + "\\" + filename);
                System.Windows.MessageBox.Show("Worksheet created. Saved as " + filename);
                
            }
            catch (Exception ex) {
                System.Windows.MessageBox.Show("An error occured: " + ex.Message);
            }
            finally {
                wkBook.Close(XlSaveAction.xlDoNotSaveChanges);
                xlApp.Quit();
            }
        }

        private void createFolderIfNotExist(string path) {
            if (!System.IO.Directory.Exists(path))
                System.IO.Directory.CreateDirectory(path);
        }

        private void createStationInfoSheet() {
            string workSheetDate = monthNum + "/1/" + yearNum;
            wkSheet = wkBook.Worksheets.get_Item(1);
            stationInfoSheet = new StationInfoSheet(wkSheet, workSheetDate, stationNumber, forest, schedule, employee, daysOff);
            stationInfoSheet.buildSheet();
            xlApp.ActiveWindow.DisplayGridlines = false;
            xlApp.ActiveWindow.DisplayHeadings = false;
           
            
        }

        public void createMonthlySummarySheet() {
            wkSheet = wkBook.Worksheets.get_Item(daysInMonth + 2);
            wkSheet.Activate();
            MonthlySummarySheet monthlySummarySheet = new MonthlySummarySheet(wkSheet, monthNum, yearNum, daysInMonth);
            monthlySummarySheet.buildSheet();

        }
        public void createDailyLogWorksheets() {
            Worksheet sourceSheet;
            Worksheet destinationSheet;
            dailyLogWorkSheet = new DailyLogWorksheet();
            wkSheet = wkBook.Worksheets.get_Item(2);
            wkSheet.Activate();

            dailyLogWorkSheet.buildSheet(wkSheet, 1, monthNum, yearNum);
            xlApp.ActiveWindow.DisplayGridlines = false;
            xlApp.ActiveWindow.DisplayHeadings = false;
            sourceSheet = wkSheet;
            
       
            for (int i = 2; i <= daysInMonth; i++) {
                sourceSheet.Activate();
                sourceSheet.get_Range("A1:CU151").Copy(Type.Missing);
                //

                //
                destinationSheet = wkBook.Worksheets.get_Item(i + 1);
                destinationSheet.Activate();
                destinationSheet.get_Range("A1:ZZ400").PasteSpecial(XlPasteType.xlPasteAll, XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
                //destinationSheet.get_Range("A1:ZZ400").PasteSpecial(XlPasteType.xlPasteFormulas, XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
                
                destinationSheet.Name = i.ToString();
                destinationSheet.get_Range("J4").Value = monthNum.ToString() + "/" + i.ToString() + "/" + yearNum.ToString();
                destinationSheet.Columns.EntireColumn.ColumnWidth = 1;
                destinationSheet.Rows.EntireRow.RowHeight = 12;
                destinationSheet.get_Range("AC1").EntireColumn.ColumnWidth = 3;
                //destinationSheet.Names.Add("leaveTypes", Properties.Resources.LeaveType);
                destinationSheet.get_Range("J4").NumberFormat = "dddd, mmmm dd";
                destinationSheet.get_Range("P7").Select();
                destinationSheet.get_Range("A69:A400").EntireRow.Hidden = true;
                destinationSheet.Protect();
                xlApp.ActiveWindow.DisplayGridlines = false;
                xlApp.ActiveWindow.DisplayHeadings = false;
            }
            sourceSheet.Protect();
        }

        public string StationNumber { get { return stationNumber; } set { stationNumber = value; } }
        public string Forest { get { return forest; } set { forest = value; } }
        public string Schedule { get { return schedule; } set {
            if (string.IsNullOrEmpty(value))
                schedule = "";
            else
                schedule = value;
        } }
        public List<Employee> Employees { get { return employee; } set { employee = value; } }
        public Dictionary<string, string> DaysOff { get { return daysOff; } set { daysOff = value; } }


    }
}
