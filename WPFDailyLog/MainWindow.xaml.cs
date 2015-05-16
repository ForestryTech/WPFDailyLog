using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Collections;

namespace WPFDailyLog
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private Dictionary<string, Label> daysOffLabels;
        private List<Forest> forestList;
        private List<string> daysOff;
        private List<Label> crewDayOffLabelList;
        private List<Label> dayOffTotalLabelList;
        private List<TextBox> employeeNames;
        private List<TextBox> employeeNamesFirst;
        private List<ComboBox> dayOff;
        private string[] months;
        private string[] years;

        public MainWindow() {
            InitializeComponent();
            loadLists();
            loadComboBox();
            loadDayOffComboBoxes();
            addDayOffLabels();
        }


        private void loadDayOffComboBoxes() {
            dayOff = new List<ComboBox>() {
                employee1DaysOff, employee2DaysOff, employee3DaysOff, employee4DaysOff, employee5DaysOff, employee6DaysOff, employee7DaysOff,
                employee8DaysOff, employee9DaysOff, employee10DaysOff };
            foreach (ComboBox day in dayOff) {
                foreach (String offDays in daysOff) {
                    day.Items.Add(offDays);
                }
            }
        }
        private void btnTest_Click(object sender, RoutedEventArgs e) {
            
        }

        private void addDayOffLabels() {
            daysOffLabels = new Dictionary<string, Label>();
            crewDayOffLabelList = new List<Label>();
            string labelName = "";
            for (int i = 1; i < 11; i++) { // rows

                for (int j = 3; j < 10; j++) { // cols
                    Label dayOffLabel = new Label();
                    labelName = "emp" + i.ToString();
                    labelName += dayForLabelName(j);
                    //dayOffLabel.Content = "OFF";
                    dayOffLabel.Name = labelName;
                    dayOffLabel.VerticalContentAlignment = VerticalAlignment.Center;
                    dayOffLabel.HorizontalContentAlignment = HorizontalAlignment.Center;
                    dayOffLabel.FontSize = 16;
                    Grid.SetRow(dayOffLabel, i);
                    Grid.SetColumn(dayOffLabel, j);
                    grdCrewInfo.Children.Add(dayOffLabel);
                    daysOffLabels.Add(labelName, dayOffLabel);
                    crewDayOffLabelList.Add(dayOffLabel);
                }
            }
        }

        private void loadComboBox() {
            foreach (Forest forest in forestList) {
                cboForest.Items.Add(forest.ForestName);
            }
            cboForest.Items.Add("");
            cboTour.Items.Add("5/8's");
            cboTour.Items.Add("4/10's");
            cboTour.Items.Add("");

            foreach (String month in months) {
                cboMonth.Items.Add(month);
            }
            foreach (String year in years) {
                cboYear.Items.Add(year);
            }
        }
        
        private string dayForLabelName(int d) {
            string day = "";
            switch (d) {
                case 3:
                    day = "Sunday";
                    break;
                case 4:
                    day = "Monday";
                    break;
                case 5:
                    day = "Tuesday";
                    break;
                case 6:
                    day = "Wednesday";
                    break;
                case 7:
                    day = "Thursday";
                    break;
                case 8:
                    day = "Friday";
                    break;
                case 9:
                    day = "Saturday";
                    break;
                default:
                    break;
            }
            return day;
        }

        private void btnTest2_Click(object sender, RoutedEventArgs e) {
            Label l = daysOffLabels["emp1Sunday"];
            Console.WriteLine("Label: " + l.Name);
        }

        private void loadLists() {
            daysOff = new List<string>() {
                "Sunday-Monday", "Monday-Tuesday", "Tuesday-Wednesday", "Wednesday-Thursday", "Thursday-Friday", "Friday-Saturday",  "Saturday-Sunday",
                "Sunday-Monday-Tuesday", "Monday-Tuesday-Wednesday", "Tuesday-Wednesday-Thursday", "Wednesday-Thursday-Friday", "Thursday-Friday-Saturday",
                "Friday-Saturday-Sunday", "Saturday-Sunday-Monday", ""
            };

            forestList = new List<Forest>() {
                new Forest("Angelus National Forest", "ANF", "0501"),
                new Forest("Cleveland National Forest", "CNF", "0502"),
                new Forest("Eldorado National Forest", "ENF", "0503"),
                new Forest("Inyo National Forest", "INF", "0504"),
                new Forest("Klamath National Forest", "KNF", "0505"),
                new Forest("Lake Tahoe Basin Mgmt Unit", "TMU", "0519"),
                new Forest("Lassen National Forest", "LNF", "0506"),
                new Forest("Los Padres National Forest", "LPF", "0507"),
                new Forest("Mendocino National Forest", "MNF", "0508"),
                new Forest("Modoc National Forest", "MDF", "0509"),
                new Forest("Plumas National Forest", "PNF", "0511"),
                new Forest("R5 Regional Office", "R5RO", "0520"),
                new Forest("San Bernardino National Forest", "BDF", "0512"),
                new Forest("Sequoia National Forest", "SQF", "0513"),
                new Forest("Shasta-Trinity National Forest", "SHF", "0514"),
                new Forest("Sierra National Forest", "SNF", "0515"),
                new Forest("Six Rivers National Forest", "SRF", "0510"),
                new Forest("Stanislaus National Forest", "STF", "0516"),
                new Forest("Tahoe National Forest", "TNF", "0517")
            };

            dayOffTotalLabelList = new List<Label>() {
                sundayTotal, mondayTotal, tuesdayTotal, wednesdayTotal, thursdayTotal, fridayTotal, saturdayTotal
            };

            employeeNames = new List<TextBox>() { employee1, employee2, employee3, employee4, employee5, employee6, employee7, employee8, employee9, employee10 };

            employeeNamesFirst = new List<TextBox>() { employee1First, employee2First, employee3First, employee4First, employee5First,
                employee6First, employee7First, employee8First, employee9First, employee10First };

            months = new string[] { "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December" };

            years = new string[] { "2014", "2015", "2016", "2017", "2018", "2019", "2020", "2021", "2022", "2023", "2024", "2025", "2026", "2027", "2028", "2029", "2030", "2031" };
        }



        private void setDayOffLabels(int row, string days) {
            bool noDaysSelected = false;
            string[] selectedDaysOff = days.Split('-');
            string crewRow = row.ToString();

            if (string.IsNullOrEmpty(days)) noDaysSelected = true;

            foreach (Label label in crewDayOffLabelList) {
                if (crewRow != "1") {
                    if (label.Name.Contains(crewRow)) {
                        if (noDaysSelected) {
                            label.Content = "";
                            label.Background = Brushes.Transparent;
                        } else {
                            if (selectedDaysOff.Count() == 2) {
                                if (label.Name.Contains(selectedDaysOff[0]) || label.Name.Contains(selectedDaysOff[1])) {
                                    label.Content = "OFF";
                                    label.Background = Brushes.LightSalmon;
                                } else {
                                    label.Content = "ON";
                                    label.Background = Brushes.LightBlue;
                                }
                            } else {
                                if (label.Name.Contains(selectedDaysOff[0]) || label.Name.Contains(selectedDaysOff[1]) || label.Name.Contains(selectedDaysOff[2])) {
                                    label.Content = "OFF";
                                    label.Background = Brushes.LightSalmon;
                                } else {
                                    label.Content = "ON";
                                    label.Background = Brushes.LightBlue;
                                }
                            }
                        }
                    }
                } else {
                        if (label.Name.Contains("1") && !(label.Name.Contains("0"))) {
                        if (noDaysSelected) {
                            label.Content = "";
                            label.Background = Brushes.Transparent;
                        } else {
                            if (selectedDaysOff.Count() == 2) {
                                if (label.Name.Contains(selectedDaysOff[0]) || label.Name.Contains(selectedDaysOff[1])) {
                                    label.Content = "OFF";
                                    label.Background = Brushes.LightSalmon;
                                } else {
                                    label.Content = "ON";
                                    label.Background = Brushes.LightBlue;
                                }
                            } else {
                                if (label.Name.Contains(selectedDaysOff[0]) || label.Name.Contains(selectedDaysOff[1]) || label.Name.Contains(selectedDaysOff[2])) {
                                    label.Content = "OFF";
                                    label.Background = Brushes.LightSalmon;
                                } else {
                                    label.Content = "ON";
                                    label.Background = Brushes.LightBlue;
                                }
                            }
                        }
                    }
                }
            }
            calculateTotalEmployeesPerDay();
        }

        private int getDaysOffRow(string cboName) {
            int row = 0;

            switch (cboName) {
                case "employee1DaysOff":
                    row = 1;
                    break;
                case "employee2DaysOff":
                    row = 2;
                    break;
                case "employee3DaysOff":
                    row = 3;
                    break;
                case "employee4DaysOff":
                    row = 4;
                    break;
                case "employee5DaysOff":
                    row = 5;
                    break;
                case "employee6DaysOff":
                    row = 6;
                    break;
                case "employee7DaysOff":
                    row = 7;
                    break;
                case "employee8DaysOff":
                    row = 8;
                    break;
                case "employee9DaysOff":
                    row = 9;
                    break;
                case "employee10DaysOff":
                    row = 10;
                    break;
            }
            return row;
        }
        
        private void employee1DaysOff_DropDownClosed(object sender, EventArgs e) {
            setDayOffLabels(1, employee1DaysOff.Text); 
        }

        private void employee2DaysOff_DropDownClosed(object sender, EventArgs e) {
            setDayOffLabels(2, employee2DaysOff.Text);
        }

        private void employee3DaysOff_DropDownClosed(object sender, EventArgs e) {
            setDayOffLabels(3, employee3DaysOff.Text);
        }

        private void employee4DaysOff_DropDownClosed(object sender, EventArgs e) {
            setDayOffLabels(4, employee4DaysOff.Text);
        }

        private void employee5DaysOff_DropDownClosed(object sender, EventArgs e) {
            setDayOffLabels(5, employee5DaysOff.Text);
        }

        private void employee6DaysOff_DropDownClosed(object sender, EventArgs e) {
            setDayOffLabels(6, employee6DaysOff.Text);
        }

        private void employee7DaysOff_DropDownClosed(object sender, EventArgs e) {
            setDayOffLabels(7, employee7DaysOff.Text);
        }

        private void employee8DaysOff_DropDownClosed(object sender, EventArgs e) {
            setDayOffLabels(8, employee8DaysOff.Text);
        }

        private void employee9DaysOff_DropDownClosed(object sender, EventArgs e) {
            setDayOffLabels(9, employee9DaysOff.Text);
        }

        private void employee10DaysOff_DropDownClosed(object sender, EventArgs e) {
            setDayOffLabels(10, employee10DaysOff.Text);
        }

        private void calculateTotalEmployeesPerDay() {
            int[] total = new int[7];
            string labelContent;
            foreach (Label label in crewDayOffLabelList) {
                labelContent = (string)label.Content;
                if (label.Name.Contains("Sunday") && labelContent == "ON")
                    total[0]++;
                if (label.Name.Contains("Monday") && labelContent == "ON")
                    total[1]++;
                if (label.Name.Contains("Tuesday") && labelContent == "ON")
                    total[2]++;
                if (label.Name.Contains("Wednesday") && labelContent == "ON")
                    total[3]++;
                if (label.Name.Contains("Thursday") && labelContent == "ON")
                    total[4]++;
                if (label.Name.Contains("Friday") && labelContent == "ON")
                    total[5]++;
                if (label.Name.Contains("Saturday") && labelContent == "ON")
                    total[6]++;
            }
            int ctr = 0;
            foreach (Label label in dayOffTotalLabelList) {
                label.Content = total[ctr].ToString();
                ctr++;
            }

        }

        private void clearTotalLabels() {
            foreach (Label label in dayOffTotalLabelList) {
                label.Content = "";
            }
        }

        private void clearDayOffLabels() {
            foreach (Label label in crewDayOffLabelList) {
                label.Content = "";
            }
        }

        private void btnGenerateWorkbook_Click(object sender, RoutedEventArgs e) {
            DailyLogBook dailyLogBook = new DailyLogBook(cboMonth.Text, cboYear.Text);
            dailyLogBook.Forest = cboForest.Text;
            dailyLogBook.StationNumber = txtStation.Text;
            dailyLogBook.Employees = getEmployeesIntoList();
            dailyLogBook.DaysOff = getEmployeeDaysOffIntoList();
            dailyLogBook.Schedule = cboTour.Text;
            dailyLogBook.CreateWorkBook();
        }

        private List<Employee> getEmployeesIntoList() {
            List<Employee> tmpEmployeeList = new List<Employee>();
            for (int i = 0; i < employeeNames.Count; i++) {
                if (string.IsNullOrEmpty(employeeNames[i].Text))
                    tmpEmployeeList.Add(new Employee("EMPTY", "EMPTY"));
                else
                    tmpEmployeeList.Add(new Employee(employeeNames[i].Text, employeeNamesFirst[i].Text));
            }
            return tmpEmployeeList;
        }

        private Dictionary<string, string> getEmployeeDaysOffIntoList() {
            try {
                string tmpEmployee;
                Dictionary<string, string> tmpDayOff = new Dictionary<string, string>();
                string tmpDay;
                for (int i = 0; i < dayOff.Count; i++) {
                    if (string.IsNullOrEmpty(employeeNames[i].Text)) {
                        tmpEmployee = "EMPTY" + i.ToString();
                        tmpDay = "EMPTY" + i.ToString();
                        tmpDayOff.Add(tmpEmployee, tmpDay);
                    } else {
                        tmpEmployee = employeeNames[i].Text + ", " + employeeNamesFirst[i].Text;
                        tmpDay = dayOff[i].Text;
                        tmpDayOff.Add(tmpEmployee, tmpDay);
                    }
                }
                return tmpDayOff;
            }
            catch (Exception e) {
                
                throw;
            }
        }

 



        private void btnViewResponse_Click(object sender, RoutedEventArgs e) {
            string newRange;
            string col;
            string row = "12";
            char c;
            for (int i = 1; i < 35; i++) {
                if ((i+64) > 90) {
                    c = (char)((i - 26) + 64);
                    col = "A" + c.ToString();
                } else {
                    c = (char)(i + 64);
                    col = c.ToString();
                }
                newRange = col + row;
                Console.WriteLine("Range: " + newRange);

            }
        }

        private void btnEnterResponse_Click(object sender, RoutedEventArgs e) {
            string path = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            path += @"\Daily_Log_2015";
            
            if (!(System.IO.Directory.Exists(path))) {
                Console.WriteLine(path + " - does not exist.");
                System.IO.Directory.CreateDirectory(path);
            } else
                Console.WriteLine(path + " - exists");
            
        }

        private void employeeDaysOff_KeyUp(object sender, KeyEventArgs e) {
            ComboBox cbo = sender as ComboBox;

            setDayOffLabels(getDaysOffRow(cbo.Name), cbo.Text);
        }

        private void employeeDaysOff_DropDownClosed(object sender, EventArgs e) {
            ComboBox cbo = sender as ComboBox;

            setDayOffLabels(getDaysOffRow(cbo.Name), cbo.Text);
        }

        private void btnYearlyReport_Click(object sender, RoutedEventArgs e) {
            
        }




    }
}
