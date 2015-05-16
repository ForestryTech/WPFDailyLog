using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace WPFDailyLog
{
    partial class StationInfoSheet
    {
        private List<DailyLogCells> headerCells;
        private List<DailyLogCells> crewInfoCells;
        private List<string> daysOffList;

        private void fillStationInfoCells() {
            headerCells = new List<DailyLogCells>() {
                new DailyLogCells("E1", date, new CellText(FontPosition.Center, 18), CellBorder.No),
                new DailyLogCells("B2:D2", "Station", new CellText(FontStyle.Bold, FontPosition.Right, 16), CellBorder.No),
                new DailyLogCells("E2", station, new CellText(FontStyle.Bold, FontPosition.Center, 16), CellBorder.No),
                new DailyLogCells("F2:H2", "Forest", new CellText(FontStyle.Bold, FontPosition.Center, 16), CellBorder.No),
                new DailyLogCells("I2:M2", forest, new CellText(FontStyle.Bold, FontPosition.Center, 16), CellBorder.No),
                new DailyLogCells("B4:D4", "Employee Name", new CellText(FontStyle.Bold, FontPosition.Center, 11), CellBorder.No),
                new DailyLogCells("E4", "Days Off", new CellText(FontStyle.Bold, FontPosition.Center, 11), CellBorder.No),
                new DailyLogCells("G4", "Sunday", new CellText(FontStyle.Bold, FontPosition.Center, 11), CellBorder.No),
                new DailyLogCells("H4", "Monday", new CellText(FontStyle.Bold, FontPosition.Center, 11), CellBorder.No),
                new DailyLogCells("I4", "Tuesday", new CellText(FontStyle.Bold, FontPosition.Center, 11), CellBorder.No),
                new DailyLogCells("J4", "Wednesday", new CellText(FontStyle.Bold, FontPosition.Center, 11), CellBorder.No),
                new DailyLogCells("K4", "Thursday", new CellText(FontStyle.Bold, FontPosition.Center, 11), CellBorder.No),
                new DailyLogCells("L4", "Friday", new CellText(FontStyle.Bold, FontPosition.Center, 11), CellBorder.No),
                new DailyLogCells("M4", "Saturday", new CellText(FontStyle.Bold, FontPosition.Center, 11), CellBorder.No),
                new DailyLogCells("E17", "Number of employees", new CellText(FontStyle.Bold, FontPosition.Center,11), CellBorder.No),
                new DailyLogCells("G17", "COUNTIF(G5:G14,\"ON\")", true, new CellText(FontStyle.Bold, FontPosition.Center), CellBorder.No),
                new DailyLogCells("H17", "COUNTIF(H5:H14,\"ON\")", true, new CellText(FontStyle.Bold, FontPosition.Center), CellBorder.No),
                new DailyLogCells("I17", "COUNTIF(I5:I14,\"ON\")", true, new CellText(FontStyle.Bold, FontPosition.Center), CellBorder.No),
                new DailyLogCells("J17", "COUNTIF(J5:J14,\"ON\")", true, new CellText(FontStyle.Bold, FontPosition.Center), CellBorder.No),
                new DailyLogCells("K17", "COUNTIF(K5:K14,\"ON\")", true, new CellText(FontStyle.Bold, FontPosition.Center), CellBorder.No),
                new DailyLogCells("L17", "COUNTIF(L5:L14,\"ON\")", true, new CellText(FontStyle.Bold, FontPosition.Center), CellBorder.No),
                new DailyLogCells("M17", "COUNTIF(M5:M14,\"ON\")", true, new CellText(FontStyle.Bold, FontPosition.Center), CellBorder.No),
                new DailyLogCells("H18:J18", "Schedule", new CellText(FontStyle.Bold, FontPosition.Center), CellBorder.No),
                new DailyLogCells("K18:M18", schedule, new CellText(FontStyle.Bold, FontPosition.Center), CellBorder.No)
            };

            crewInfoCells = new List<DailyLogCells>() {
                new DailyLogCells("A5", "Captain", new CellText(FontPosition.Left), CellBorder.Yes),
                new DailyLogCells("A6", "Engineer", new CellText(FontPosition.Left), CellBorder.Yes),
                new DailyLogCells("A7", "AFEO", new CellText(FontPosition.Left), CellBorder.Yes),
                new DailyLogCells("A8", "Senior FF", new CellText(FontPosition.Left), CellBorder.Yes),
                new DailyLogCells("A9", "Firefighter", new CellText(FontPosition.Left), CellBorder.Yes),
                new DailyLogCells("A10", "Firefighter", new CellText(FontPosition.Left), CellBorder.Yes),
                new DailyLogCells("A11", "Firefighter", new CellText(FontPosition.Left), CellBorder.Yes),
                new DailyLogCells("A12", "WT/AFEO", new CellText(FontPosition.Left), CellBorder.Yes),
                new DailyLogCells("A13", "Firefighter", new CellText(FontPosition.Left), CellBorder.Yes),
                new DailyLogCells("A14", "Firefighter", new CellText(FontPosition.Left), CellBorder.Yes),

                new DailyLogCells("B5:D5", "", new CellText(FontPosition.Center), CellBorder.Yes),
                new DailyLogCells("B6:D6", "", new CellText(FontPosition.Center), CellBorder.Yes),
                new DailyLogCells("B7:D7", "", new CellText(FontPosition.Center), CellBorder.Yes),
                new DailyLogCells("B8:D8", "", new CellText(FontPosition.Center), CellBorder.Yes),
                new DailyLogCells("B9:D9", "", new CellText(FontPosition.Center), CellBorder.Yes),
                new DailyLogCells("B10:D10", "", new CellText(FontPosition.Center), CellBorder.Yes),
                new DailyLogCells("B11:D11", "", new CellText(FontPosition.Center), CellBorder.Yes),
                new DailyLogCells("B12:D12", "", new CellText(FontPosition.Center), CellBorder.Yes),
                new DailyLogCells("B13:D13", "", new CellText(FontPosition.Center), CellBorder.Yes),
                new DailyLogCells("B14:D14", "", new CellText(FontPosition.Center), CellBorder.Yes),

                new DailyLogCells("E5", new CellText(FontPosition.Left)),
                new DailyLogCells("E6", new CellText(FontPosition.Left)),
                new DailyLogCells("E7", new CellText(FontPosition.Left)),
                new DailyLogCells("E8", new CellText(FontPosition.Left)),
                new DailyLogCells("E9", new CellText(FontPosition.Left)),
                new DailyLogCells("E10", new CellText(FontPosition.Left)),
                new DailyLogCells("E11", new CellText(FontPosition.Left)),
                new DailyLogCells("E12", new CellText(FontPosition.Left)),
                new DailyLogCells("E13", new CellText(FontPosition.Left)),
                new DailyLogCells("E14", new CellText(FontPosition.Left)),

                new DailyLogCells("F5", true),
                new DailyLogCells("F6", true),
                new DailyLogCells("F7", true),
                new DailyLogCells("F8", true),
                new DailyLogCells("F9", true),
                new DailyLogCells("F10", true),
                new DailyLogCells("F11", true),
                new DailyLogCells("F12", true),
                new DailyLogCells("F13", true),
                new DailyLogCells("F14", true),

            };

            for (int i = 5; i < 15; i++) {
                //Sunday
                crewInfoCells.Add(new DailyLogCells("G" + i.ToString(), "IF($E" + i.ToString() + "=\"\",\"\",IF(ISERROR(SEARCH(G$4,$E"+i.ToString()+")),\"ON\",\"OFF\"))", true, new CellText(FontPosition.Center), CellBorder.Yes));
                //Monday
                crewInfoCells.Add(new DailyLogCells("H" + i.ToString(), "IF($E" + i.ToString() + "=\"\",\"\",IF(ISERROR(SEARCH(H$4,$E" + i.ToString() + ")),\"ON\",\"OFF\"))", true, new CellText(FontPosition.Center), CellBorder.Yes));
                //Tuesday
                crewInfoCells.Add(new DailyLogCells("I" + i.ToString(), "IF($E" + i.ToString() + "=\"\",\"\",IF(ISERROR(SEARCH(I$4,$E" + i.ToString() + ")),\"ON\",\"OFF\"))", true, new CellText(FontPosition.Center), CellBorder.Yes));
                //Wednesday
                crewInfoCells.Add(new DailyLogCells("J" + i.ToString(), "IF($E" + i.ToString() + "=\"\",\"\",IF(ISERROR(SEARCH(J$4,$E" + i.ToString() + ")),\"ON\",\"OFF\"))", true, new CellText(FontPosition.Center), CellBorder.Yes));
                //Thursday
                crewInfoCells.Add(new DailyLogCells("K" + i.ToString(), "IF($E" + i.ToString() + "=\"\",\"\",IF(ISERROR(SEARCH(K$4,$E" + i.ToString() + ")),\"ON\",\"OFF\"))", true, new CellText(FontPosition.Center), CellBorder.Yes));
                //Friday
                crewInfoCells.Add(new DailyLogCells("L" + i.ToString(), "IF($E" + i.ToString() + "=\"\",\"\",IF(ISERROR(SEARCH(L$4,$E" + i.ToString() + ")),\"ON\",\"OFF\"))", true, new CellText(FontPosition.Center), CellBorder.Yes));
                //Saturday
                crewInfoCells.Add(new DailyLogCells("M" + i.ToString(), "IF($E" + i.ToString() + "=\"\",\"\",IF(ISERROR(SEARCH(M$4,$E" + i.ToString() + ")),\"ON\",\"OFF\"))", true, new CellText(FontPosition.Center), CellBorder.Yes));
            }

            daysOffList = new List<string>() {
                "Sunday-Monday", "Monday-Tuesday", "Tuesday-Wednesday", "Wednesday-Thursday", "Thursday-Friday", "Friday-Saturday",  "Saturday-Sunday",
                "Sunday-Monday-Tuesday", "Monday-Tuesday-Wednesday", "Tuesday-Wednesday-Thursday", "Wednesday-Thursday-Friday", "Thursday-Friday-Saturday",
                "Friday-Saturday-Sunday", "Saturday-Sunday-Monday", ""
            };
        }
    }
}
//Sunday
/* new DailyLogCells("G5","=IF($E5=\"\",\"\",IF(ISERROR(SEARCH(G$4,$E5)),\"ON\",\"OFF\"))", true, new CellText(FontPosition.Center), CellBorder.Yes),
new DailyLogCells("G6","=IF($E6=\"\",\"\",IF(ISERROR(SEARCH(G$4,$E6)),\"ON\",\"OFF\"))", true, new CellText(FontPosition.Center), CellBorder.Yes),
new DailyLogCells("G7","=IF($E7=\"\",\"\",IF(ISERROR(SEARCH(G$4,$E7)),\"ON\",\"OFF\"))", true, new CellText(FontPosition.Center), CellBorder.Yes),
new DailyLogCells("G8","=IF($E8=\"\",\"\",IF(ISERROR(SEARCH(G$4,$E8)),\"ON\",\"OFF\"))", true, new CellText(FontPosition.Center), CellBorder.Yes),
new DailyLogCells("G9","=IF($E9=\"\",\"\",IF(ISERROR(SEARCH(G$4,$E9)),\"ON\",\"OFF\"))", true, new CellText(FontPosition.Center), CellBorder.Yes),
new DailyLogCells("G10","=IF($E10=\"\",\"\",IF(ISERROR(SEARCH(G$4,$E10)),\"ON\",\"OFF\"))", true, new CellText(FontPosition.Center), CellBorder.Yes),
new DailyLogCells("G11","=IF($E11=\"\",\"\",IF(ISERROR(SEARCH(G$4,$E11)),\"ON\",\"OFF\"))", true, new CellText(FontPosition.Center), CellBorder.Yes),
new DailyLogCells("G12","=IF($E12=\"\",\"\",IF(ISERROR(SEARCH(G$4,$E12)),\"ON\",\"OFF\"))", true, new CellText(FontPosition.Center), CellBorder.Yes),
new DailyLogCells("G13","=IF($E13=\"\",\"\",IF(ISERROR(SEARCH(G$4,$E13)),\"ON\",\"OFF\"))", true, new CellText(FontPosition.Center), CellBorder.Yes),
new DailyLogCells("G14","=IF($E14=\"\",\"\",IF(ISERROR(SEARCH(G$4,$E14)),\"ON\",\"OFF\"))", true, new CellText(FontPosition.Center), CellBorder.Yes),
//Monday
new DailyLogCells("H5","=IF($E5=\"\",\"\",IF(ISERROR(SEARCH(H$4,$E5)),\"ON\",\"OFF\"))", true, new CellText(FontPosition.Center), CellBorder.Yes),
new DailyLogCells("H6","=IF($E6=\"\",\"\",IF(ISERROR(SEARCH(H$4,$E6)),\"ON\",\"OFF\"))", true, new CellText(FontPosition.Center), CellBorder.Yes),
new DailyLogCells("H7","=IF($E7=\"\",\"\",IF(ISERROR(SEARCH(H$4,$E5)),\"ON\",\"OFF\"))", true, new CellText(FontPosition.Center), CellBorder.Yes),
new DailyLogCells("H8","=IF($E8=\"\",\"\",IF(ISERROR(SEARCH(H$4,$E5)),\"ON\",\"OFF\"))", true, new CellText(FontPosition.Center), CellBorder.Yes),
new DailyLogCells("H9","=IF($E9=\"\",\"\",IF(ISERROR(SEARCH(H$4,$E5)),\"ON\",\"OFF\"))", true, new CellText(FontPosition.Center), CellBorder.Yes),
new DailyLogCells("H10","=IF($E10=\"\",\"\",IF(ISERROR(SEARCH(H$4,$E5)),\"ON\",\"OFF\"))", true, new CellText(FontPosition.Center), CellBorder.Yes),
new DailyLogCells("H11","=IF($E11=\"\",\"\",IF(ISERROR(SEARCH(H$4,$E5)),\"ON\",\"OFF\"))", true, new CellText(FontPosition.Center), CellBorder.Yes),
new DailyLogCells("H12","=IF($E12=\"\",\"\",IF(ISERROR(SEARCH(H$4,$E5)),\"ON\",\"OFF\"))", true, new CellText(FontPosition.Center), CellBorder.Yes),
new DailyLogCells("H13","=IF($E13=\"\",\"\",IF(ISERROR(SEARCH(H$4,$E5)),\"ON\",\"OFF\"))", true, new CellText(FontPosition.Center), CellBorder.Yes),
new DailyLogCells("H14","=IF($E14=\"\",\"\",IF(ISERROR(SEARCH(H$4,$E5)),\"ON\",\"OFF\"))", true, new CellText(FontPosition.Center), CellBorder.Yes),  */