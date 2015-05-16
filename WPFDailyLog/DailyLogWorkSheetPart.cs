using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WPFDailyLog
{
    partial class DailyLogWorksheet
    {
        private List<DailyLogCells> dailyLogCells;
        //private List<DailyLogCells> nonBorderDailyLogCells;
        //private List<DailyLogCells> workSheetHeaders;
        //private List<DailyLogCells> crewNamesAndStatus;
        //private List<DailyLogCells> activityCellsA;
        //private List<DailyLogCells> activityCellsB;
        //private List<DailyLogCells> dailyActivityHeader;
        //private List<DailyLogCells> dailyActivityCells;
        private List<Forest> forestList;
        //private List<String> activityList;
        private List<String> Status;
        private List<String> activityTypes;

        private void fillDailyLogCells() {
            Status = new List<string>() {
                "On Duty",
                "Off Duty",
                "Cover",
                "A/L",
                "S/L",
                "Holiday",
                "LWOP",
                "AWOL",
                "6th/7th Day",
                "Holiday Worked",
                "Admin",
                "Military",
                "Severity"
            };

            activityTypes = new List<String>() {
                "Station Administration", "Station Maintenance", "Engine PM/Readiness", "Physical Training", "Training",
                "Wildland Fire Base hours", "Wildland Fire OT", "Overhead Assignments", "Traffic Collision", "Medical Aid",
                "Hazmat", "Structure Fire", "Vehicle Fire", "Stand-by / Cover", "LE Assist / S&R", "False Alarm / Cancel",
                "Extended Staffing", "Other Emergency", "Non fire OT", "Pre-suppresion", "Fuels / RX Burning", "Safety",
                "Recruitment / Outreach", "Hiring", "Recreation Assist", "Prevention / Patrol", "Engineering assist",
                "Public assist / Media", "Resource assist", "WCT(Work Capacity Test", "Physicals / Drug test", "Local / Forest cadre",
                "Regional cadre", "Regional committee", "Meeting / Briefing", "Fire rehab / Admin", "Facilities Maintenance"
            };

            dailyLogCells = new List<DailyLogCells>() {
                new DailyLogCells("A2:BT3", "VLOOKUP('Station Info'!I2,AC72:AE90,2) & \" Station \" & 'Station Info'!E2", true, new CellText(FontStyle.Bold,12), CellBorder.No),
                new DailyLogCells("BU2:CL3", "Pay Period", new CellText(FontPosition.Left, 16), CellBorder.No),
                new DailyLogCells("CM2:CT3", "VLOOKUP(J4,U73:W150,3)", true, new CellText(FontStyle.Bold, 16), CellBorder.No),
                new DailyLogCells("A4:I4", "Date:", new CellText(FontStyle.Bold, FontPosition.Right, 10), CellBorder.No),
                new DailyLogCells("J4:AI4", sheetDate, new CellText(FontPosition.Left, 10), CellBorder.No),
                new DailyLogCells("AW4:BO4", "Duty Hours:", new CellText(FontStyle.Bold), CellBorder.No),
                new DailyLogCells("BP4:BT4", "IF('Station Info'!K18=\"5/8's\",8,IF('Station Info'!K18=\"4/10's\",10,\"\"))", true, CellBorder.No),
                new DailyLogCells("A6:O6", "Crew",new CellText(FontStyle.Bold)),
                new DailyLogCells("P6:AA6", "Status", new CellText(FontStyle.Bold)),
                new DailyLogCells("AB6:AD6", "Hours", new CellText(FontStyle.Bold, 10)),
                new DailyLogCells("AE6:AT6", "Comments", new CellText(FontStyle.Bold)),
                new DailyLogCells("AU6:BL6", "Activity", new CellText(FontStyle.Bold)),
                new DailyLogCells("BM6:BT6", "Hours", new CellText(FontStyle.Bold)),
                new DailyLogCells("BU6:CL6", "Activity", new CellText(FontStyle.Bold)),
                new DailyLogCells("CM6:CT6", "Hours", new CellText(FontStyle.Bold)),
                // Crew Cells                   // Stauts Cells                  // Hours Cell                   // Comments Cell
                new DailyLogCells("A7:O8", "IF('Station Info'!B5=\"\",\"\",'Station Info'!B5)", true),     new DailyLogCells("P7:AA8"),    new DailyLogCells("AB7:AD8"),   new DailyLogCells("AE7:AT8"),
                new DailyLogCells("A9:O10", "IF('Station Info'!B6=\"\",\"\",'Station Info'!B6)", true),    new DailyLogCells("P9:AA10"),   new DailyLogCells("AB9:AD10"),  new DailyLogCells("AE9:AT10"),
                new DailyLogCells("A11:O12", "IF('Station Info'!B7=\"\",\"\",'Station Info'!B7)", true),   new DailyLogCells("P11:AA12"),  new DailyLogCells("AB11:AD12"), new DailyLogCells("AE11:AT12"),
                new DailyLogCells("A13:O14", "IF('Station Info'!B8=\"\",\"\",'Station Info'!B8)", true),   new DailyLogCells("P13:AA14"),  new DailyLogCells("AB13:AD14"), new DailyLogCells("AE13:AT14"),
                new DailyLogCells("A15:O16", "IF('Station Info'!B9=\"\",\"\",'Station Info'!B9)", true),   new DailyLogCells("P15:AA16"),  new DailyLogCells("AB15:AD16"), new DailyLogCells("AE15:AT16"),
                new DailyLogCells("A17:O18", "IF('Station Info'!B10=\"\",\"\",'Station Info'!B10)", true),   new DailyLogCells("P17:AA18"),  new DailyLogCells("AB17:AD18"), new DailyLogCells("AE17:AT18"),
                new DailyLogCells("A19:O20", "IF('Station Info'!B11=\"\",\"\",'Station Info'!B11)", true),   new DailyLogCells("P19:AA20"),  new DailyLogCells("AB19:AD20"), new DailyLogCells("AE19:AT20"),
                new DailyLogCells("A21:O22", "IF('Station Info'!B12=\"\",\"\",'Station Info'!B12)", true),   new DailyLogCells("P21:AA22"),  new DailyLogCells("AB21:AD22"), new DailyLogCells("AE21:AT22"),
                new DailyLogCells("A23:O24", "IF('Station Info'!B13=\"\",\"\",'Station Info'!B13)", true),   new DailyLogCells("P23:AA24"),  new DailyLogCells("AB23:AD24"), new DailyLogCells("AE23:AT24"),
                new DailyLogCells("A25:O26", "IF('Station Info'!B14=\"\",\"\",'Station Info'!B14)", true),   new DailyLogCells("P25:AA26"),  new DailyLogCells("AB25:AD26"), new DailyLogCells("AE25:AT26"),
                // Activity cells                                                                                                            //Hours
                new DailyLogCells("AU7:BL7", Properties.Resources.StationAdministration, new CellText(FontPosition.Left)),                  new DailyLogCells("BM7:BT7","SUMIF($AE$28:$AT$64,AU7,$Z$28:$AD$64)", true),
                new DailyLogCells("AU8:BL8", Properties.Resources.StationMaintenance, new CellText(FontPosition.Left)),                     new DailyLogCells("BM8:BT8", "SUMIF($AE$28:$AT$64,AU8,$Z$28:$AD$64)", true),
                new DailyLogCells("AU9:BL9", Properties.Resources.EnginePMReadiness, new CellText(FontPosition.Left)),                      new DailyLogCells("BM9:BT9", "SUMIF($AE$28:$AT$64,AU9,$Z$28:$AD$64)", true),
                new DailyLogCells("AU10:BL10", Properties.Resources.PhysicalTraining, new CellText(FontPosition.Left)),                     new DailyLogCells("BM10:BT10", "SUMIF($AE$28:$AT$64,AU10,$Z$28:$AD$64)", true),
                new DailyLogCells("AU11:BL11", Properties.Resources.Training, new CellText(FontPosition.Left)),                             new DailyLogCells("BM11:BT11", "SUMIF($AE$28:$AT$64,AU11,$Z$28:$AD$64)", true),
                new DailyLogCells("AU12:BL12", Properties.Resources.WildlandFireBase, new CellText(FontPosition.Left, FontColor.Red)),      new DailyLogCells("BM12:BT12", "SUMIF($AE$28:$AT$64,AU12,$Z$28:$AD$64)", true),
                new DailyLogCells("AU13:BL13", Properties.Resources.WildlandFireOT, new CellText(FontPosition.Left, FontColor.Red)),        new DailyLogCells("BM13:BT13","SUMIF($AE$28:$AT$64,AU13,$Z$28:$AD$64)", true),
                new DailyLogCells("AU14:BL14", Properties.Resources.OverheadAssignments, new CellText(FontPosition.Left, FontColor.Red)),   new DailyLogCells("BM14:BT14", "SUMIF($AE$28:$AT$64,AU14,$Z$28:$AD$64)", true),
                new DailyLogCells("AU15:BL15", Properties.Resources.TrafficCollision, new CellText(FontPosition.Left, FontColor.Red)),      new DailyLogCells("BM15:BT15", "SUMIF($AE$28:$AT$64,AU15,$Z$28:$AD$64)", true),
                new DailyLogCells("AU16:BL16", Properties.Resources.MedicalAid, new CellText(FontPosition.Left, FontColor.Red)),            new DailyLogCells("BM16:BT16", "SUMIF($AE$28:$AT$64,AU16,$Z$28:$AD$64)", true),
                new DailyLogCells("AU17:BL17", Properties.Resources.Hazmat, new CellText(FontPosition.Left, FontColor.Red)),                new DailyLogCells("BM17:BT17", "SUMIF($AE$28:$AT$64,AU17,$Z$28:$AD$64)", true),
                new DailyLogCells("AU18:BL18", Properties.Resources.StructureFire, new CellText(FontPosition.Left, FontColor.Red)),         new DailyLogCells("BM18:BT18","SUMIF($AE$28:$AT$64,AU18,$Z$28:$AD$64)", true),
                new DailyLogCells("AU19:BL19", Properties.Resources.VehicleFire, new CellText(FontPosition.Left, FontColor.Red)),           new DailyLogCells("BM19:BT19", "SUMIF($AE$28:$AT$64,AU19,$Z$28:$AD$64)", true),
                new DailyLogCells("AU20:BL20", Properties.Resources.StandByCover, new CellText(FontPosition.Left, FontColor.Red)),          new DailyLogCells("BM20:BT20", "SUMIF($AE$28:$AT$64,AU20,$Z$28:$AD$64)", true),
                new DailyLogCells("AU21:BL21", Properties.Resources.LEAssistSR, new CellText(FontPosition.Left, FontColor.Red)),            new DailyLogCells("BM21:BT21", "SUMIF($AE$28:$AT$64,AU21,$Z$28:$AD$64)", true),
                new DailyLogCells("AU22:BL22", Properties.Resources.FalseAlarmCancel, new CellText(FontPosition.Left, FontColor.Red)),      new DailyLogCells("BM22:BT22", "SUMIF($AE$28:$AT$64,AU22,$Z$28:$AD$64)", true),
                new DailyLogCells("AU23:BL23", Properties.Resources.ExtendedStaffing, new CellText(FontPosition.Left)),                     new DailyLogCells("BM23:BT23", "SUMIF($AE$28:$AT$64,AU23,$Z$28:$AD$64)", true), 
                new DailyLogCells("AU24:BL24", Properties.Resources.OtherEmergency, new CellText(FontPosition.Left, FontColor.Red)),        new DailyLogCells("BM24:BT24", "SUMIF($AE$28:$AT$64,AU24,$Z$28:$AD$64)", true),
                new DailyLogCells("AU25:BL25", Properties.Resources.NonFireOT, new CellText(FontPosition.Left)),                            new DailyLogCells("BM25:BT25", "SUMIF($AE$28:$AT$64,AU25,$Z$28:$AD$64)", true),
                new DailyLogCells("AU26:BL26", Properties.Resources.PreSuppression, new CellText(FontPosition.Left)),                       new DailyLogCells("BM26:BT26", "SUMIF($AE$28:$AT$64,AU26,$Z$28:$AD$64)", true),

                new DailyLogCells("BU7:CL7", Properties.Resources.FuelsRXBurning, new CellText(FontPosition.Left)),                         new DailyLogCells("CM7:CT7", "SUMIF($AE$28:$AT$64,BU7,$Z$28:$AD$64)", true),
                new DailyLogCells("BU8:CL8", Properties.Resources.Safety, new CellText(FontPosition.Left)),                                 new DailyLogCells("CM8:CT8", "SUMIF($AE$28:$AT$64,BU8,$Z$28:$AD$64)", true),
                new DailyLogCells("BU9:CL9", Properties.Resources.RecruitmentOutreach, new CellText(FontPosition.Left)),                    new DailyLogCells("CM9:CT9", "SUMIF($AE$28:$AT$64,BU9,$Z$28:$AD$64)", true),
                new DailyLogCells("BU10:CL10", Properties.Resources.Hiring, new CellText(FontPosition.Left)),                               new DailyLogCells("CM10:CT10", "SUMIF($AE$28:$AT$64,BU10,$Z$28:$AD$64)", true),
                new DailyLogCells("BU11:CL11", Properties.Resources.RecreationAssist, new CellText(FontPosition.Left)),                     new DailyLogCells("CM11:CT11", "SUMIF($AE$28:$AT$64,BU11,$Z$28:$AD$64)", true),
                new DailyLogCells("BU12:CL12", Properties.Resources.PreventionPatrol, new CellText(FontPosition.Left)),                     new DailyLogCells("CM12:CT12", "SUMIF($AE$28:$AT$64,BU12,$Z$28:$AD$64)", true),
                new DailyLogCells("BU13:CL13", Properties.Resources.EngineeringAssist, new CellText(FontPosition.Left)),                    new DailyLogCells("CM13:CT13", "SUMIF($AE$28:$AT$64,BU13,$Z$28:$AD$64)", true),
                new DailyLogCells("BU14:CL14", Properties.Resources.PublicMedia, new CellText(FontPosition.Left)),                          new DailyLogCells("CM14:CT14", "SUMIF($AE$28:$AT$64,BU14,$Z$28:$AD$64)", true),
                new DailyLogCells("BU15:CL15", Properties.Resources.ResourceAssist, new CellText(FontPosition.Left)),                       new DailyLogCells("CM15:CT15", "SUMIF($AE$28:$AT$64,BU15,$Z$28:$AD$64)", true),
                new DailyLogCells("BU16:CL16", Properties.Resources.WCT, new CellText(FontPosition.Left)),                                  new DailyLogCells("CM16:CT16", "SUMIF($AE$28:$AT$64,BU16,$Z$28:$AD$64)", true),
                new DailyLogCells("BU17:CL17", Properties.Resources.PhysicalsDrugTest, new CellText(FontPosition.Left)),                    new DailyLogCells("CM17:CT17", "SUMIF($AE$28:$AT$64,BU17,$Z$28:$AD$64)", true),
                new DailyLogCells("BU18:CL18", Properties.Resources.LocalForestCadre, new CellText(FontPosition.Left)),                     new DailyLogCells("CM18:CT18", "SUMIF($AE$28:$AT$64,BU18,$Z$28:$AD$64)", true),
                new DailyLogCells("BU19:CL19", Properties.Resources.RegionalCadre, new CellText(FontPosition.Left)),                        new DailyLogCells("CM19:CT19", "SUMIF($AE$28:$AT$64,BU19,$Z$28:$AD$64)", true),
                new DailyLogCells("BU20:CL20", Properties.Resources.RegionalCommittee, new CellText(FontPosition.Left)),                    new DailyLogCells("CM20:CT20", "SUMIF($AE$28:$AT$64,BU20,$Z$28:$AD$64)", true),
                new DailyLogCells("BU21:CL21", Properties.Resources.MeetingBriefing, new CellText(FontPosition.Left)),                      new DailyLogCells("CM21:CT21", "SUMIF($AE$28:$AT$64,BU21,$Z$28:$AD$64)", true),
                new DailyLogCells("BU22:CL22", Properties.Resources.OtherHours, new CellText(FontPosition.Left)),                           new DailyLogCells("CM22:CT22", Properties.Resources.OtherHoursFormula, true),
                new DailyLogCells("BU23:CL23", Properties.Resources.FireRehabAdmin, new CellText(FontPosition.Left)),                       new DailyLogCells("CM23:CT23", "SUMIF($AE$28:$AT$64,BU23,$Z$28:$AD$64)", true),
                new DailyLogCells("BU24:CL24", Properties.Resources.Severity, new CellText(FontPosition.Left, FontColor.Green)),            new DailyLogCells("CM24:CT24", Properties.Resources.SeverityFormula, true),
                new DailyLogCells("BU25:CL25", "Facilities Maintenance", new CellText(FontPosition.Left)), new DailyLogCells("CM25:CT25", "SUMIF($AE$28:$AT$64,BU25,$Z$28:$AD$64)", true),
                new DailyLogCells("BU26:CL26", "Total Hours", new CellText(FontStyle.Bold, FontPosition.Left)), new DailyLogCells("CM26:CT26", "SUM(BM7:BT26)+SUM(CM7:CT25)", true),
                new DailyLogCells("A27:O27", "Time", new CellText(FontStyle.Bold)),
                new DailyLogCells("P27:T27", "# People", new CellText(FontStyle.Bold)),
                new DailyLogCells("U27:Y27", "# Hours", new CellText(FontStyle.Bold)),
                new DailyLogCells("Z27:AD27", "Total Hrs", new CellText(FontStyle.Bold)),
                new DailyLogCells("AE27:AT27", "Activity", new CellText(FontStyle.Bold)),
                new DailyLogCells("AU27:CT27", "Description", new CellText(FontStyle.Bold)),
                new DailyLogCells("A28:O29", new CellText(FontPosition.Left)), new DailyLogCells("P28:T29"), new DailyLogCells("U28:Y29"), new DailyLogCells("Z28:AD29","SUM(P28*U28)", true), new DailyLogCells("AE28:AT29", new CellText(FontPosition.Left)), new DailyLogCells("AU28:CT29", new CellText(FontPosition.Left)),
                new DailyLogCells("A30:O31", new CellText(FontPosition.Left)), new DailyLogCells("P30:T31"), new DailyLogCells("U30:Y31"), new DailyLogCells("Z30:AD31","SUM(P30*U30)", true), new DailyLogCells("AE30:AT31", new CellText(FontPosition.Left)), new DailyLogCells("AU30:CT31", new CellText(FontPosition.Left)),
                new DailyLogCells("A32:O33", new CellText(FontPosition.Left)), new DailyLogCells("P32:T33"), new DailyLogCells("U32:Y33"), new DailyLogCells("Z32:AD33","SUM(P32*U32)", true), new DailyLogCells("AE32:AT33", new CellText(FontPosition.Left)), new DailyLogCells("AU32:CT33", new CellText(FontPosition.Left)),
                new DailyLogCells("A34:O35", new CellText(FontPosition.Left)), new DailyLogCells("P34:T35"), new DailyLogCells("U34:Y35"), new DailyLogCells("Z34:AD35","SUM(P34*U34)", true), new DailyLogCells("AE34:AT35", new CellText(FontPosition.Left)), new DailyLogCells("AU34:CT35", new CellText(FontPosition.Left)),
                new DailyLogCells("A36:O37", new CellText(FontPosition.Left)), new DailyLogCells("P36:T37"), new DailyLogCells("U36:Y37"), new DailyLogCells("Z36:AD37","SUM(P36*U36)", true), new DailyLogCells("AE36:AT37", new CellText(FontPosition.Left)), new DailyLogCells("AU36:CT37", new CellText(FontPosition.Left)),
                new DailyLogCells("A38:O39", new CellText(FontPosition.Left)), new DailyLogCells("P38:T39"), new DailyLogCells("U38:Y39"), new DailyLogCells("Z38:AD39","SUM(P38*U38)", true), new DailyLogCells("AE38:AT39", new CellText(FontPosition.Left)), new DailyLogCells("AU38:CT39", new CellText(FontPosition.Left)),
                new DailyLogCells("A40:O41", new CellText(FontPosition.Left)), new DailyLogCells("P40:T41"), new DailyLogCells("U40:Y41"), new DailyLogCells("Z40:AD41","SUM(P40*U40)", true), new DailyLogCells("AE40:AT41", new CellText(FontPosition.Left)), new DailyLogCells("AU40:CT41", new CellText(FontPosition.Left)),
                new DailyLogCells("A42:O43", new CellText(FontPosition.Left)), new DailyLogCells("P42:T43"), new DailyLogCells("U42:Y43"), new DailyLogCells("Z42:AD43","SUM(P42*U42)", true), new DailyLogCells("AE42:AT43", new CellText(FontPosition.Left)), new DailyLogCells("AU42:CT43", new CellText(FontPosition.Left)),
                new DailyLogCells("A44:O45", new CellText(FontPosition.Left)), new DailyLogCells("P44:T45"), new DailyLogCells("U44:Y45"), new DailyLogCells("Z44:AD45","SUM(P44*U44)", true), new DailyLogCells("AE44:AT45", new CellText(FontPosition.Left)), new DailyLogCells("AU44:CT45", new CellText(FontPosition.Left)),
                new DailyLogCells("A46:O47", new CellText(FontPosition.Left)), new DailyLogCells("P46:T47"), new DailyLogCells("U46:Y47"), new DailyLogCells("Z46:AD47","SUM(P46*U46)", true), new DailyLogCells("AE46:AT47", new CellText(FontPosition.Left)), new DailyLogCells("AU46:CT47", new CellText(FontPosition.Left)),
                new DailyLogCells("A48:O49", new CellText(FontPosition.Left)), new DailyLogCells("P48:T49"), new DailyLogCells("U48:Y49"), new DailyLogCells("Z48:AD49","SUM(P48*U48)", true), new DailyLogCells("AE48:AT49", new CellText(FontPosition.Left)), new DailyLogCells("AU48:CT49", new CellText(FontPosition.Left)),
                new DailyLogCells("A50:O51", new CellText(FontPosition.Left)), new DailyLogCells("P50:T51"), new DailyLogCells("U50:Y51"), new DailyLogCells("Z50:AD51","SUM(P50*U50)", true), new DailyLogCells("AE50:AT51", new CellText(FontPosition.Left)), new DailyLogCells("AU50:CT51", new CellText(FontPosition.Left)),
                new DailyLogCells("A52:O53", new CellText(FontPosition.Left)), new DailyLogCells("P52:T53"), new DailyLogCells("U52:Y53"), new DailyLogCells("Z52:AD53","SUM(P52*U52)", true), new DailyLogCells("AE52:AT53", new CellText(FontPosition.Left)), new DailyLogCells("AU52:CT53", new CellText(FontPosition.Left)),
                new DailyLogCells("A54:O55", new CellText(FontPosition.Left)), new DailyLogCells("P54:T55"), new DailyLogCells("U54:Y55"), new DailyLogCells("Z54:AD55","SUM(P54*U54)", true), new DailyLogCells("AE54:AT55", new CellText(FontPosition.Left)), new DailyLogCells("AU54:CT55", new CellText(FontPosition.Left)),
                new DailyLogCells("A56:O57", new CellText(FontPosition.Left)), new DailyLogCells("P56:T57"), new DailyLogCells("U56:Y57"), new DailyLogCells("Z56:AD57","SUM(P56*U56)", true), new DailyLogCells("AE56:AT57", new CellText(FontPosition.Left)), new DailyLogCells("AU56:CT57", new CellText(FontPosition.Left)),
                new DailyLogCells("A58:O59", new CellText(FontPosition.Left)), new DailyLogCells("P58:T59"), new DailyLogCells("U58:Y59"), new DailyLogCells("Z58:AD59","SUM(P58*U58)", true), new DailyLogCells("AE58:AT59", new CellText(FontPosition.Left)), new DailyLogCells("AU58:CT59", new CellText(FontPosition.Left)),
                new DailyLogCells("A60:O61", new CellText(FontPosition.Left)), new DailyLogCells("P60:T61"), new DailyLogCells("U60:Y61"), new DailyLogCells("Z60:AD61","SUM(P60*U60)", true), new DailyLogCells("AE60:AT61", new CellText(FontPosition.Left)), new DailyLogCells("AU60:CT61", new CellText(FontPosition.Left)),
                new DailyLogCells("A62:O63", new CellText(FontPosition.Left)), new DailyLogCells("P62:T63"), new DailyLogCells("U62:Y63"), new DailyLogCells("Z62:AD63","SUM(P62*U62)", true), new DailyLogCells("AE62:AT63", new CellText(FontPosition.Left)), new DailyLogCells("AU62:CT63", new CellText(FontPosition.Left)),
                new DailyLogCells("A64:O65", new CellText(FontPosition.Left)), new DailyLogCells("P64:T65"), new DailyLogCells("U64:Y65"), new DailyLogCells("Z64:AD65","SUM(P64*U64)", true), new DailyLogCells("AE64:AT65", new CellText(FontPosition.Left)), new DailyLogCells("AU64:CT65", new CellText(FontPosition.Left)),
                new DailyLogCells("A66:O66"), new DailyLogCells("P66:Y66", "Total Hours", new CellText(FontStyle.Bold)), new DailyLogCells("Z66:AD66","SUM(Z28:AD65)", true), new DailyLogCells("AE66:CT66")

            };

            forestList = new List<Forest>() {
                new Forest("Angelus National Forest", "ANF", "0501"),
                new Forest("Cleveland National Forest", "CNF", "0502"),
                new Forest("Eldorado National Forest", "ENF", "0503"),
                new Forest("Inyo National Forest", "INF", "0504"),
                new Forest("Klamath National Forest", "KNF", "0505"),
                new Forest("Lake Tahoe Basin Mgmt Unit", "TMU", "0519"),
                new Forest("Lassen National Forest", "LNF", "0506"),
                new Forest("Los Padres", "LPF", "0507"),
                new Forest("Mendocino", "MNF", "0508"),
                new Forest("Modoc", "MDF", "0509"),
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

            
        }
    }
}
