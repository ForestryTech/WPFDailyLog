using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WPFDailyLog
{
    partial class MonthlySummarySheet : LogWorkSheet
    {
        private List<DailyLogCells> rowHeadings;


        private void fillMonthlySummaryLists() {
            rowHeadings = new List<DailyLogCells>() {
                new DailyLogCells("A2", "ACTIVITIES", new CellText(FontStyle.Bold, FontPosition.Left), CellBorder.No),
                new DailyLogCells("A3", Properties.Resources.StationAdministration, new CellText(FontPosition.Left)),
                new DailyLogCells("A4", Properties.Resources.StationMaintenance, new CellText(FontPosition.Left)),
                new DailyLogCells("A5", Properties.Resources.EnginePMReadiness, new CellText(FontPosition.Left)),
                new DailyLogCells("A6", Properties.Resources.PhysicalTraining, new CellText(FontPosition.Left)),
                new DailyLogCells("A7", Properties.Resources.Training, new CellText(FontPosition.Left)),
                new DailyLogCells("A8", Properties.Resources.LocalForestCadre, new CellText(FontPosition.Left)),
                new DailyLogCells("A9", Properties.Resources.RegionalCadre, new CellText(FontPosition.Left)),
                new DailyLogCells("A10", Properties.Resources.MeetingBriefing, new CellText(FontPosition.Left)),
                new DailyLogCells("A11", Properties.Resources.MeetingBriefing, new CellText(FontPosition.Left)),
                new DailyLogCells("A12", Properties.Resources.WildlandFireBase, new CellText(FontPosition.Left)),
                new DailyLogCells("A13", Properties.Resources.WildlandFireOT, new CellText(FontPosition.Left)),
                new DailyLogCells("A14", Properties.Resources.OverheadAssignments, new CellText(FontPosition.Left)),
                new DailyLogCells("A15", Properties.Resources.TrafficCollision, new CellText(FontPosition.Left)),
                new DailyLogCells("A16", Properties.Resources.MedicalAid, new CellText(FontPosition.Left)),
                new DailyLogCells("A17", Properties.Resources.Hazmat, new CellText(FontPosition.Left)),
                new DailyLogCells("A18", Properties.Resources.StructureFire, new CellText(FontPosition.Left)),
                new DailyLogCells("A19", Properties.Resources.VehicleFire, new CellText(FontPosition.Left)),
                new DailyLogCells("A20", Properties.Resources.StandByCover, new CellText(FontPosition.Left)),
                new DailyLogCells("A21", Properties.Resources.LEAssistSR, new CellText(FontPosition.Left)),
                new DailyLogCells("A22", Properties.Resources.FalseAlarmCancel, new CellText(FontPosition.Left)),
                new DailyLogCells("A23", Properties.Resources.ExtendedStaffing, new CellText(FontPosition.Left)),
                new DailyLogCells("A24", Properties.Resources.OtherEmergency, new CellText(FontPosition.Left)),
                new DailyLogCells("A25", Properties.Resources.NonFireOT, new CellText(FontPosition.Left)),
                new DailyLogCells("A26", Properties.Resources.OtherHours, new CellText(FontPosition.Left)),
                new DailyLogCells("A27", Properties.Resources.FireRehabAdmin, new CellText(FontPosition.Left)),
                new DailyLogCells("A29", "PROJECTS", new CellText(FontStyle.Bold, FontPosition.Left), CellBorder.No),
                new DailyLogCells("A30", Properties.Resources.FuelsRXBurning, new CellText(FontPosition.Left)),
                new DailyLogCells("A31", Properties.Resources.PreSuppression, new CellText(FontPosition.Left)),
                new DailyLogCells("A32", Properties.Resources.Safety, new CellText(FontPosition.Left)),
                new DailyLogCells("A33", Properties.Resources.RecruitmentOutreach, new CellText(FontPosition.Left)),
                new DailyLogCells("A34", Properties.Resources.Hiring, new CellText(FontPosition.Left)),
                new DailyLogCells("A35", Properties.Resources.RecreationAssist, new CellText(FontPosition.Left)),
                new DailyLogCells("A36", Properties.Resources.PreventionPatrol, new CellText(FontPosition.Left)),
                new DailyLogCells("A37", Properties.Resources.EngineeringAssist, new CellText(FontPosition.Left)),
                new DailyLogCells("A38", Properties.Resources.PublicMedia, new CellText(FontPosition.Left)),
                new DailyLogCells("A39", Properties.Resources.ResourceAssist, new CellText(FontPosition.Left)),
                new DailyLogCells("A40", Properties.Resources.WCT, new CellText(FontPosition.Left)),
                new DailyLogCells("A41", Properties.Resources.PhysicalsDrugTest, new CellText(FontPosition.Left)),
                new DailyLogCells("A42", Properties.Resources.FacilitiesMaintenance, new CellText(FontPosition.Left)),
                new DailyLogCells("A44", "PERSONNEL", new CellText(FontStyle.Bold, FontPosition.Left), CellBorder.No),
                new DailyLogCells("A45", "# of Personnel", new CellText(FontPosition.Left)),
                new DailyLogCells("A46", "Total Hours", new CellText(FontPosition.Left)),
                new DailyLogCells("A49", "STATS", new CellText(FontStyle.Bold, FontPosition.Left), CellBorder.No),
                new DailyLogCells("A50", "Off Forest Assignments", new CellText(FontPosition.Left)),
                new DailyLogCells("A51", "    Module", new CellText(FontPosition.Left)),
                new DailyLogCells("A52", "    Single Resource", new CellText(FontPosition.Left)),
                new DailyLogCells("A53", "Wildland Fires", new CellText(FontPosition.Left)),
                new DailyLogCells("A54", "Vehicle Fires", new CellText(FontPosition.Left)),
                new DailyLogCells("A55", "Structure Fires", new CellText(FontPosition.Left)), 
                new DailyLogCells("A56", "Local Single Resource", new CellText(FontPosition.Left)),
                new DailyLogCells("A57", "False Alarm / Smoke Check", new CellText(FontPosition.Left)),
                new DailyLogCells("A58", "Hazmat", new CellText(FontPosition.Left)),
                new DailyLogCells("A59", "Medical Aid", new CellText(FontPosition.Left)),
                new DailyLogCells("A60", "Traffic Accident", new CellText(FontPosition.Left)),
                new DailyLogCells("A61", "Cover Assignment", new CellText(FontPosition.Left)),
                new DailyLogCells("A62", "Cancelations", new CellText(FontPosition.Left)),
                new DailyLogCells("A63", "Firefighter Standby", new CellText(FontPosition.Left)),
                new DailyLogCells("A64", "Other Response", new CellText(FontPosition.Left)),
                new DailyLogCells("A65", "Total # of Responses", new CellText(FontPosition.Left)),
                new DailyLogCells("A67", "Out of Region Days", new CellText(FontStyle.Bold, FontPosition.Left), CellBorder.No),
                new DailyLogCells("A68", "    Module", new CellText(FontPosition.Left)),
                new DailyLogCells("A69", "    Single Resource", new CellText(FontPosition.Left)),
                new DailyLogCells("A71", "In Region Assignment Days", new CellText(FontStyle.Bold, FontPosition.Left), CellBorder.No),
                new DailyLogCells("A72", "    Module", new CellText(FontPosition.Left)),
                new DailyLogCells("A73", "    Single Resource", new CellText(FontPosition.Left)),
                new DailyLogCells(columnHeads[columnHeads.Count - 2] + "45:" + columnHeads[columnHeads.Count - 1] + "45", "Total Monthly Hours", new CellText(FontStyle.Bold, FontPosition.Center), CellBorder.No)
            };
        }


        
    }
}
