using System;
using Microsoft.Office.Interop.Excel;

namespace DraftSystem
{
    class Program
    {
        static void Main(string[] args)
        {
            Application xlApp = new Application();
            Workbook xlWorkBook = xlApp.Workbooks
                .Open(@"C:\FF\211439_1_19_2020_copy.xlsx");
            // Read Defensive Data
            AllDefensiveRecords.ReadExcel(xlWorkBook);
            // Read Kicker Data
            KickerRecord.ReadExcel(xlWorkBook);
            // Read Offensive Data
            AllOffensiveRecords.ReadExcel(xlWorkBook);
            // Read Schedule
            ScheduleRecord.ReadExcel(xlWorkBook);
            // Read Team Players
            TeamPlayersRecord.ReadExcel(xlWorkBook);
            // Read Teams
            Team.ReadExcel(xlWorkBook);
            xlWorkBook.Close();

            // For each team summarize offensive record
            // For each tema summarize defensive record
            foreach (Team t in Team.Teams)
            {
                t.SummarizeRecord(true);
                t.SummarizeRecord(false);
            }
            
            // Find average def values
            foreach (Team t in Team.Teams)
            {
                t.DefenseSummary();
            }
            // Find average offensive coef
            foreach (Team t in Team.Teams)
            {
                t.OffenseSummary();
            }
            // Calculate team counts 
            foreach (Team t in Team.Teams)
            {
                t.PosCounts();
            }
            // For each team find pos / metric coeficients
            foreach (Team t in Team.Teams)
            {
                t.CalculatePosCoef();
            }
            // For each player summarize coeficients
            Player.Init();
            foreach (Player p in Player.AllPlayers)
            {
                p.FindPrevCoef();
            }
            // For each team find team players expected coeficients
            foreach(Team t in Team.Teams)
            {
                t.UpdatePlayerCoef();
            }
            // For each team update schedule
            foreach(Team t in Team.Teams)
            {
                t.UpdateSchedule();
            }
            // For each player find expected points
            foreach(Player p in Player.AllPlayers)
            {
                p.FindExpectedPoints();
            }

            xlApp = new Application();
            xlApp.Workbooks.Add();
            _Worksheet workSheet = xlApp.ActiveSheet;
            workSheet.Name = "Coeficients";

            int col = 1;
            workSheet.Cells[1, col++] = "Name";
            workSheet.Cells[1, col++] = "Position";

            foreach (string m in Enum.GetNames(typeof(Metric)))
            {
                workSheet.Cells[1, col++] = m;
            }

            col = 1;
            int row = 2;
            foreach (Player p in Player.AllPlayers)
            {
                workSheet.Cells[row, col++] = p.Name;
                workSheet.Cells[row, col++] = p.Position.ToString();
                foreach (Metric m in Enum.GetValues(typeof(Metric)))
                {
                    workSheet.Cells[row, col++] = p.realCoef[m];
                }
                row++;
                col = 1;
            }

            xlApp.Worksheets.Add();
            workSheet = xlApp.ActiveSheet;
            workSheet.Name = "Weeks";

            col = 1;
            workSheet.Cells[1, col++] = "Name";
            workSheet.Cells[1, col++] = "Position";
            workSheet.Cells[1, col++] = "Team";
            workSheet.Cells[1, col++] = "Drafted";
            workSheet.Cells[1, col++] = "ADP";
            workSheet.Cells[1, col++] = "Tier";

            for (int i = 1; i < 17; i++)
            {
                workSheet.Cells[1, col++] = "Week " + i;
            }

            col = 1;
            row = 2;
            foreach (Player p in Player.AllPlayers)
            {
                workSheet.Cells[row, col++] = p.Name;
                workSheet.Cells[row, col++] = p.Position.ToString();
                workSheet.Cells[row, col++] = p.Team.ShortName;
                col += 3;
                for (int i = 1; i < 17; i++)
                {
                    if (p.expectedPoints.ContainsKey(i))
                    {
                        workSheet.Cells[row, col++] = p.expectedPoints[i];
                    }
                    else
                    {
                        workSheet.Cells[row, col++] = 0;
                    }
                }
                row++;
                col = 1;
            }

            workSheet.SaveAs(@"C:\FF\testNew.xlsx");
            xlApp.Quit();
        }
    }
}
