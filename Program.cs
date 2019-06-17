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
                .Open(@"C:\Users\Dimitar\OneDrive\Documents\Copy of FantasyWeek1_14_2019.xlsx");
            // Read Defensive Data
            DefensiveRecord.ReadExcel(xlWorkBook);
            // Read Kicker Data
            KickerRecord.ReadExcel(xlWorkBook);
            // Read Offensive Data
            OffensiveRecord.ReadExcel(xlWorkBook);
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

            Console.Write("");
            xlApp = new Application();
            xlApp.Workbooks.Add();
            string[,] offTable = new string[33, 26];
            string[,] defTable = new string[33, 26];
            int i = 0;
            foreach (Team t in Team.Teams)
            {
                offTable[i, 0] = t.Name;
                defTable[i, 0] = t.Name;
                int j = 1;
                foreach(Metric m in Enum.GetValues(typeof(Metric)))
                {
                    offTable[i, j] = t.offensiveSummary[m].ToString();
                    defTable[i, j++] = t.defensiveSummary[m].ToString();
                }
                i++;
            }
            _Worksheet workSheet = xlApp.ActiveSheet;
            workSheet.Name = "testOff";
            workSheet.UsedRange.Value = offTable;
            xlApp.Worksheets.Add();
            workSheet.Name = "testDef";
            workSheet.UsedRange.Value = defTable;
            workSheet.SaveAs(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\testNew.xlsx");
            xlApp.Quit();

            // For each player summarize coeficients
            // For each team find pos / metric coeficients
            // Calculate team counts
            // For each team find team players expected coeficients
            // For each player find expected points

        }
    }
}
