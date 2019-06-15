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
            Console.Write("");

            // For each team summarize offensive record


            // For each tema summarize defensive record


            // For each player summarize coeficients
            // For each team find pos / metric coeficients
            // Calculate team counts
            // For each team find team players expected coeficients
            // For each player find expected points

        }
    }
}
