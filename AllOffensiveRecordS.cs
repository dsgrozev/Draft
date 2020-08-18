using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;

namespace DraftSystem
{
    class AllOffensiveRecords
    {
        static internal List<AllOffensiveRecords> Records = new List<AllOffensiveRecords>();
        public string PlayerName;
        public string Position;
        public string Team;
        public string VsTeam;
        public int WeekNumber;
        public int PassYds;
        public int PassTD;
        public int PassInt;
        public int RushYds;
        public int RushTD;
        public int Rec;
        public int RecYds;
        public int RecTD;
        public int RetTD;
        public int TwoPt;
        public int FumLost;

        public AllOffensiveRecords(
            string playerName,
            string position,
            string team,
            string vsTeam,
            int weekNumber,
            int passYds,
            int passTd,
            int passInt,
            int rushYds,
            int rushTd,
            int rec,
            int recYds,
            int recTd,
            int retTd,
            int twoPt,
            int fumLost)
        {
            PlayerName = playerName ?? throw new ArgumentNullException(nameof(playerName));
            Position = position ?? throw new ArgumentNullException(nameof(position));
            Team = team ?? throw new ArgumentNullException(nameof(team));
            VsTeam = vsTeam ?? throw new ArgumentNullException(nameof(vsTeam));
            WeekNumber = weekNumber;
            PassYds = passYds;
            PassTD = passTd;
            PassInt = passInt;
            RushYds = rushYds;
            RushTD = rushTd;
            Rec = rec;
            RecYds = recYds;
            RecTD = recTd;
            RetTD = retTd;
            TwoPt = twoPt;
            FumLost = fumLost;
        }

        internal static void ReadExcel(Workbook xlWorkBook)
        {
            _Worksheet sheet = xlWorkBook.Sheets["Offence Data"];
            object[,] range = sheet.UsedRange.Value;
            for (int i = 2; i <= range.GetUpperBound(0); i++)
            {
                int j = 1;
                Records.Add(new AllOffensiveRecords(
                    (string)range[i, j++],
                    (string)range[i, j++],
                    (string)range[i, j++],
                    (string)range[i, j++],
                    Convert.ToInt32(range[i, j++]),
                    Convert.ToInt32(range[i, j++]),
                    Convert.ToInt32(range[i, j++]),
                    Convert.ToInt32(range[i, j++]),
                    Convert.ToInt32(range[i, j++]),
                    Convert.ToInt32(range[i, j++]),
                    Convert.ToInt32(range[i, j++]),
                    Convert.ToInt32(range[i, j++]),
                    Convert.ToInt32(range[i, j++]),
                    Convert.ToInt32(range[i, j++]),
                    Convert.ToInt32(range[i, j++]),
                    Convert.ToInt32(range[i, j++])
                ));
            }
        }
    }
}