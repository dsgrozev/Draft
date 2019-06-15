using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;

namespace DraftSystem
{
    class OffensiveRecord
    {
        static internal List<OffensiveRecord> Records = new List<OffensiveRecord>();
        internal string PlayerName;
        internal string Position;
        internal string Team;
        internal string VsTeam;
        internal int WeekNumber;
        public int PassYds;
        public int PassTd;
        public int PassInt;
        public int RushYds;
        public int RushTd;
        public int Rec;
        public int RecYds;
        public int RecTd;
        public int RetTd;
        public int TwoPt;
        public int FumLost;

        public OffensiveRecord(
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
            PassTd = passTd;
            PassInt = passInt;
            RushYds = rushYds;
            RushTd = rushTd;
            Rec = rec;
            RecYds = recYds;
            RecTd = recTd;
            RetTd = retTd;
            TwoPt = twoPt;
            FumLost = fumLost;
        }

        internal static void ReadExcel(Workbook xlWorkBook)
        {
            _Worksheet sheet = xlWorkBook.Sheets["Offense Data"];
            object[,] range = sheet.UsedRange.Value;
            for (int i = 2; i <= range.GetUpperBound(0); i++)
            {
                int j = 1;
                Records.Add(new OffensiveRecord(
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
