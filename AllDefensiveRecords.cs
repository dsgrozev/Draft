using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;

namespace DraftSystem
{
    class AllDefensiveRecords
    {
        static internal List<AllDefensiveRecords> Records = new List<AllDefensiveRecords>();
        internal string PlayerName;
        internal string Team;
        internal string VsTeam;
        internal int WeekNumber;
        public int PtsVs;
        public int Sack;
        public int DefInt;
        public int FumRec;
        public int DefTD;
        public int Safe;
        public int BlkKick;
        public int DefRetTD;

        public AllDefensiveRecords(
            string playerName,
            string team,
            string vsTeam,
            int weekNumber,
            int ptsVs,
            int sack,
            int defInt,
            int fumRec,
            int defTd,
            int safe,
            int blkKick,
            int defRetTd)
        {
            PlayerName = playerName ?? throw new ArgumentNullException(nameof(playerName));
            Team = team ?? throw new ArgumentNullException(nameof(team));
            VsTeam = vsTeam ?? throw new ArgumentNullException(nameof(vsTeam));
            WeekNumber = weekNumber;
            PtsVs = ptsVs;
            Sack = sack;
            DefInt = defInt;
            FumRec = fumRec;
            DefTD = defTd;
            Safe = safe;
            BlkKick = blkKick;
            DefRetTD = defRetTd;
        }

        internal static void ReadExcel(Workbook xlWorkBook)
        {
            _Worksheet sheet = xlWorkBook.Sheets["Defense Data"];
            object[,] range = sheet.UsedRange.Value;
            for (int i = 2; i <= range.GetUpperBound(0); i++)
            {
                int j = 1;
                Records.Add(new AllDefensiveRecords(
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
                    Convert.ToInt32(range[i, j++])
                ));
            }
        }
    }
}
