using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;

namespace DraftSystem
{
    class DefensiveRecord
    {
        static internal List<DefensiveRecord> Records = new List<DefensiveRecord>();
        internal string PlayerName;
        internal string Team;
        internal string VsTeam;
        internal int WeekNumber;
        internal int PtsVs;
        internal int Sack;
        internal int DefInt;
        internal int FumRec;
        internal int DefTd;
        internal int Safe;
        internal int BlkKick;
        internal int DefRetTd;

        public DefensiveRecord(
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
            DefTd = defTd;
            Safe = safe;
            BlkKick = blkKick;
            DefRetTd = defRetTd;
        }

        internal static void ReadExcel(Workbook xlWorkBook)
        {
            _Worksheet sheet = xlWorkBook.Sheets["Defense Data"];
            object[,] range = sheet.UsedRange.Value;
            for (int i = 2; i <= range.GetUpperBound(0); i++)
            {
                int j = 1;
                Records.Add(new DefensiveRecord(
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
