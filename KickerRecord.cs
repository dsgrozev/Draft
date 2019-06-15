using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;

namespace DraftSystem
{
    class KickerRecord
    {
        static internal List<KickerRecord> Records = new List<KickerRecord>();
        internal string PlayerName;
        internal string Team;
        internal string VsTeam;
        internal int WeekNumber;
        internal int Fg19;
        internal int Fg29;
        internal int Fg39;
        internal int Fg49;
        internal int Fg50;
        internal int Pat;

        public KickerRecord(
            string playerName,
            string team,
            string vsTeam,
            int weekNumber,
            int fg19,
            int fg29,
            int fg39,
            int fg49,
            int fg50,
            int pat)
        {
            PlayerName = playerName ?? throw new ArgumentNullException(nameof(playerName));
            Team = team ?? throw new ArgumentNullException(nameof(team));
            VsTeam = vsTeam ?? throw new ArgumentNullException(nameof(vsTeam));
            WeekNumber = weekNumber;
            Fg19 = fg19;
            Fg29 = fg29;
            Fg39 = fg39;
            Fg49 = fg49;
            Fg50 = fg50;
            Pat = pat;
        }

        internal static void ReadExcel(Workbook xlWorkBook)
        {
            _Worksheet sheet = xlWorkBook.Sheets["Kicker Data"];
            object[,] range = sheet.UsedRange.Value;
            for (int i = 2; i <= range.GetUpperBound(0); i++)
            {
                int j = 1;
                Records.Add(new KickerRecord(
                    (string)range[i, j++],
                    (string)range[i, j++],
                    (string)range[i, j++],
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
