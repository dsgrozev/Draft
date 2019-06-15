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
        internal int FG19;
        internal int FG29;
        internal int FG39;
        internal int FG49;
        internal int FG50;
        internal int PAT;

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
            FG19 = fg19;
            FG29 = fg29;
            FG39 = fg39;
            FG49 = fg49;
            FG50 = fg50;
            PAT = pat;
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
        public static int GetMetricValue(this object kickerRecord, string metric)
        {
            return (int)kickerRecord.GetType().GetProperty(metric).GetValue(kickerRecord);
        }
    }
}
