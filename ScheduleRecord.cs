using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;

namespace DraftSystem
{
    class ScheduleRecord
    {
        static internal List<ScheduleRecord> Records = new List<ScheduleRecord>();
        internal string Team;
        internal string Week1;
        internal string Week2;
        internal string Week3;
        internal string Week4;
        internal string Week5;
        internal string Week6;
        internal string Week7;
        internal string Week8;
        internal string Week9;
        internal string Week10;
        internal string Week11;
        internal string Week12;
        internal string Week13;
        internal string Week14;
        internal string Week15;
        internal string Week16;

        public ScheduleRecord(
            string team,
            string week1,
            string week2,
            string week3,
            string week4,
            string week5,
            string week6,
            string week7,
            string week8,
            string week9,
            string week10,
            string week11,
            string week12,
            string week13,
            string week14,
            string week15,
            string week16)
        {
            Team = team ?? throw new ArgumentNullException(nameof(team));
            Week1 = week1 ?? throw new ArgumentNullException(nameof(week1));
            Week2 = week2 ?? throw new ArgumentNullException(nameof(week2));
            Week3 = week3 ?? throw new ArgumentNullException(nameof(week3));
            Week4 = week4 ?? throw new ArgumentNullException(nameof(week4));
            Week5 = week5 ?? throw new ArgumentNullException(nameof(week5));
            Week6 = week6 ?? throw new ArgumentNullException(nameof(week6));
            Week7 = week7 ?? throw new ArgumentNullException(nameof(week7));
            Week8 = week8 ?? throw new ArgumentNullException(nameof(week8));
            Week9 = week9 ?? throw new ArgumentNullException(nameof(week9));
            Week10 = week10 ?? throw new ArgumentNullException(nameof(week10));
            Week11 = week11 ?? throw new ArgumentNullException(nameof(week11));
            Week12 = week12 ?? throw new ArgumentNullException(nameof(week12));
            Week13 = week13 ?? throw new ArgumentNullException(nameof(week13));
            Week14 = week14 ?? throw new ArgumentNullException(nameof(week14));
            Week15 = week15 ?? throw new ArgumentNullException(nameof(week15));
            Week16 = week16 ?? throw new ArgumentNullException(nameof(week16));
        }

        internal static void ReadExcel(Workbook xlWorkBook)
        {
            _Worksheet sheet = xlWorkBook.Sheets["Schedule"];
            object[,] range = sheet.UsedRange.Value;
            for (int i = 2; i <= range.GetUpperBound(0); i++)
            {
                int j = 1;
                Records.Add(new ScheduleRecord(
                    (string)range[i, j++],
                    (string)range[i, j++],
                    (string)range[i, j++],
                    (string)range[i, j++],
                    (string)range[i, j++],
                    (string)range[i, j++],
                    (string)range[i, j++],
                    (string)range[i, j++],
                    (string)range[i, j++],
                    (string)range[i, j++],
                    (string)range[i, j++],
                    (string)range[i, j++],
                    (string)range[i, j++],
                    (string)range[i, j++],
                    (string)range[i, j++],
                    (string)range[i, j++],
                    (string)range[i, j++]
                ));
            }
        }
    }
}
