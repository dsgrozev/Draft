using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;

namespace DraftSystem
{
    class TeamPlayersRecord
    {
        static internal List<TeamPlayersRecord> Records = new List<TeamPlayersRecord>();
        internal string PlayerName;
        internal string Team;
        internal string Position;
        internal int Suspension;

        public TeamPlayersRecord(string playerName, string team, string position, int suspension)
        {
            PlayerName = playerName ?? throw new ArgumentNullException(nameof(playerName));
            Team = team ?? throw new ArgumentNullException(nameof(team));
            Position = position ?? throw new ArgumentNullException(nameof(position));
            Suspension = suspension;
        }

        internal static void ReadExcel(Workbook xlWorkBook)
        {
            _Worksheet sheet = xlWorkBook.Sheets["TeamPlayers"];
            object[,] range = sheet.UsedRange.Value;
            for (int i = 2; i <= range.GetUpperBound(0); i++)
            {
                int j = 1;
                Records.Add(new TeamPlayersRecord(
                    (string)range[i, j++],
                    (string)range[i, j++],
                    (string)range[i, j++],
                    Convert.ToInt32(range[i, j++])
                ));
            }
        }
    }
}
