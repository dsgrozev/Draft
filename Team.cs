using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;

namespace DraftSystem
{

    class Team
    {
        static internal List<Team> Teams = new List<Team>();
        internal string Name;
        internal string ShortName;

        public Team(string name, string shortName)
        {
            Name = name;
            ShortName = shortName;
        }

        internal static void ReadExcel(Workbook xlWorkBook)
        {
            _Worksheet sheet = xlWorkBook.Sheets["Teams"];
            object[,] range = sheet.UsedRange.Value;
            for (int i = 2; i <= range.GetUpperBound(0); i++)
            {
                Teams.Add(new Team((string)range[i, 1], (string)range[i, 2]));
            }
        }

        internal Team FindTeamByName(string name) => Teams.Find(x => x.Name == name);
        internal Team FindTeamByShortName(string shortName) => 
            Teams.Find(x => x.ShortName == shortName);
    }
}
