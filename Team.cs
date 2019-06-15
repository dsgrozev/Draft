using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;

namespace DraftSystem
{

    class Team
    {
        static internal List<Team> Teams = new List<Team>();
        internal string Name;
        internal string ShortName;
        internal Dictionary<int, Dictionary<Metric, int>> offensiveRecord =
            new Dictionary<int, Dictionary<Metric, int>>();
        internal Dictionary<int, Dictionary<Metric, int>> defensiveRecord =
            new Dictionary<int, Dictionary<Metric, int>>();
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

        internal void SummarizeRecord(bool offensive)
        {
            Dictionary<int, Dictionary<Metric, int>> record = 
                offensive ? offensiveRecord : defensiveRecord;
            for (int i = 1; i < 18; i++)
            {
                Dictionary<Metric, int> weekRecord = new Dictionary<Metric, int>();
                record.Add(i, weekRecord);
                IEnumerable<OffensiveRecord> recsOff = 
                        OffensiveRecord.Records.Where(x =>
                        (offensive ? x.Team : x.VsTeam) == Name &&
                        x.WeekNumber == i);
                IEnumerable<DefensiveRecord> recsDef =
                        DefensiveRecord.Records.Where(x =>
                        (offensive ? x.Team : x.VsTeam) == Name &&
                        x.WeekNumber == i);
                IEnumerable<KickerRecord> recsKick =
                        KickerRecord.Records.Where(x =>
                        (offensive ? x.Team : x.VsTeam) == Name &&
                        x.WeekNumber == i);
                foreach (Metric m in Enum.GetValues(typeof(Metric)))
                {
                    int sum = 0;
                    foreach (var rec in recsOff)
                    {
                        sum += Ext.GetMetricValue(rec, m);
                    }
                    foreach (var rec in recsDef)
                    {
                        sum += Ext.GetMetricValue(rec, m);
                    }
                    foreach (var rec in recsKick)
                    {
                        sum += Ext.GetMetricValue(rec, m);
                    }
                    weekRecord.Add(m, sum);
                }
            }
        }
    }
}
