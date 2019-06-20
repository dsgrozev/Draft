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
        internal Dictionary<int, Dictionary<Metric, int>> OffensiveRecord =
            new Dictionary<int, Dictionary<Metric, int>>();
        internal Dictionary<int, Dictionary<Metric, int>> DefensiveRecord =
            new Dictionary<int, Dictionary<Metric, int>>();
        internal Dictionary<Metric, double> OffensiveSummary = new Dictionary<Metric, double>();
        internal Dictionary<Metric, double> DefensiveSummary = new Dictionary<Metric, double>();
        internal Dictionary<Position, int> PositionCounts = new Dictionary<Position, int>();
        internal Dictionary<Metric, Dictionary<Position, double>> MetricByPositionCoeficients = 
            new Dictionary<Metric, Dictionary<Position, double>>();
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
                offensive ? OffensiveRecord : DefensiveRecord;
            for (int i = 1; i < AllOffensiveRecordS.Records.Select(x => x.WeekNumber).Max() + 1; i++)
            {
                Dictionary<Metric, int> weekRecord = new Dictionary<Metric, int>();
                IEnumerable<AllOffensiveRecordS> recsOff = 
                        AllOffensiveRecordS.Records.Where(x =>
                        (offensive ? x.Team : x.VsTeam) == Name &&
                        x.WeekNumber == i);
                IEnumerable<AllDefensiveRecords> recsDef =
                        AllDefensiveRecords.Records.Where(x =>
                        (offensive ? x.Team : x.VsTeam) == Name &&
                        x.WeekNumber == i);
                IEnumerable<KickerRecord> recsKick =
                        KickerRecord.Records.Where(x =>
                        (offensive ? x.Team : x.VsTeam) == Name &&
                        x.WeekNumber == i);

                if (recsDef.Count() > 0)
                {
                    record.Add(i, weekRecord);
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

        internal void PosCounts()
        {
            IEnumerable<AllOffensiveRecordS> recsOff = AllOffensiveRecordS.Records.Where(x => x.Team == Name);
            int weeks = AllDefensiveRecords.Records.Where(x => x.Team == Name).Count();
            foreach (Position p in Enum.GetValues(typeof(Position)))
            {
                if (p != Position.DEF && p != Position.K)
                {
                    List<int> counts = new List<int>();
                    for (int i = 1; i <= weeks; i++)
                    {
                        int players = recsOff.Where(x => x.Position == p.ToString() && x.WeekNumber == i).Count();
                        if (players != 0)
                        {
                            counts.Add(players);
                        }
                    }
                    PositionCounts[p] = Convert.ToInt32(FindWeightedAverage(counts));
                }
            }
        }

        internal void CalculatePosCoef()
        {

        }

        internal void DefenseSummary()
        {
            foreach (Metric m in Enum.GetValues(typeof(Metric)))
            {
                IEnumerable<int> recs = DefensiveRecord.Select(x => x.Value[m]);
                DefensiveSummary.Add(m, FindWeightedAverage(recs));
            }
        }

        internal void OffenseSummary()
        {
            Dictionary<Metric, List<double>> offCoef = new Dictionary<Metric, List<double>>();
            foreach (KeyValuePair<int, Dictionary<Metric, int>> rec in OffensiveRecord)
            {
                int week = rec.Key;
                IEnumerable<AllDefensiveRecords> game = 
                    AllDefensiveRecords.Records.Where(x => x.WeekNumber == week && x.Team == Name);
                if (game.Count() == 0)
                {
                    continue;
                }
                string oppName = game.First().VsTeam;
                Team oppTeam = FindTeamByName(oppName);
                foreach (Metric m in Enum.GetValues(typeof(Metric)))
                {
                    double defAverage = oppTeam.DefensiveSummary[m];
                    double coef = 0;
                    if (defAverage != 0)
                    {
                        coef = rec.Value[m] / defAverage;
                    }
                    if (!offCoef.ContainsKey(m))
                    {
                        offCoef[m] = new List<double>();
                    }
                    offCoef[m].Add(coef);
                }
            }
            foreach (KeyValuePair<Metric, List<double>> list in offCoef)
            {
                OffensiveSummary.Add(list.Key, FindWeightedAverage(list.Value));
            }
        }

        private double FindWeightedAverage(List<double> recs)
        {
            double[] arr = recs.ToArray();
            int count = arr.Length;
            int divider = count;
            double sum = arr.Sum();

            if (count > 4)
            {
                sum += arr[count - 1] * 3;
                sum += arr[count - 2] * 3;
                sum += arr[count - 3];
                sum += arr[count - 4];
                divider += 8;
            }
            else if (count > 2)
            {
                sum += arr[count - 1];
                sum += arr[count - 2];
                divider += 2;
            }

            return sum / divider;
        }

        private double FindWeightedAverage(IEnumerable<int> recs)
        {
            int[] arr = recs.ToArray();
            int count = arr.Length;
            int divider = count;
            int sum = arr.Sum();

            if (count > 4)
            {
                sum += arr[count - 1] * 3;
                sum += arr[count - 2] * 3;
                sum += arr[count - 3];
                sum += arr[count - 4];
                divider += 8;
            }else if (count > 2)
            {
                sum += arr[count - 1];
                sum += arr[count - 2];
                divider += 2;
            }

            return (double)sum / divider;
        }
    }
}
