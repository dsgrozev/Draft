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
        internal Dictionary<int, Team> Schedule = new Dictionary<int, Team>();
        
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

        internal static Team FindTeamByName(string name) => Teams.Find(x => x.Name == name);
        internal static Team FindTeamByShortName(string shortName) => 
            Teams.Find(x => x.ShortName == shortName);

        internal void SummarizeRecord(bool offensive)
        {
            Dictionary<int, Dictionary<Metric, int>> record =
                offensive ? OffensiveRecord : DefensiveRecord;
            for (int i = 1; i < AllOffensiveRecords.Records.Select(x => x.WeekNumber).Max() + 1; i++)
            {
                Dictionary<Metric, int> weekRecord = new Dictionary<Metric, int>();
                IEnumerable<AllOffensiveRecords> recsOff = 
                        AllOffensiveRecords.Records.Where(x =>
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

        internal void UpdateSchedule()
        {
            ScheduleRecord sr = ScheduleRecord.Records.First(x => x.Team == this.Name);
            for (int i = 1; i < 17; i++)
            {
                string shortName = sr.GetOppByWeek(i);
                if (shortName != "Bye")
                {
                    this.Schedule.Add(i, FindTeamByShortName(shortName));
                }
            }
        }

        internal void UpdatePlayerCoef()
        {
            foreach (Position p in Enum.GetValues(typeof(Position)))
            {
                var players = Player.AllPlayers.Where(x => x.Team == this && x.Position == p);
                foreach (Metric m in Enum.GetValues(typeof(Metric)))
                {
                    double sumCoef = this.MetricByPositionCoeficients[m][p];
                    double sumPlayers = players.Sum(x => x.isNew ? 0 : x.coef[m]);
                    if (sumPlayers >= sumCoef)
                    {
                        foreach(Player pl in players)
                        {
                            pl.realCoef[m] = pl.isNew ? 0 : pl.coef[m]; 
                        }
                    }
                    else
                    {
                        if (players.Any(x => x.isNew))
                        {
                            double add = (sumCoef - sumPlayers) / players.Count(x => x.isNew);
                            foreach (Player pl in players)
                            {
                                pl.realCoef[m] = pl.isNew ? add : pl.coef[m];
                            }
                        }
                        else
                        {
                            double add = (sumCoef - sumPlayers) / players.Count();
                            foreach (Player pl in players)
                            {
                                pl.realCoef[m] = pl.coef[m] + add;
                            }
                        }

                    }
                }
            }
        }

        internal void PosCounts()
        {
            IEnumerable<AllOffensiveRecords> recsOff = AllOffensiveRecords.Records.Where(x => x.Team == Name);
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
            var allPlayers = AllOffensiveRecords.Records.Where(x => x.Team == Name);
            foreach (Metric m in Enum.GetValues(typeof(Metric)))
            {
                foreach (Position pos in Enum.GetValues(typeof(Position)))
                {
                    List<double> values = new List<double>();
                    for (int i = 1; i <= allPlayers.Max(x => x.WeekNumber); i++) 
                    {
                        if (!OffensiveRecord.ContainsKey(i))
                        {
                            continue;
                        }
                        double sum = 
                            allPlayers.Where(x => x.WeekNumber == i && x.Position == pos.ToString()).Sum(x => x.GetMetricValue(m));

                        if (OffensiveRecord[i][m] != 0)
                        {
                            values.Add(sum / OffensiveRecord[i][m]);
                        }
                        
                    }
                    double avg = FindWeightedAverage(values);

                    if(!MetricByPositionCoeficients.ContainsKey(m))
                    {
                        MetricByPositionCoeficients.Add(m, new Dictionary<Position, double>());
                    }
                    MetricByPositionCoeficients[m].Add(pos, avg);
                }
            }
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

        public static double FindWeightedAverage(List<double> recs)
        {
            double[] arr = recs.ToArray();
            int count = arr.Length;
            int divider = count;
            if (count == 0)
            {
                return 0;
            }
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

        public static double FindWeightedAverage(IEnumerable<int> recs)
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
