using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace DraftSystem
{
    class Player
    {
        internal static readonly List<Player> AllPlayers = new List<Player>();
        internal readonly string Name;
        internal readonly Team Team;
        internal readonly Position Position;
        internal int suspension;
        internal bool isNew = true;
        internal Dictionary<Metric, double> coef = new Dictionary<Metric, double>();
        internal Dictionary<Metric, double> realCoef = new Dictionary<Metric, double>();
        internal Dictionary<int, double> expectedPoints = new Dictionary<int, double>();

        public Player(string name, Team team, Position position, int suspension)
        {
            Name = name ?? throw new ArgumentNullException(nameof(name));
            this.Team = team ?? throw new ArgumentNullException(nameof(team));
            this.Position = position;
            this.suspension = suspension;
        }

        internal static void Init()
        {
            foreach(var line in TeamPlayersRecord.Records)
            {
                AllPlayers.Add(new Player(line.PlayerName, Team.FindTeamByName(line.Team),
                    (Position)Enum.Parse(typeof(Position), line.Position), line.Suspension));
            }
            foreach(Team t in Team.Teams)
            {
                AllPlayers.Add(new Player(t.Name, t, Position.DEF, 0));
            }
        }

        internal void FindPrevCoef()
        {
            Dictionary<Metric, List<double>> playerRecord = new Dictionary<Metric, List<double>>();
            IEnumerable records;
            if (this.Position == Position.DEF)
            {
                records = AllDefensiveRecords.Records.Where(x => x.Team == this.Name);
            }
            else if (this.Position == Position.K)
            {
                records = KickerRecord.Records.Where(x => x.PlayerName.StartsWith(this.Name));
            }
            else
            {
                records = AllOffensiveRecords.Records.Where(x => x.PlayerName.StartsWith(this.Name));
            }
            foreach (var r in records)
            {
                Team oppTeam = r.GetOppTeam();
                this.isNew = false;
                foreach (Metric m in Enum.GetValues(typeof(Metric)))
                {
                    double rec = oppTeam.DefensiveSummary[m] == 0 ? 0 : 
                        1.0 * r.GetMetricValue(m) / oppTeam.DefensiveSummary[m];
                    if (!playerRecord.ContainsKey(m))
                    {
                        playerRecord.Add(m, new List<double>());
                    }
                    playerRecord[m].Add(rec);
                }
            }
            if (!this.isNew)
            {
                foreach (Metric m in Enum.GetValues(typeof(Metric)))
                {
                    coef.Add(m, Team.FindWeightedAverage(playerRecord[m]));
                }
            }
        }

        internal void FindExpectedPoints()
        {
            foreach(int week in this.Team.Schedule.Keys)
            {
                if (week <= suspension)
                {
                    this.expectedPoints.Add(week, 0);
                    continue;
                }
                Team oppTeam = this.Team.Schedule[week];
                double sum = 0;
                foreach (Metric m in Enum.GetValues(typeof(Metric)))
                {
                    sum += Ext.PointsPerMetric(m, this.realCoef[m] * oppTeam.DefensiveSummary[m]);
                }
                this.expectedPoints.Add(week, sum);
            }
        }
    }
}
