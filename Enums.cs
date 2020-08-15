using System;
using System.Collections.Generic;
using System.Reflection;

namespace DraftSystem
{
    internal static class Ext
    {
        internal static Dictionary<Metric, double> points = new Dictionary<Metric, double>()
        {
            {Metric.PassYds, .04},
            {Metric.PassTD, 6},
            {Metric.PassInt, -2},
            {Metric.RushYds, .1},
            {Metric.RushTD, 6},
            {Metric.Rec, .5},
            {Metric.RecYds, .1 },
            {Metric.RecTD, 6},
            {Metric.RetTD, 6},
            {Metric.TwoPt, 2},
            {Metric.FumLost, -2},
            {Metric.FG19, 2},
            {Metric.FG29, 2},
            {Metric.FG39, 2},
            {Metric.FG49, 3},
            {Metric.FG50, 4},
            {Metric.PAT, 1},
            {Metric.Sack, 1},
            {Metric.DefInt, 2},
            {Metric.FumRec, 2},
            {Metric.DefTD, 6},
            {Metric.Safe, 2},
            {Metric.BlkKick, 2},
            {Metric.DefRetTD, 6}
        };
        public static int GetMetricValue(this object record, Metric metric)
        {
            FieldInfo p = record.GetType().GetField(metric.ToString());
            if (p == null)
            {
                return 0;
            }
            return (int)p.GetValue(record);
        }

        public static Team GetOppTeam (this object record)
        {
            FieldInfo p = record.GetType().GetField("VsTeam");
            return Team.FindTeamByName((string)p.GetValue(record));
        }

        internal static double PointsPerMetric(Metric m, double value)
        {
            if (m == Metric.PtsVs)
            {
                if (value == 0)
                {
                    return 10;
                }
                if (value <= 6)
                {
                    return 7;
                }
                if (value <= 13)
                {
                    return 4;
                }
                if (value <= 20)
                {
                    return 1;
                }
                if (value <= 27)
                {
                    return 0;
                }
                if (value <= 34)
                {
                    return -1;
                }
                return -4;
            }
            else
            {
                return points[m] * value;
            }
        }
    }
    enum Position
    {
        QB,
        RB,
        WR,
        TE,
        K,
        DEF
    }

    enum Metric
    {
        PassYds,
        PassTD,
        PassInt,
        RushYds,
        RushTD,
        Rec,
        RecYds,
        RecTD,
        RetTD,
        TwoPt,
        FumLost,
        FG19,
        FG29,
        FG39,
        FG49,
        FG50,
        PAT,
        PtsVs,
        Sack,
        DefInt,
        FumRec,
        DefTD,
        Safe,
        BlkKick,
        DefRetTD
    }
}
