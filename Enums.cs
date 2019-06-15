using System.Reflection;

namespace DraftSystem
{
    internal static class Ext
    {
        public static int GetMetricValue(this object record, Metric metric)
        {
            FieldInfo p = record.GetType().GetField(metric.ToString());
            if (p == null)
            {
                return 0;
            }
            return (int)p.GetValue(record);
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
