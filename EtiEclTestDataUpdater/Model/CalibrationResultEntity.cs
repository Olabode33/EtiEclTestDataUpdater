using System;
using System.Collections.Generic;
using System.Text;

namespace EtiEclTestDataUpdater.Model
{
    public class CalibrationResultBehaviouralTerm
    {
        public long AffiliateId { get; set; }
        public double Expired { get; set; }
        public double NonExpired { get; set; }
    }

    public class CalibrationResultCcfSummary
    {
        public long AffiliateId { get; set; }
        public double OdCcf { get; set; }
        public double CardCcf { get; set; }
        public double OverallCcf { get; set; }
    }
}
