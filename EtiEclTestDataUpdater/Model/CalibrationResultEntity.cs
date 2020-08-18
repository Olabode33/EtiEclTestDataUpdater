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

    public class CalibrationResultHaircutSummary
    {
        public long AffiliateId { get; set; }
        public double Debenture { get; set; }
        public double Cash { get; set; }
        public double Inventory { get; set; }
        public double PlantEquipment { get; set; }
        public double ResidentialProperty { get; set; }
        public double CommercialProperty { get; set; }
        public double Receivables { get; set; }
        public double Shares { get; set; }
        public double Vehicle { get; set; }
    }

    public class CalibrationResultRecoveryRate
    {
        public long AffiliateId { get; set; }
        public double CommercialRecoveryRate { get; set; }
        public double ConsumerRecoveryRate { get; set; }
        public double CorporateRecoveryRate { get; set; }
    }

    public class CalibrationResultPdCrDr
    {
        public long AffiliateId { get; set; }
        public double CureRate { get; set; }
        public double RedefaultRate { get; set; }
        public double RedefaultFactor { get; set; }
        public double CommercialCureRate { get; set; }
        public double CommercialRedefaultRate { get; set; }
        public double ConsumerCureRate { get; set; }
        public double ConsumerRedefaultRate { get; set; }
    }

    public class CalibrationResultMarginalDefaultRate
    {
        public long AffiliateId { get; set; }
        public int Month { get; set; }
        public double Comm1 { get; set; }
        public double Cons1 { get; set; }
        public double Comm2 { get; set; }
        public double Cons2 { get; set; }
    }

    public class CalibrationResultPd12Month
    {
        public long AffiliateId { get; set; }
        public int Rating { get; set; }
        public double Pd12Months { get; set; }
    }
}
