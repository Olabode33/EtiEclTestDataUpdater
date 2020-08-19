using System;
using System.Collections.Generic;
using System.Text;

namespace EtiEclTestDataUpdater.Model
{
    public class GeneralAssumptionEntity
    {
        public string Key { get; set; }
        public double Value { get; set; }
        public long AffiliateId { get; set; }
    }

    public class CollateralTypeBase
    {
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

    public class LgdCorAssumption: CollateralTypeBase
    {
        public long AffiliateId { get; set; }
        public int LgdGroup { get; set; }
        public double CollateralValue { get; set; }
    }

    public class PdSnPAssumption
    {

        public string CreditRating { get; set; }
        public string PdRating { get; set; }
        public long AffiliateId { get; set; }
        public int PdGroup { get; set; }
    }

    public class PdSnPCummulativeDefaultRatesAssumption
    {
        public string Rating { get; set; }
        public int Year { get; set; }
        public long AffiliateId { get; set; }
        public double Value { get; set; }
    }

    public class ImpairmentCreditRating
    {
        public string Key { get; set; }
        public string Value { get; set; }
        public long AffiliateId { get; set; }
    }


}
