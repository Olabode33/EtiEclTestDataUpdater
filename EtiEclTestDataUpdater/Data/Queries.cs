using EtiEclTestDataUpdater.Model;
using System;
using System.Collections.Generic;
using System.Text;

namespace EtiEclTestDataUpdater.Data
{
    public static class Queries
    {
        #region Update Calibration Result Query
        public static string UpdateCalibrationEadBehaviouralTerm(CalibrationResultBehaviouralTerm input)
        {
            return $"Update [CalibrationResult_EAD_Behavioural_Terms] " +
                   $" set Assumption_Expired = {input.Expired}, " +
                   $" Assumption_NonExpired = {input.NonExpired} " +
                   $" where CalibrationId = (select top 1 id from CalibrationRunEadBehaviouralTerms where OrganizationUnitId = {input.AffiliateId} and [Status] = 7);";
        }

        public static string UpdateCalibrationEadCcfSummary(CalibrationResultCcfSummary input)
        {
            return $"Update [CalibrationResult_EAD_CCF_Summary] " +
                   $" set [OD_CCF] = {input.OdCcf}, " +
                   $" [Card_CCF] = {input.CardCcf}, " +
                   $" [Overall_CCF] = {input.OverallCcf} " +
                   $" where CalibrationId = (select top 1 id from CalibrationRunEadCcfSummary where OrganizationUnitId = {input.AffiliateId} and [Status] = 7);";
        }

        public static string UpdateCalibrationHaircutSummary(CalibrationResultHaircutSummary input)
        {
            return $"Update [CalibrationResult_LGD_HairCut_Summary] " +
                   $" set [Debenture] = {input.Debenture}, [Cash] = {input.Cash}, [Inventory] = {input.Inventory}, [Plant_And_Equipment] = {input.PlantEquipment}, " +
                   $" [Residential_Property] = {input.ResidentialProperty}, [Commercial_Property] = {input.CommercialProperty}, [Receivables] = {input.Receivables}, " +
                   $" [Shares] = {input.Shares}, [Vehicle] = {input.Vehicle} " +
                   $" where CalibrationId = (select top 1 id from CalibrationRunLgdHairCut where OrganizationUnitId = {input.AffiliateId} and [Status] = 7);";
        }

        public static string UpdateCalibrationRecoveryRate(CalibrationResultRecoveryRate input)
        {
            return $"Update [CalibrationResult_LGD_RecoveryRate] " +
                   $" set [Corporate_RecoveryRate] = {input.CorporateRecoveryRate}, " +
                   $" [Commercial_RecoveryRate] = {input.CommercialRecoveryRate}, " +
                   $" [Consumer_RecoveryRate] = {input.ConsumerRecoveryRate} " +
                   $" where CalibrationId = (select top 1 id from CalibrationRunLgdRecoveryRate where OrganizationUnitId = {input.AffiliateId} and [Status] = 7);";
        }

        public static string UpdateCalibrationPdCrDr(CalibrationResultPdCrDr input)
        {
            return $"Update [CalibrationResult_PD_12Months_Summary] " +
                   $" set [Cure_Rate] = {input.CureRate}, " +
                   $" [Redefault_Rate] = {input.RedefaultRate}, " +
                   $" [Redefault_Factor] = {input.RedefaultFactor} " +
                   $" where CalibrationId = (select top 1 id from CalibrationRunPdCrDrs where OrganizationUnitId = {input.AffiliateId} and [Status] = 7);";
        }

        public static string UpdateCalibrationPdCommsConsMarginalDefaultRate(CalibrationResultMarginalDefaultRate input)
        {
            return $"Update [CalibrationResult_PD_CommsCons_MarginalDefaultRate] " +
                   $" set  " +
                   $" [Comm1] = {input.Comm1}, " +
                   $" [Cons1] = {input.Cons1}, " +
                   $" [Comm2] = {input.Comm2}, " +
                   $" [Cons2] = {input.Cons2} " +
                   $" where [Month] = {input.Month} and  CalibrationId = (select top 1 id from CalibrationRunPdCrDrs where OrganizationUnitId = {input.AffiliateId} and [Status] = 7);";
        }

        public static string UpdateCalibrationPd12Month(CalibrationResultPd12Month input)
        {
            return $"Update [CalibrationResult_PD_12Months] " +
                   $" set [Months_PDs_12] = {input.Pd12Months} " +
                   $" where [Rating] = {input.Rating} and  CalibrationId = (select top 1 id from CalibrationRunPdCrDrs where OrganizationUnitId = {input.AffiliateId} and [Status] = 7);";
        }
        #endregion

        #region Assumption Update Query
        public static string UpdateEadInputAssumptionByKey(GeneralAssumptionEntity input)
        {
            return $"Update [EadInputAssumptions] " +
                   $" set [Value] = {input.Value} " +
                   $" where [Key] = '{input.Key}' and [OrganizationUnitId] = {input.AffiliateId};";
        }

        public static string UpdateLgdCorAssumption(LgdCorAssumption input)
        {
            var collateralValueKey = input.LgdGroup == 3 ? "CostOfRecoveryHighCollateralValue" : "CostOfRecoveryLowCollateralValue";
            return $"Update [LgdInputAssumptions]  set [Value] = {input.CollateralValue}  where [LgdGroup] = {input.LgdGroup} and [Key] = '{collateralValueKey}' and [OrganizationUnitId] = {input.AffiliateId}; " + 
                   $"Update [LgdInputAssumptions]  set [Value] = {input.Debenture}  where [LgdGroup] = {input.LgdGroup} and [InputName] = 'Debenture' and [OrganizationUnitId] = {input.AffiliateId}; " + 
                   $"Update [LgdInputAssumptions]  set [Value] = {input.Cash}  where [LgdGroup] = {input.LgdGroup} and [InputName] = 'Cash' and [OrganizationUnitId] = {input.AffiliateId}; " + 
                   $"Update [LgdInputAssumptions]  set [Value] = {input.Inventory}  where [LgdGroup] = {input.LgdGroup} and [InputName] = 'Inventory' and [OrganizationUnitId] = {input.AffiliateId}; " + 
                   $"Update [LgdInputAssumptions]  set [Value] = {input.PlantEquipment}  where [LgdGroup] = {input.LgdGroup} and [InputName] = 'Plant & Equipment' and [OrganizationUnitId] = {input.AffiliateId}; " + 
                   $"Update [LgdInputAssumptions]  set [Value] = {input.ResidentialProperty}  where [LgdGroup] = {input.LgdGroup} and [InputName] = 'Residential Property' and [OrganizationUnitId] = {input.AffiliateId}; " + 
                   $"Update [LgdInputAssumptions]  set [Value] = {input.CommercialProperty}  where [LgdGroup] = {input.LgdGroup} and [InputName] = 'Commercial Property' and [OrganizationUnitId] = {input.AffiliateId}; " + 
                   $"Update [LgdInputAssumptions]  set [Value] = {input.Receivables}  where [LgdGroup] = {input.LgdGroup} and [InputName] = 'Receivables' and [OrganizationUnitId] = {input.AffiliateId}; " + 
                   $"Update [LgdInputAssumptions]  set [Value] = {input.Shares}  where [LgdGroup] = {input.LgdGroup} and [InputName] = 'Shares' and [OrganizationUnitId] = {input.AffiliateId}; " + 
                   $"Update [LgdInputAssumptions]  set [Value] = {input.Vehicle}  where [LgdGroup] = {input.LgdGroup} and [InputName] = 'Vehicle' and [OrganizationUnitId] = {input.AffiliateId}; " + 
                   $"\n";
        }

        public static string UpdateLgdCollateralGrowthAssumption(LgdCorAssumption input)
        {
            return $"Update [LgdInputAssumptions]  set [Value] = {input.Debenture}  where [LgdGroup] = {input.LgdGroup} and [InputName] = 'Debenture' and [OrganizationUnitId] = {input.AffiliateId}; " + 
                   $"Update [LgdInputAssumptions]  set [Value] = {input.Cash}  where [LgdGroup] = {input.LgdGroup} and [InputName] = 'Cash' and [OrganizationUnitId] = {input.AffiliateId}; " + 
                   $"Update [LgdInputAssumptions]  set [Value] = {input.Inventory}  where [LgdGroup] = {input.LgdGroup} and [InputName] = 'Inventory' and [OrganizationUnitId] = {input.AffiliateId}; " + 
                   $"Update [LgdInputAssumptions]  set [Value] = {input.PlantEquipment}  where [LgdGroup] = {input.LgdGroup} and [InputName] = 'Plant & Equipment' and [OrganizationUnitId] = {input.AffiliateId}; " + 
                   $"Update [LgdInputAssumptions]  set [Value] = {input.ResidentialProperty}  where [LgdGroup] = {input.LgdGroup} and [InputName] = 'Residential Property' and [OrganizationUnitId] = {input.AffiliateId}; " + 
                   $"Update [LgdInputAssumptions]  set [Value] = {input.CommercialProperty}  where [LgdGroup] = {input.LgdGroup} and [InputName] = 'Commercial Property' and [OrganizationUnitId] = {input.AffiliateId}; " + 
                   $"Update [LgdInputAssumptions]  set [Value] = {input.Receivables}  where [LgdGroup] = {input.LgdGroup} and [InputName] = 'Receivables' and [OrganizationUnitId] = {input.AffiliateId}; " + 
                   $"Update [LgdInputAssumptions]  set [Value] = {input.Shares}  where [LgdGroup] = {input.LgdGroup} and [InputName] = 'Shares' and [OrganizationUnitId] = {input.AffiliateId}; " + 
                   $"Update [LgdInputAssumptions]  set [Value] = {input.Vehicle}  where [LgdGroup] = {input.LgdGroup} and [InputName] = 'Vehicle' and [OrganizationUnitId] = {input.AffiliateId}; " + 
                   $"\n";
        }

        public static string UpdateLgdCollateralTypeAssumption(GeneralAssumptionEntity input)
        {
            return $"Update [LgdInputAssumptions]  set [Value] = {input.Value}  where [LgdGroup] = 8 and [InputName] = '{input.Key}' and [OrganizationUnitId] = {input.AffiliateId}; ";
        }

        public static string UpdatePdSnPAssumption(PdSnPAssumption input)
        {
            return $"Update [PdInputAssumptions]  set [Value] = '{input.PdRating}'  where [PdGroup] = {input.PdGroup} and [InputName] = '{input.CreditRating}' and [OrganizationUnitId] = {input.AffiliateId}; ";
        }

        public static string UpdatePdSnPCummulativeDefaultRatesAssumption(PdSnPCummulativeDefaultRatesAssumption input)
        {
            return $"Update [PdInputAssumptionSnPCummulativeDefaultRates] set [Value] = '{input.Value}'  where [Rating] = '{input.Rating}' and [Years] = '{input.Year}' and [OrganizationUnitId] = {input.AffiliateId}; ";
        }

        public static string UpdateImpairmentScenarioInputAssumption(ImpairmentCreditRating input)
        {
            return $"Update [Assumptions] set [Value] = '{input.Value}'  where [InputName] = '{input.Key}' and [OrganizationUnitId] = {input.AffiliateId}; ";
        }

        public static string UpdateImpairmentCreditRatingAssumption(ImpairmentCreditRating input)
        {
            return $"Update [Assumptions] set [Value] = '{input.Value}'  where [InputName] = 'Credit Rating Rank {input.Key}' and [OrganizationUnitId] = {input.AffiliateId}; ";
        }
        #endregion
    }
}
