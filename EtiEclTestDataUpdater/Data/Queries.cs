using EtiEclTestDataUpdater.Model;
using System;
using System.Collections.Generic;
using System.Text;

namespace EtiEclTestDataUpdater.Data
{
    public static class Queries
    {
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
    }
}
