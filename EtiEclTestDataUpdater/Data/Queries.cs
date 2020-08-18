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
    }
}
