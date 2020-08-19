using EtiEclTestDataUpdater.Data;
using System;
using System.Collections.Generic;
using System.Text;

namespace EtiEclTestDataUpdater.Processor
{
    public class CoreProcessor
    {
        private DataAccess _dataAccess;
        private CalibrationResultBehaviouralTermsProcessor _calibrationResultBehaviouralTermsProcessor;
        private CalibrationResultCcfSummaryProcessor _calibrationResultCcfSummaryProcessor;
        private CalibrationResultHaircutProcessor _calibrationResultHaircutProcessor;
        private CalibrationResultRecoveryRateProcessor _calibrationResultRecoveryRateProcessor;
        private CalibrationResultPdCureRateRrRfProcessor _calibrationResultPdCureRateRrRfProcessor;
        private CalibrationResultPdMarginalDefaultRateProcessor _calibrationResultPdMarginalDefaultRateProcessor;
        private CalibrationResultPd12MonthProcessor _calibrationResultPd12MonthProcessor;
        private AssumptionEadProcessor _assumptionEadProcessor;
        private AssumptionLgdCorProcessor _assumptionLgdCorProcessor;
        private AssumptionLgdCollateralGrowthProcessor _assumptionLgdCollateralGrowthProcessor;
        private AssumptionLgdCollateralTypeProcessor _assumptionLgdCollateralTypeProcessor;
        private AssumptionPdSnpProcessor _assumptionPdSnpProcessor;
        private AssumptionPdSnpCummulativeDefaultRatesProcessor _assumptionPdSnpCummulativeProcessor;
        private AssumptionImpairmentScenarioInputProcessor _assumptionImpairmentScenarioInputProcessor;
        private AssumptionImpairmentCreditRatingProcessor _assumptionImpairmentCreditRatingProcessor;

        public CoreProcessor()
        {
            _dataAccess = new DataAccess();
            _calibrationResultBehaviouralTermsProcessor = new CalibrationResultBehaviouralTermsProcessor(_dataAccess);
            _calibrationResultCcfSummaryProcessor = new CalibrationResultCcfSummaryProcessor(_dataAccess);
            _calibrationResultHaircutProcessor = new CalibrationResultHaircutProcessor(_dataAccess);
            _calibrationResultRecoveryRateProcessor = new CalibrationResultRecoveryRateProcessor(_dataAccess);
            _calibrationResultPdCureRateRrRfProcessor = new CalibrationResultPdCureRateRrRfProcessor(_dataAccess);
            _calibrationResultPdMarginalDefaultRateProcessor = new CalibrationResultPdMarginalDefaultRateProcessor(_dataAccess);
            _calibrationResultPd12MonthProcessor = new CalibrationResultPd12MonthProcessor(_dataAccess);
            _assumptionEadProcessor = new AssumptionEadProcessor(_dataAccess);
            _assumptionLgdCorProcessor = new AssumptionLgdCorProcessor(_dataAccess);
            _assumptionLgdCollateralGrowthProcessor = new AssumptionLgdCollateralGrowthProcessor(_dataAccess);
            _assumptionLgdCollateralTypeProcessor = new AssumptionLgdCollateralTypeProcessor(_dataAccess);
            _assumptionPdSnpProcessor = new AssumptionPdSnpProcessor(_dataAccess);
            _assumptionPdSnpCummulativeProcessor = new AssumptionPdSnpCummulativeDefaultRatesProcessor(_dataAccess);
            _assumptionImpairmentScenarioInputProcessor = new AssumptionImpairmentScenarioInputProcessor(_dataAccess);
            _assumptionImpairmentCreditRatingProcessor = new AssumptionImpairmentCreditRatingProcessor(_dataAccess);
        }

        public void Run()
        {
            _calibrationResultBehaviouralTermsProcessor.Execute();
            _calibrationResultCcfSummaryProcessor.Execute();
            _calibrationResultHaircutProcessor.Execute();
            _calibrationResultRecoveryRateProcessor.Execute();
            _calibrationResultPdCureRateRrRfProcessor.Execute();
            _calibrationResultPdMarginalDefaultRateProcessor.Execute();
            _calibrationResultPd12MonthProcessor.Execute();
            _assumptionEadProcessor.Execute();
            _assumptionLgdCorProcessor.Execute();
            _assumptionLgdCollateralGrowthProcessor.Execute();
            _assumptionLgdCollateralTypeProcessor.Execute();
            _assumptionPdSnpProcessor.Execute();
            _assumptionPdSnpCummulativeProcessor.Execute();
            _assumptionImpairmentScenarioInputProcessor.Execute();
            _assumptionImpairmentCreditRatingProcessor.Execute();

            Console.WriteLine("All Completed: " + this.GetType().Name);
        }
    }
}
