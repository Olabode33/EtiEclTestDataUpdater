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
        }
    }
}
