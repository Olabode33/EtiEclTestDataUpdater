using EtiEclTestDataUpdater.Data;
using System;
using System.Collections.Generic;
using System.Text;

namespace EtiEclTestDataUpdater.Processor
{
    public class CoreProcessor
    {
        private CalibrationResultBehaviouralTermsProcessor _calibrationResultBehaviouralTermsProcessor;
        private CalibrationResultCcfSummaryProcessor _calibrationResultCcfSummaryProcessor;
        private DataAccess _dataAccess;

        public CoreProcessor()
        {
            _dataAccess = new DataAccess();
            _calibrationResultBehaviouralTermsProcessor = new CalibrationResultBehaviouralTermsProcessor(_dataAccess);
            _calibrationResultCcfSummaryProcessor = new CalibrationResultCcfSummaryProcessor(_dataAccess);
        }

        public void Run()
        {
            _calibrationResultBehaviouralTermsProcessor.Execute();
            _calibrationResultCcfSummaryProcessor.Execute();
        }
    }
}
