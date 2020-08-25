using EtiEclTestDataUpdater.Data;
using EtiEclTestDataUpdater.Model;
using Microsoft.Extensions.FileSystemGlobbing.Internal.PatternContexts;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;

namespace EtiEclTestDataUpdater.Processor
{
    public class VarianceProcessor
    {
        public long _affiliateId;
        private DataAccess _dataAccess;
        private ExcelWorkbook _workBook;

        public VarianceProcessor(long affiliateId)
        {
            _affiliateId = affiliateId;
            _dataAccess = new DataAccess();
        }

        public void Run()
        {
            //Open File
            var filePath = $"{Path.Combine(_dataAccess.GetVarianceTemplateFileFolder(), "Variance_Checker_Template.xlsx")}";
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var package = new ExcelPackage(new FileInfo(filePath));
            _workBook = package.Workbook;

            //Run Processor
            ExtractCalibrationEadBehaviouralTerm();
            ExtractCalibrationEadCcfSummary();
            ExtractCalibrationLgdHaircut();
            ExtractCalibrationLgdRecoveryRate();
            ExtractCalibrationPd12Months();
            ExtractCalibrationPdCrDrSummary();


            //Close File
            package.Save();
            package.Dispose();
        }

        private void ExtractCalibrationEadBehaviouralTerm()
        {
            Console.WriteLine("Extracting Calibration Result EAD Behavioural Term...");
            var qry = Queries.GetCalibrationResultEadBehaviouralTerms(_affiliateId);
            var dataTable = _dataAccess.GetDataFromQuery(qry);

            if (dataTable.Rows.Count == 0)
            {
                Console.WriteLine("No records found.");
                return;
            }

            DataRow dataRow = dataTable.Rows[0];
            var item = _dataAccess.ParseDataToObject(new CalibrationResultBehaviouralTermForExtraction(), dataRow);

            ExcelWorksheet worksheet = _workBook.Worksheets[0]; //EAD_Input Sheet

            worksheet.Cells[7, 3].Value = item.Expired;
            worksheet.Cells[6, 3].Value = item.NonExpired;
        }

        private void ExtractCalibrationEadCcfSummary()
        {
            Console.WriteLine("Extracting Calibration Result EAD CCF Summary...");
            var qry = Queries.GetCalibrationResultEadCcfSummary(_affiliateId);
            var dataTable = _dataAccess.GetDataFromQuery(qry);

            if (dataTable.Rows.Count == 0)
            {
                Console.WriteLine("No records found.");
                return;
            }

            DataRow dataRow = dataTable.Rows[0];
            var item = _dataAccess.ParseDataToObject(new CalibrationResultCcfSummary(), dataRow);

            ExcelWorksheet worksheet = _workBook.Worksheets[0]; //EAD_Input Sheet

            for (int i = 8; i < 11; i++)
            {
                worksheet.Cells[i, 3].Value = item.OverallCcf;
            }
        }

        private void ExtractCalibrationLgdHaircut()
        {
            Console.WriteLine("Extracting Calibration Result LGD Haircut...");
            var qry = Queries.GetCalibrationResultLgdHaircut(_affiliateId);
            var dataTable = _dataAccess.GetDataFromQuery(qry);

            if (dataTable.Rows.Count == 0)
            {
                Console.WriteLine("No records found.");
                return;
            }

            DataRow dataRow = dataTable.Rows[0];
            var item = _dataAccess.ParseDataToObject(new CalibrationResultHaircutSummary(), dataRow);

            ExcelWorksheet worksheet = _workBook.Worksheets[1]; //LGD-Assumption Sheet

            worksheet.Cells[75, 1].Value = item.Debenture;
            worksheet.Cells[75, 2].Value = item.Cash;
            worksheet.Cells[75, 3].Value = item.Inventory;
            worksheet.Cells[75, 4].Value = item.PlantEquipment;
            worksheet.Cells[75, 5].Value = item.ResidentialProperty;
            worksheet.Cells[75, 6].Value = item.CommercialProperty;
            worksheet.Cells[75, 7].Value = item.Receivables;
            worksheet.Cells[75, 8].Value = item.Shares;
            worksheet.Cells[75, 9].Value = item.Vehicle;
        }

        private void ExtractCalibrationLgdRecoveryRate()
        {
            Console.WriteLine("Extracting Calibration Result LGD Recovery Rate...");
            var qry = Queries.GetCalibrationResultLgdRecoveryRate(_affiliateId);
            var dataTable = _dataAccess.GetDataFromQuery(qry);

            if (dataTable.Rows.Count == 0)
            {
                Console.WriteLine("No records found.");
                return;
            }

            DataRow dataRow = dataTable.Rows[0];
            var item = _dataAccess.ParseDataToObject(new CalibrationResultRecoveryRate(), dataRow);

            ExcelWorksheet worksheet = _workBook.Worksheets[1]; //LGD-Assumption Sheet

            //Commercial
            for (int i = 3; i < 8; i++)
            {
                worksheet.Cells[i, 8].Value = item.CommercialRecoveryRate;
            }
            //Consumer
            for (int i = 8; i < 11; i++)
            {
                worksheet.Cells[i, 8].Value = item.ConsumerRecoveryRate;
            }
            //Corporate
            for (int i = 11; i < 16; i++)
            {
                worksheet.Cells[i, 8].Value = item.CorporateRecoveryRate;
            }
        }

        private void ExtractCalibrationPd12Months()
        {
            Console.WriteLine("Extracting Calibration Result PD 12 Month Rating...");
            var qry = Queries.GetCalibrationResultPd12Months(_affiliateId);
            var dataTable = _dataAccess.GetDataFromQuery(qry);

            if (dataTable.Rows.Count == 0)
            {
                Console.WriteLine("No records found.");
                return;
            }

            if (dataTable.Rows.Count != 10)
            {
                Console.WriteLine("Record not equal to 10.");
            }


            ExcelWorksheet worksheet = _workBook.Worksheets[1]; //LGD-Assumption Sheet
            ExcelWorksheet pdWorksheet = _workBook.Worksheets[2]; //PD_Assumption Sheet

            for (int i = 0; i < 10; i++)
            {
                DataRow dataRow = dataTable.Rows[i];
                var item = _dataAccess.ParseDataToObject(new CalibrationResultPd12Month(), dataRow);
                worksheet.Cells[i + 20, 3].Value = item.Pd12Months;
                pdWorksheet.Cells[i + 8, 5].Value = item.Pd12Months;
            }
        }

        private void ExtractCalibrationPdCrDrSummary()
        {
            Console.WriteLine("Extracting Calibration Result PD CR DR Summary...");
            var qry = Queries.GetCalibrationResultPdCrDrSummary(_affiliateId);
            var dataTable = _dataAccess.GetDataFromQuery(qry);

            if (dataTable.Rows.Count == 0)
            {
                Console.WriteLine("No records found.");
                return;
            }


            DataRow dataRow = dataTable.Rows[0];
            var item = _dataAccess.ParseDataToObject(new CalibrationResultPdCrDr(), dataRow);

            ExcelWorksheet worksheet = _workBook.Worksheets[1]; //LGD-Assumption Sheet
            ExcelWorksheet pdWorksheet = _workBook.Worksheets[2]; //PD_Assumption Sheet

            for (int i = 3; i < 16; i++)
            {
                worksheet.Cells[i, 13].Value = item.CureRate;
            }
            pdWorksheet.Cells[2, 4].Value = item.RedefaultFactor;
        }
    }
}
