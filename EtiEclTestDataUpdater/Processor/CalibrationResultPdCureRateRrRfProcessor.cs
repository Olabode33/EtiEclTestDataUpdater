using EtiEclTestDataUpdater.Data;
using EtiEclTestDataUpdater.Model;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace EtiEclTestDataUpdater.Processor
{
    public class CalibrationResultPdCureRateRrRfProcessor
    {
        private readonly DataAccess _dataAccess;

        public CalibrationResultPdCureRateRrRfProcessor(DataAccess dataAccess)
        {
            _dataAccess = dataAccess;
        }

        public void Execute()
        {
            var dataList = ReadFromExcel();
            var qryBuilder = new StringBuilder();

            if (dataList.Count > 0)
            {
                foreach (var item in dataList)
                {
                    var qry = Queries.UpdateCalibrationPdCrDr(item) + "\n";
                    qryBuilder.Append(qry);
                }

                _dataAccess.ExecuteQuery(qryBuilder.ToString());
                Console.WriteLine("Completed: " + this.GetType().Name);
            }

        }

        public List<CalibrationResultPdCrDr> ReadFromExcel()
        {
            var dataList = new List<CalibrationResultPdCrDr>();
            var filePath = $"{Path.Combine(_dataAccess.GetFilePath(), "AssumptionTemplate.xlsx")}";

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[6]; //Calibration_CureRate_RR_RF  Sheet
                int rows = worksheet.Dimension.Rows;

                for (int i = 2; i <= rows; i++)
                {
                    var AffiliateId = worksheet.Cells[i, 1].Value;
                    var cureRate = worksheet.Cells[i, 2].Value;
                    var redefaultRate = worksheet.Cells[i, 3].Value;
                    var redefaultFactor = worksheet.Cells[i, 4].Value;

                    if (AffiliateId == null)
                    {
                        //Console.WriteLine("Row is empty: " + i.ToString());
                    }
                    else if(string.IsNullOrWhiteSpace(AffiliateId.ToString()))
                    {
                        //Console.WriteLine("Row is empty: " + i.ToString());
                    }
                    else
                    {
                        var data = new CalibrationResultPdCrDr();
                        try { data.AffiliateId = Convert.ToInt64(AffiliateId); } catch { data.AffiliateId = -1; }
                        try { data.CureRate = Convert.ToDouble(cureRate); } catch { data.CureRate = 0.0; }
                        try { data.RedefaultRate = Convert.ToDouble(redefaultRate); } catch { data.RedefaultRate = 0.0; }
                        try { data.RedefaultFactor = Convert.ToDouble(redefaultFactor); } catch { data.RedefaultFactor = 0.0; }

                        dataList.Add(data);
                    }
                }
            }

            return dataList;
        }
    }
}
