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
    public class CalibrationResultPdMarginalDefaultRateProcessor
    {
        private readonly DataAccess _dataAccess;

        public CalibrationResultPdMarginalDefaultRateProcessor(DataAccess dataAccess)
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
                    var qry = Queries.UpdateCalibrationPdCommsConsMarginalDefaultRate(item) + "\n";
                    qryBuilder.Append(qry);
                }

                _dataAccess.ExecuteQuery(qryBuilder.ToString());
            }

        }

        public List<CalibrationResultMarginalDefaultRate> ReadFromExcel()
        {
            var dataList = new List<CalibrationResultMarginalDefaultRate>();
            var filePath = $"{Path.Combine(_dataAccess.GetFilePath(), "AssumptionTemplate.xlsx")}";

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[6]; //Calibration_PD_MariginalDefault  Sheet
                int rows = worksheet.Dimension.Rows;

                for (int i = 2; i < rows; i++)
                {
                    var AffiliateId = worksheet.Cells[i, 1].Value;
                    var month = worksheet.Cells[i, 2].Value;
                    var cons1 = worksheet.Cells[i, 3].Value;
                    var cons2 = worksheet.Cells[i, 4].Value;
                    var comm1 = worksheet.Cells[i, 5].Value;
                    var comm2 = worksheet.Cells[i, 6].Value;

                    if (AffiliateId == null)
                    {
                        Console.WriteLine("Row is empty: " + i.ToString());
                    }
                    else if(string.IsNullOrWhiteSpace(AffiliateId.ToString()))
                    {
                        Console.WriteLine("Row is empty: " + i.ToString());
                    }
                    else
                    {
                        var data = new CalibrationResultMarginalDefaultRate();
                        try { data.AffiliateId = Convert.ToInt64(AffiliateId); } catch { data.AffiliateId = -1; }
                        try { data.Month = Convert.ToInt32(month); } catch { data.Month = 0; }
                        try { data.Comm1 = Convert.ToDouble(comm1); } catch { data.Comm1 = 0.0; }
                        try { data.Cons1 = Convert.ToDouble(cons1); } catch { data.Cons1 = 0.0; }
                        try { data.Comm2 = Convert.ToDouble(comm2); } catch { data.Comm2 = 0.0; }
                        try { data.Cons2 = Convert.ToDouble(cons2); } catch { data.Cons2 = 0.0; }

                        dataList.Add(data);
                    }
                }
            }

            return dataList;
        }
    }
}
