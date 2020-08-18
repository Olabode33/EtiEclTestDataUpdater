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
    public class CalibrationResultRecoveryRateProcessor
    {
        private readonly DataAccess _dataAccess;

        public CalibrationResultRecoveryRateProcessor(DataAccess dataAccess)
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
                    var qry = Queries.UpdateCalibrationRecoveryRate(item) + "\n";
                    qryBuilder.Append(qry);
                }

                _dataAccess.ExecuteQuery(qryBuilder.ToString());
            }

        }

        public List<CalibrationResultRecoveryRate> ReadFromExcel()
        {
            var dataList = new List<CalibrationResultRecoveryRate>();
            var filePath = $"{Path.Combine(_dataAccess.GetFilePath(), "AssumptionTemplate.xlsx")}";

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[4]; //Recovery Rate Sheet
                int rows = worksheet.Dimension.Rows;

                for (int i = 2; i < rows; i++)
                {
                    var AffiliateId = worksheet.Cells[i, 1].Value;
                    var commercial = worksheet.Cells[i, 2].Value;
                    var consumer = worksheet.Cells[i, 3].Value;
                    var corporate = worksheet.Cells[i, 4].Value;

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
                        var data = new CalibrationResultRecoveryRate();
                        try { data.AffiliateId = Convert.ToInt64(AffiliateId); } catch { data.AffiliateId = -1; }
                        try { data.CommercialRecoveryRate = Convert.ToDouble(commercial); } catch { data.CommercialRecoveryRate = 0.0; }
                        try { data.ConsumerRecoveryRate = Convert.ToDouble(consumer); } catch { data.ConsumerRecoveryRate = 0.0; }
                        try { data.CorporateRecoveryRate = Convert.ToDouble(corporate); } catch { data.CorporateRecoveryRate = 0.0; }

                        dataList.Add(data);
                    }
                }
            }

            return dataList;
        }
    }
}
