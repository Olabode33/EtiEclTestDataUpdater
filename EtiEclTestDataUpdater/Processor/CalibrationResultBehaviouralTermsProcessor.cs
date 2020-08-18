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
    public class CalibrationResultBehaviouralTermsProcessor
    {
        private readonly DataAccess _dataAccess;

        public CalibrationResultBehaviouralTermsProcessor(DataAccess dataAccess)
        {
            _dataAccess = dataAccess;
        }

        public void Execute()
        {
            var dataList = ReadFromExcel();
            var qryBuilder = new StringBuilder();

            foreach (var item in dataList)
            {
                var qry = Queries.UpdateCalibrationEadBehaviouralTerm(item) + "\n";
                qryBuilder.Append(qry);
            }

            _dataAccess.ExecuteQuery(qryBuilder.ToString());

        }

        public List<CalibrationResultBehaviouralTerm> ReadFromExcel()
        {
            var dataList = new List<CalibrationResultBehaviouralTerm>();
            var filePath = $"{Path.Combine(_dataAccess.GetFilePath(), "AssumptionTemplate.xlsx")}";

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[1];
                int rows = worksheet.Dimension.Rows;

                for (int i = 2; i < rows; i++)
                {
                    var AffiliateId = worksheet.Cells[i, 1].Value;
                    var Expired = worksheet.Cells[i, 2].Value;
                    var NonExpired = worksheet.Cells[i, 3].Value;

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
                        var data = new CalibrationResultBehaviouralTerm();
                        try { data.AffiliateId = Convert.ToInt64(AffiliateId); } catch { data.AffiliateId = -1; }
                        try { data.Expired = Convert.ToDouble(Expired); } catch { data.Expired = 0.0; }
                        try { data.NonExpired = Convert.ToDouble(NonExpired); } catch { data.NonExpired = 0.0; }

                        dataList.Add(data);
                    }
                }
            }

            return dataList;
        }
    }
}
