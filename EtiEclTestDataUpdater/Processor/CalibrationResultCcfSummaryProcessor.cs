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
    public class CalibrationResultCcfSummaryProcessor
    {
        private readonly DataAccess _dataAccess;

        public CalibrationResultCcfSummaryProcessor(DataAccess dataAccess)
        {
            _dataAccess = dataAccess;
        }

        public void Execute()
        {
            var dataList = ReadFromExcel();
            var qryBuilder = new StringBuilder();

            foreach (var item in dataList)
            {
                var qry = Queries.UpdateCalibrationEadCcfSummary(item) + "\n";
                qryBuilder.Append(qry);
            }

            _dataAccess.ExecuteQuery(qryBuilder.ToString());

        }

        public List<CalibrationResultCcfSummary> ReadFromExcel()
        {
            var dataList = new List<CalibrationResultCcfSummary>();
            var filePath = $"{Path.Combine(_dataAccess.GetFilePath(), "AssumptionTemplate.xlsx")}";

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[2]; //CCF Sheet
                int rows = worksheet.Dimension.Rows;

                for (int i = 2; i < rows; i++)
                {
                    var AffiliateId = worksheet.Cells[i, 1].Value;
                    var od = worksheet.Cells[i, 2].Value;
                    var card = worksheet.Cells[i, 3].Value;
                    var overall = worksheet.Cells[i, 4].Value;

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
                        var data = new CalibrationResultCcfSummary();
                        try { data.AffiliateId = Convert.ToInt64(AffiliateId); } catch { data.AffiliateId = -1; }
                        try { data.OdCcf = Convert.ToDouble(od); } catch { data.OdCcf = 0.0; }
                        try { data.CardCcf = Convert.ToDouble(card); } catch { data.CardCcf = 0.0; }
                        try { data.OverallCcf = Convert.ToDouble(overall); } catch { data.OverallCcf = 0.0; }

                        dataList.Add(data);
                    }
                }
            }

            return dataList;
        }
    }
}
