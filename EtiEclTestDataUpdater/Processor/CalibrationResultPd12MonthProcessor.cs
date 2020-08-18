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
    public class CalibrationResultPd12MonthProcessor
    {
        private readonly DataAccess _dataAccess;

        public CalibrationResultPd12MonthProcessor(DataAccess dataAccess)
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
                    var qry = Queries.UpdateCalibrationPd12Month(item) + "\n";
                    qryBuilder.Append(qry);
                }

                _dataAccess.ExecuteQuery(qryBuilder.ToString());
            }

        }

        public List<CalibrationResultPd12Month> ReadFromExcel()
        {
            var dataList = new List<CalibrationResultPd12Month>();
            var filePath = $"{Path.Combine(_dataAccess.GetFilePath(), "AssumptionTemplate.xlsx")}";

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[7]; //Calibration_PD_CR_DR  Sheet
                int rows = worksheet.Dimension.Rows;

                for (int i = 2; i < rows; i++)
                {
                    var AffiliateId = worksheet.Cells[i, 1].Value;
                    
                    for (int rating = 1; rating <= 10; rating++)
                    {
                        var pd12Month = worksheet.Cells[i, rating + 1].Value;

                        if (AffiliateId == null)
                        {
                            Console.WriteLine("Row is empty: " + i.ToString());
                        }
                        else if (string.IsNullOrWhiteSpace(AffiliateId.ToString()))
                        {
                            Console.WriteLine("Row is empty: " + i.ToString());
                        }
                        else
                        {
                            var data = new CalibrationResultPd12Month();
                            try { data.AffiliateId = Convert.ToInt64(AffiliateId); } catch { data.AffiliateId = -1; }
                            data.Rating = rating;
                            try { data.Pd12Months = Convert.ToDouble(pd12Month); } catch { data.Pd12Months = 0.0; }

                            dataList.Add(data);
                        }
                    }
                }
            }

            return dataList;
        }
    }
}
