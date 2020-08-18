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
    public class CalibrationResultHaircutProcessor
    {
        private readonly DataAccess _dataAccess;

        public CalibrationResultHaircutProcessor(DataAccess dataAccess)
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
                    var qry = Queries.UpdateCalibrationHaircutSummary(item) + "\n";
                    qryBuilder.Append(qry);
                }

                _dataAccess.ExecuteQuery(qryBuilder.ToString());
            }

        }

        public List<CalibrationResultHaircutSummary> ReadFromExcel()
        {
            var dataList = new List<CalibrationResultHaircutSummary>();
            var filePath = $"{Path.Combine(_dataAccess.GetFilePath(), "AssumptionTemplate.xlsx")}";

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[3]; //Haircut Sheet
                int rows = worksheet.Dimension.Rows;

                for (int i = 2; i < rows; i++)
                {
                    var AffiliateId = worksheet.Cells[i, 1].Value;
                    var debenture = worksheet.Cells[i, 2].Value;
                    var cash = worksheet.Cells[i, 3].Value;
                    var inventory = worksheet.Cells[i, 4].Value;
                    var plant = worksheet.Cells[i, 5].Value;
                    var residential = worksheet.Cells[i, 6].Value;
                    var commercial = worksheet.Cells[i, 7].Value;
                    var recievable = worksheet.Cells[i, 8].Value;
                    var shares = worksheet.Cells[i, 9].Value;
                    var vehicle = worksheet.Cells[i, 10].Value;

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
                        var data = new CalibrationResultHaircutSummary();
                        try { data.AffiliateId = Convert.ToInt64(AffiliateId); } catch { data.AffiliateId = -1; }
                        try { data.Debenture = Convert.ToDouble(debenture); } catch { data.Debenture = 0.0; }
                        try { data.Cash = Convert.ToDouble(cash); } catch { data.Cash = 0.0; }
                        try { data.Inventory = Convert.ToDouble(inventory); } catch { data.Inventory = 0.0; }
                        try { data.PlantEquipment = Convert.ToDouble(plant); } catch { data.PlantEquipment = 0.0; }
                        try { data.ResidentialProperty = Convert.ToDouble(residential); } catch { data.ResidentialProperty = 0.0; }
                        try { data.CommercialProperty = Convert.ToDouble(commercial); } catch { data.CommercialProperty = 0.0; }
                        try { data.Receivables = Convert.ToDouble(recievable); } catch { data.Receivables = 0.0; }
                        try { data.Shares = Convert.ToDouble(shares); } catch { data.Shares = 0.0; }
                        try { data.Vehicle = Convert.ToDouble(vehicle); } catch { data.Vehicle = 0.0; }

                        dataList.Add(data);
                    }
                }
            }

            return dataList;
        }
    }
}
