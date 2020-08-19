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
    public class AssumptionLgdCollateralGrowthProcessor
    {
        private readonly DataAccess _dataAccess;

        public AssumptionLgdCollateralGrowthProcessor(DataAccess dataAccess)
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
                    var qry = Queries.UpdateLgdCollateralGrowthAssumption(item) + "\n";
                    qryBuilder.Append(qry);
                }

                _dataAccess.ExecuteQuery(qryBuilder.ToString());
                Console.WriteLine("Completed: " + this.GetType().Name);
            }

        }

        public List<LgdCorAssumption> ReadFromExcel()
        {
            var dataList = new List<LgdCorAssumption>();
            var filePath = $"{Path.Combine(_dataAccess.GetFilePath(), "AssumptionTemplate.xlsx")}";

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[8]; //CollateralGrowth Sheet
                int rows = worksheet.Dimension.Rows;

                for (int i = 2; i <= rows; i++)
                {
                    var AffiliateId = worksheet.Cells[i, 1].Value;
                    var scenario = worksheet.Cells[i, 2].Value;
                    var debenture = worksheet.Cells[i, 3].Value;
                    var cash = worksheet.Cells[i, 4].Value;
                    var inventory = worksheet.Cells[i, 5].Value;
                    var plantEquipment = worksheet.Cells[i, 6].Value;
                    var residential = worksheet.Cells[i, 7].Value;
                    var commercial = worksheet.Cells[i, 8].Value;
                    var receivables = worksheet.Cells[i, 9].Value;
                    var shares = worksheet.Cells[i, 10].Value;
                    var vehicle = worksheet.Cells[i, 11].Value;

                    if (AffiliateId == null)
                    {
                        //Console.WriteLine("Row is empty: " + i.ToString());
                    }
                    else if(string.IsNullOrWhiteSpace(AffiliateId.ToString()) && string.IsNullOrWhiteSpace(scenario.ToString()) 
                                                                              && (scenario.ToString().ToLower().Contains("best") || scenario.ToString().ToLower().Contains("optimistic") || scenario.ToString().ToLower().Contains("downturn")))
                    {
                        //Console.WriteLine("Row is empty: " + i.ToString());
                    }
                    else
                    {
                        var data = new LgdCorAssumption();
                        try { data.AffiliateId = Convert.ToInt64(AffiliateId); } catch { data.AffiliateId = -1; }
                        data.LgdGroup = scenario.ToString().ToLower().Contains("best") ? 5 : (scenario.ToString().ToLower().Contains("optimistic") ? 6 : 7);
                        try { data.Debenture = Convert.ToDouble(debenture); } catch { data.Debenture = 0.0; }
                        try { data.Cash = Convert.ToDouble(cash); } catch { data.Cash = 0.0; }
                        try { data.Inventory = Convert.ToDouble(inventory); } catch { data.Inventory = 0.0; }
                        try { data.PlantEquipment = Convert.ToDouble(plantEquipment); } catch { data.PlantEquipment = 0.0; }
                        try { data.ResidentialProperty = Convert.ToDouble(residential); } catch { data.ResidentialProperty = 0.0; }
                        try { data.CommercialProperty = Convert.ToDouble(commercial); } catch { data.CommercialProperty = 0.0; }
                        try { data.Receivables = Convert.ToDouble(receivables); } catch { data.Receivables = 0.0; }
                        try { data.Shares = Convert.ToDouble(shares); } catch { data.Shares = 0.0; }
                        try { data.Vehicle = Convert.ToDouble(vehicle); } catch { data.Vehicle = 0.0; }

                        dataList.Add(data);

                        if (data.LgdGroup == 5)
                        {
                            var dataOptimitic = new LgdCorAssumption();
                            try { dataOptimitic.AffiliateId = Convert.ToInt64(AffiliateId); } catch { dataOptimitic.AffiliateId = -1; }
                            dataOptimitic.LgdGroup = 6;
                            dataOptimitic.Debenture = data.Debenture;
                            dataOptimitic.Cash = data.Cash;
                            dataOptimitic.Inventory = data.Inventory;
                            dataOptimitic.PlantEquipment = data.PlantEquipment;
                            dataOptimitic.ResidentialProperty = data.ResidentialProperty;
                            dataOptimitic.CommercialProperty = data.CommercialProperty;
                            dataOptimitic.Receivables = data.Receivables;
                            dataOptimitic.Shares = data.Shares;
                            dataOptimitic.Vehicle = data.Vehicle;

                            dataList.Add(dataOptimitic);

                            var dataDownturn = new LgdCorAssumption();
                            try { dataDownturn.AffiliateId = Convert.ToInt64(AffiliateId); } catch { dataDownturn.AffiliateId = -1; }
                            dataDownturn.LgdGroup = 7;
                            dataDownturn.Debenture = dataOptimitic.Debenture * 0.92 - 0.08;
                            dataDownturn.Cash = dataOptimitic.Cash * 0.92 - 0.08;
                            dataDownturn.Inventory = dataOptimitic.Inventory * 0.92 - 0.08;
                            dataDownturn.PlantEquipment = dataOptimitic.PlantEquipment * 0.92 - 0.08;
                            dataDownturn.ResidentialProperty = dataOptimitic.ResidentialProperty * 0.92 - 0.08;
                            dataDownturn.CommercialProperty = dataOptimitic.CommercialProperty * 0.92 - 0.08;
                            dataDownturn.Receivables = dataOptimitic.Receivables * 0.92 - 0.08;
                            dataDownturn.Shares = dataOptimitic.Shares * 0.92 - 0.08;
                            dataDownturn.Vehicle = dataOptimitic.Vehicle * 0.92 - 0.08;

                            dataList.Add(dataDownturn);
                        }

                    }
                }
            }

            return dataList;
        }
    }
}
