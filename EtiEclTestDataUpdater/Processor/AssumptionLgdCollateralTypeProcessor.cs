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
    public class AssumptionLgdCollateralTypeProcessor
    {
        private readonly DataAccess _dataAccess;

        public AssumptionLgdCollateralTypeProcessor(DataAccess dataAccess)
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
                    var qry = Queries.UpdateLgdCollateralTypeAssumption(item) + "\n";
                    qryBuilder.Append(qry);
                }

                _dataAccess.ExecuteQuery(qryBuilder.ToString());
                Console.WriteLine("Completed: " + this.GetType().Name);
            }

        }

        public List<GeneralAssumptionEntity> ReadFromExcel()
        {
            var dataList = new List<GeneralAssumptionEntity>();
            var filePath = $"{Path.Combine(_dataAccess.GetFilePath(), "AssumptionTemplate.xlsx")}";

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[11]; //CollateralType Sheet
                int rows = worksheet.Dimension.Rows;

                for (int i = 2; i <= rows; i++)
                {
                    var AffiliateId = worksheet.Cells[i, 1].Value;
                    var collateralType = worksheet.Cells[i, 2].Value;
                    var ttrYears = worksheet.Cells[i, 3].Value;

                    if (AffiliateId == null)
                    {
                        //Console.WriteLine("Row is empty: " + i.ToString());
                    }
                    else if(string.IsNullOrWhiteSpace(AffiliateId.ToString()) || string.IsNullOrWhiteSpace(collateralType.ToString()))
                    {
                        //Console.WriteLine("Row is empty: " + i.ToString());
                    }
                    else
                    {
                        var data = new GeneralAssumptionEntity();
                        try { data.AffiliateId = Convert.ToInt64(AffiliateId); } catch { data.AffiliateId = -1; }
                        try { 
                            data.Key = collateralType.ToString().ToLower().Contains("plant") ? "Plant & Equipment" : 
                                        (collateralType.ToString().ToLower().Contains("residential") ? "Residential Property" :
                                        (collateralType.ToString().ToLower().Contains("commercial") ? "Commercial Property" : collateralType.ToString())); 
                        } catch { data.Key = ""; }
                        try { data.Value = Convert.ToDouble(ttrYears); } catch { data.Value = 0.0; }
                        dataList.Add(data);
                    }
                }
            }

            return dataList;
        }
    }
}
