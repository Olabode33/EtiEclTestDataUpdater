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
    public class AssumptionPdSnpCummulativeDefaultRatesProcessor
    {
        private readonly DataAccess _dataAccess;

        public AssumptionPdSnpCummulativeDefaultRatesProcessor(DataAccess dataAccess)
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
                    var qry = Queries.UpdatePdSnPCummulativeDefaultRatesAssumption(item) + "\n";
                    qryBuilder.Append(qry);
                }

                _dataAccess.ExecuteQuery(qryBuilder.ToString());
                Console.WriteLine("Completed: " + this.GetType().Name);
            }

        }

        public List<PdSnPCummulativeDefaultRatesAssumption> ReadFromExcel()
        {
            var dataList = new List<PdSnPCummulativeDefaultRatesAssumption>();
            var filePath = $"{Path.Combine(_dataAccess.GetFilePath(), "AssumptionTemplate.xlsx")}";

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[13]; //PD_CummulativeRate Sheet
                int rows = worksheet.Dimension.Rows;

                for (int i = 2; i <= rows; i++)
                {
                    var AffiliateId = worksheet.Cells[i, 1].Value;
                    var rating = worksheet.Cells[i, 2].Value;

                    for (int year = 1; year <= 15; year++)
                    {
                        var value = worksheet.Cells[i, year + 2].Value;

                        if (AffiliateId == null)
                        {
                            //Console.WriteLine("Row is empty: " + i.ToString());
                        }
                        else if (string.IsNullOrWhiteSpace(AffiliateId.ToString()))
                        {
                            //Console.WriteLine("Row is empty: " + i.ToString());
                        }
                        else
                        {
                            var data = new PdSnPCummulativeDefaultRatesAssumption();
                            data.Rating = rating.ToString();
                            data.Year = year;
                            try { data.AffiliateId = Convert.ToInt64(AffiliateId); } catch { data.AffiliateId = -1; }
                            try { data.Value = Convert.ToDouble(value); } catch { data.Value = 0.0; }
                            dataList.Add(data);

                        }
                    }
                }
            }

            return dataList;
        }
    }
}
