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
    public class AssumptionImpairmentScenarioInputProcessor
    {
        private readonly DataAccess _dataAccess;

        public AssumptionImpairmentScenarioInputProcessor(DataAccess dataAccess)
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
                    var qry = Queries.UpdateImpairmentScenarioInputAssumption(item) + "\n";
                    qryBuilder.Append(qry);
                }

                _dataAccess.ExecuteQuery(qryBuilder.ToString());
                Console.WriteLine("Completed: " + this.GetType().Name);
            }

        }

        public List<ImpairmentCreditRating> ReadFromExcel()
        {
            var dataList = new List<ImpairmentCreditRating>();
            var filePath = $"{Path.Combine(_dataAccess.GetFilePath(), "AssumptionTemplate.xlsx")}";

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[14]; //Impairment_ScenarioInput Sheet
                int rows = worksheet.Dimension.Rows;

                for (int i = 2; i <= rows; i++)
                {
                    var AffiliateId = worksheet.Cells[i, 1].Value;
                    var inputName = worksheet.Cells[i, 2].Value;
                    var value = worksheet.Cells[i, 3].Value;

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
                        var data = new ImpairmentCreditRating();
                        try { data.AffiliateId = Convert.ToInt64(AffiliateId); } catch { data.AffiliateId = -1; }
                        try { data.Key = inputName.ToString(); } catch { data.Key = ""; }
                        try { data.Value = value.ToString(); } catch { data.Value = value.ToString(); }

                        dataList.Add(data);
                    }
                }
            }

            return dataList;
        }
    }
}
