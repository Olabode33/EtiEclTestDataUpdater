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
    public class AssumptionPdSnpProcessor
    {
        private readonly DataAccess _dataAccess;

        public AssumptionPdSnpProcessor(DataAccess dataAccess)
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
                    var qry = Queries.UpdatePdSnPAssumption(item) + "\n";
                    qryBuilder.Append(qry);
                }

                _dataAccess.ExecuteQuery(qryBuilder.ToString());
                Console.WriteLine("Completed: " + this.GetType().Name);
            }

        }

        public List<PdSnPAssumption> ReadFromExcel()
        {
            var dataList = new List<PdSnPAssumption>();
            var filePath = $"{Path.Combine(_dataAccess.GetFilePath(), "AssumptionTemplate.xlsx")}";

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[12]; //PD_SandP Sheet
                int rows = worksheet.Dimension.Rows;

                for (int i = 2; i <= rows; i++)
                {
                    var AffiliateId = worksheet.Cells[i, 1].Value;
                    var creditRating = worksheet.Cells[i, 2].Value;
                    var etiCreditPolicy = worksheet.Cells[i, 3].Value;
                    var bestFit = worksheet.Cells[i, 4].Value;

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
                        var data_etiCreditPolicy = new PdSnPAssumption();
                        data_etiCreditPolicy.PdGroup = 2;
                        try { data_etiCreditPolicy.AffiliateId = Convert.ToInt64(AffiliateId); } catch { data_etiCreditPolicy.AffiliateId = -1; }
                        try { data_etiCreditPolicy.CreditRating = creditRating.ToString(); } catch { data_etiCreditPolicy.CreditRating = ""; }
                        try { data_etiCreditPolicy.PdRating = etiCreditPolicy.ToString(); } catch { data_etiCreditPolicy.PdRating = ""; }
                        dataList.Add(data_etiCreditPolicy);

                        var data_bestFit = new PdSnPAssumption();
                        data_bestFit.PdGroup = 3;
                        try { data_bestFit.AffiliateId = Convert.ToInt64(AffiliateId); } catch { data_bestFit.AffiliateId = -1; }
                        try { data_bestFit.CreditRating = creditRating.ToString(); } catch { data_bestFit.CreditRating = ""; }
                        try { data_bestFit.PdRating = bestFit.ToString(); } catch { data_bestFit.PdRating = ""; }
                        dataList.Add(data_bestFit);
                    }
                }
            }

            return dataList;
        }
    }
}
