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
    public class AssumptionEadProcessor
    {
        private readonly DataAccess _dataAccess;

        public AssumptionEadProcessor(DataAccess dataAccess)
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
                    var qry = Queries.UpdateEadInputAssumptionByKey(item) + "\n";
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
                ExcelWorksheet worksheet = package.Workbook.Worksheets[8]; //Assumption_EAD Sheet
                int rows = worksheet.Dimension.Rows;

                for (int i = 2; i <= rows; i++)
                {
                    var AffiliateId = worksheet.Cells[i, 1].Value;
                    var conversionFactorObe = worksheet.Cells[i, 2].Value;
                    var prepaymentFactor = worksheet.Cells[i, 3].Value;

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
                        var data_conversionFactorObe = new GeneralAssumptionEntity();
                        try { data_conversionFactorObe.AffiliateId = Convert.ToInt64(AffiliateId); } catch { data_conversionFactorObe.AffiliateId = -1; }
                        data_conversionFactorObe.Key = AssumptionKey.KEY_ConversionFactorOBE;
                        try { data_conversionFactorObe.Value = Convert.ToDouble(conversionFactorObe); } catch { data_conversionFactorObe.Value = 0.0; }
                        dataList.Add(data_conversionFactorObe);

                        var data_prepaymentFactor = new GeneralAssumptionEntity();
                        try { data_prepaymentFactor.AffiliateId = Convert.ToInt64(AffiliateId); } catch { data_prepaymentFactor.AffiliateId = -1; }
                        data_prepaymentFactor.Key = AssumptionKey.KEY_PrePaymentFactor;
                        try { data_prepaymentFactor.Value = Convert.ToDouble(prepaymentFactor); } catch { data_prepaymentFactor.Value = 0.0; }
                        dataList.Add(data_prepaymentFactor);
                    }
                }
            }

            return dataList;
        }
    }
}
