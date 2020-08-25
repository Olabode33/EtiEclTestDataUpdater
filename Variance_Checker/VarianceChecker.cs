using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Text;

namespace Variance_Checker
{
    class VarianceChecker
    {
        string path = "";
        public void ExtractEADFromExcel()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            path = (Environment.CurrentDirectory + @"\InputAndFramework\EAD.xlsx");
            using (var package = new ExcelPackage(new FileInfo(path)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                var rows = new[] { 9, 12, 13, 15, 16 }; 
                var result = new List<string>();
                for (int i = 0; i < rows.Length; i++)
                {
                    result.Add(worksheet.Cells[rows[i], 5].Value.ToString());
                }
            }
        }

        public void ExtractImpairmentFromExcel()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            path = (Environment.CurrentDirectory + @"\InputAndFramework\Framework.xlsx");
            using (var package = new ExcelPackage(new FileInfo(path)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[2];

                var result = new List<string>();
                var rows = new [] {18, 19, 20, 21, 26, 27, 29, 30, 33, 34, 35, 36, 37, 38, 39, 40, 42, 43};

                for (int i = 0; i < rows.Length; i++)
                {
                    result.Add(worksheet.Cells[rows[i], 4].Value.ToString());
                }

                var rank = new List<string>();
                for (int i = 0; i < 20; i++)
                {
                    rank.Add(worksheet.Cells[i + 24, 9].Value.ToString());
                }
            }

        }

        public void ExtractLGDFromExcel()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            path = (Environment.CurrentDirectory + @"\InputAndFramework\LGD.xlsx");
            using (var package = new ExcelPackage(new FileInfo(path)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[4];

                var unsecuredRec = new List<dynamic>();
                var rows = new[] { 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17 };

                for (int i = 0; i < rows.Length; i++)
                {
                    unsecuredRec.Add(new DynamicClass
                    {
                        c = worksheet.Cells[rows[i], 3].Value.ToString(),
                        d = worksheet.Cells[rows[i], 4].Value.ToString(),
                        e = worksheet.Cells[rows[i], 5].Value.ToString(),
                        f = worksheet.Cells[rows[i], 6].Value.ToString(),
                        g = worksheet.Cells[rows[i], 7].Value.ToString(),
                        h = worksheet.Cells[rows[i], 8].Value.ToString(),
                        i = worksheet.Cells[rows[i], 9].Value.ToString(),
                        j = worksheet.Cells[rows[i], 10].Value.ToString(),
                        k = worksheet.Cells[rows[i], 11].Value.ToString(),
                        l = worksheet.Cells[rows[i], 12].Value.ToString(),
                        m = worksheet.Cells[rows[i], 13].Value.ToString(),

                    });
                }

                var pdAssumptions = new List<string>();
                rows = new[] { 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35 };

                for (int i = 0; i < rows.Length; i++)
                {
                    pdAssumptions.Add(worksheet.Cells[rows[i], 3].Value.ToString());
                }

                var costOfRecovery = new List<dynamic>();
                rows = new[] { 39, 40 };

                for (int i = 0; i < rows.Length; i++)
                {
                    costOfRecovery.Add(new DynamicClass
                    {
                        c = worksheet.Cells[rows[i], 3].Value.ToString(),
                        d = worksheet.Cells[rows[i], 4].Value.ToString(),
                        e = worksheet.Cells[rows[i], 5].Value.ToString(),
                        f = worksheet.Cells[rows[i], 6].Value.ToString(),
                        g = worksheet.Cells[rows[i], 7].Value.ToString(),
                        h = worksheet.Cells[rows[i], 8].Value.ToString(),
                        i = worksheet.Cells[rows[i], 9].Value.ToString(),
                        j = worksheet.Cells[rows[i], 10].Value.ToString(),
                        k = worksheet.Cells[rows[i], 11].Value.ToString(),
                        l = worksheet.Cells[rows[i], 12].Value.ToString()

                    });
                }

                var mes = new List<dynamic>();
                rows = new[] { 45, 46, 47 };

                for (int i = 0; i < rows.Length; i++)
                {
                    mes.Add(new DynamicClass
                    {
                        c = worksheet.Cells[rows[i], 3].Value.ToString(),
                        d = worksheet.Cells[rows[i], 4].Value.ToString(),
                        e = worksheet.Cells[rows[i], 5].Value.ToString(),
                        f = worksheet.Cells[rows[i], 6].Value.ToString(),
                        g = worksheet.Cells[rows[i], 7].Value.ToString(),
                        h = worksheet.Cells[rows[i], 8].Value.ToString(),
                        i = worksheet.Cells[rows[i], 9].Value.ToString(),
                        j = worksheet.Cells[rows[i], 10].Value.ToString(),
                        k = worksheet.Cells[rows[i], 11].Value.ToString()
                    });
                }

                var coltype = new List<string>();
                rows = new[] { 51, 52, 53, 54, 55, 56, 57, 58, 59 };

                for (int i = 0; i < rows.Length; i++)
                {
                    coltype.Add(worksheet.Cells[rows[i], 3].Value.ToString());
                }

                worksheet = package.Workbook.Worksheets[6];

                var haircuts = new List<dynamic>();
                rows = new[] { 4 };

                for (int i = 0; i < rows.Length; i++)
                {
                    haircuts.Add(new DynamicClass
                    {
                        b = worksheet.Cells[rows[i], 2].Value.ToString(),
                        c = worksheet.Cells[rows[i], 3].Value.ToString(),
                        d = worksheet.Cells[rows[i], 4].Value.ToString(),
                        e = worksheet.Cells[rows[i], 5].Value.ToString(),
                        f = worksheet.Cells[rows[i], 6].Value.ToString(),
                        g = worksheet.Cells[rows[i], 7].Value.ToString(),
                        h = worksheet.Cells[rows[i], 8].Value.ToString(),
                        i = worksheet.Cells[rows[i], 9].Value.ToString(),
                        j = worksheet.Cells[rows[i], 10].Value.ToString(),
                        k = worksheet.Cells[rows[i], 11].Value.ToString(),
                        l = worksheet.Cells[rows[i], 12].Value.ToString(),
                        m = worksheet.Cells[rows[i], 13].Value.ToString()
                    });
                }

                worksheet = package.Workbook.Worksheets[7];

                var bes = new List<dynamic>();

                for (int i = 4; i <= 4; i++)
                {
                    bes.Add(new DynamicClass
                    {
                        b = worksheet.Cells[i, 2].Value.ToString(),
                        c = worksheet.Cells[i, 3].Value.ToString(),
                        d = worksheet.Cells[i, 4].Value.ToString(),
                        e = worksheet.Cells[i, 5].Value.ToString(),
                        f = worksheet.Cells[i, 6].Value.ToString(),
                        g = worksheet.Cells[i, 7].Value.ToString(),
                        h = worksheet.Cells[i, 8].Value.ToString(),
                        i = worksheet.Cells[i, 9].Value.ToString(),
                        j = worksheet.Cells[i, 10].Value.ToString(),
                        k = worksheet.Cells[i, 11].Value.ToString(),
                        l = worksheet.Cells[i, 12].Value.ToString(),
                        m = worksheet.Cells[i, 13].Value.ToString(),
                        n = worksheet.Cells[i, 14].Value.ToString()

                    });
                }

                worksheet = package.Workbook.Worksheets[8];

                var oes = new List<dynamic>();

                for (int i = 4; i <= 4; i++)
                {
                    oes.Add(new DynamicClass
                    {
                        b = worksheet.Cells[i, 2].Value.ToString(),
                        c = worksheet.Cells[i, 3].Value.ToString(),
                        d = worksheet.Cells[i, 4].Value.ToString(),
                        e = worksheet.Cells[i, 5].Value.ToString(),
                        f = worksheet.Cells[i, 6].Value.ToString(),
                        g = worksheet.Cells[i, 7].Value.ToString(),
                        h = worksheet.Cells[i, 8].Value.ToString(),
                        i = worksheet.Cells[i, 9].Value.ToString(),
                        j = worksheet.Cells[i, 10].Value.ToString(),
                        k = worksheet.Cells[i, 11].Value.ToString(),
                        l = worksheet.Cells[i, 12].Value.ToString(),
                        m = worksheet.Cells[i, 13].Value.ToString(),
                        n = worksheet.Cells[i, 14].Value.ToString()

                    });
                    
                }

                worksheet = package.Workbook.Worksheets[9];

                var des = new List<dynamic>();

                for (int i = 4; i <= 4; i++)
                {
                    des.Add(new DynamicClass
                    {
                        b = worksheet.Cells[i, 2].Value.ToString(),
                        c = worksheet.Cells[i, 3].Value.ToString(),
                        d = worksheet.Cells[i, 4].Value.ToString(),
                        e = worksheet.Cells[i, 5].Value.ToString(),
                        f = worksheet.Cells[i, 6].Value.ToString(),
                        g = worksheet.Cells[i, 7].Value.ToString(),
                        h = worksheet.Cells[i, 8].Value.ToString(),
                        i = worksheet.Cells[i, 9].Value.ToString(),
                        j = worksheet.Cells[i, 10].Value.ToString(),
                        k = worksheet.Cells[i, 11].Value.ToString(),
                        l = worksheet.Cells[i, 12].Value.ToString(),
                        m = worksheet.Cells[i, 13].Value.ToString(),
                        n = worksheet.Cells[i, 14].Value.ToString()

                    });
                }

            }

        }

        public void PopulateExcel()
        {

        }
    }
}
