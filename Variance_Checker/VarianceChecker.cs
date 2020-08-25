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
        string output = Environment.CurrentDirectory + @"\Template\Variance_Checker_Template.xlsx";
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

                PopulateEAD(result);
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
                var rows = new[] { 18, 19, 20, 21, 26, 27, 29, 30, 33, 34, 35, 36, 37, 38, 39, 40, 42, 43 };

                for (int i = 0; i < rows.Length; i++)
                {
                    result.Add(worksheet.Cells[rows[i], 4].Value.ToString());
                }

                var rank = new List<string>();
                for (int i = 0; i < 20; i++)
                {
                    rank.Add(worksheet.Cells[i + 24, 9].Value.ToString());
                }

                PopulateImpairment(result, rank);

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

                    });
                }

                PopulateLGD(unsecuredRec, pdAssumptions, costOfRecovery, mes, coltype, haircuts, bes, oes, des);

            }

        }

        public void ExtractPDFromExcel()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            path = (Environment.CurrentDirectory + @"\InputAndFramework\PD.xlsx");
            using (var package = new ExcelPackage(new FileInfo(path)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                var rows = new[] { 12, 13 };
                var result = new List<string>();

                for (int i = 0; i < rows.Length; i++)
                {
                    result.Add(worksheet.Cells[rows[i], 5].Value.ToString());
                }

                worksheet = package.Workbook.Worksheets[5];
                var twelveMPD = new List<dynamic>();
                rows = new[] { 4, 5, 6, 7, 8, 9, 10, 11, 12, 13 };

                for (int i = 0; i < rows.Length; i++)
                {
                    twelveMPD.Add(new DynamicClass
                    {
                        b = worksheet.Cells[rows[i], 2].Value.ToString(),
                        c = worksheet.Cells[rows[i], 3].Value.ToString(),
                        d = worksheet.Cells[rows[i], 4].Value.ToString(),
                        e = worksheet.Cells[rows[i], 5].Value.ToString(),
                    });
                }

                var spCum = new List<dynamic>();
                rows = new[] { 5, 6, 7, 8, 9, 10, 11 };

                for (int i = 0; i < rows.Length; i++)
                {
                    spCum.Add(new DynamicClass
                    {
                        g = worksheet.Cells[rows[i], 7].Value.ToString(),
                        h = worksheet.Cells[rows[i], 8].Value.ToString(),
                        i = worksheet.Cells[rows[i], 9].Value.ToString(),
                        j = worksheet.Cells[rows[i], 10].Value.ToString(),
                        k = worksheet.Cells[rows[i], 11].Value.ToString(),
                        l = worksheet.Cells[rows[i], 12].Value.ToString(),
                        m = worksheet.Cells[rows[i], 13].Value.ToString(),
                        n = worksheet.Cells[rows[i], 14].Value.ToString(),
                        o = worksheet.Cells[rows[i], 15].Value.ToString(),
                        p = worksheet.Cells[rows[i], 16].Value.ToString(),
                        q = worksheet.Cells[rows[i], 17].Value.ToString(),
                        r = worksheet.Cells[rows[i], 18].Value.ToString(),
                        s = worksheet.Cells[rows[i], 19].Value.ToString(),
                        t = worksheet.Cells[rows[i], 20].Value.ToString(),
                        u = worksheet.Cells[rows[i], 21].Value.ToString(),
                        v = worksheet.Cells[rows[i], 22].Value.ToString(),
                        w = worksheet.Cells[rows[i], 23].Value.ToString(),
                    });
                }

                PopulatePD(result, twelveMPD, spCum);
            }
        }

        public void ExtractPDMagDefFromExcel()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            path = (Environment.CurrentDirectory + @"\InputAndFramework\PD.xlsx");
            using (var package = new ExcelPackage(new FileInfo(path)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[6];
                var mdr = new List<dynamic>();

                for (int i = 5; i < 245; i++)
                {
                    mdr.Add(new DynamicClass
                    {
                        c = worksheet.Cells[i, 3].Value.ToString(),
                        d = worksheet.Cells[i, 4].Value.ToString(),
                        e = worksheet.Cells[i, 5].Value.ToString(),
                        f = worksheet.Cells[i, 6].Value.ToString(),
                    });
                }

                PopulatePDMagDef(mdr);
            }
        }

        public void ExtractPDMacroFromExcel()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            path = (Environment.CurrentDirectory + @"\InputAndFramework\PD.xlsx");
            using (var package = new ExcelPackage(new FileInfo(path)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[7];
                var hi = new List<dynamic>();

                for (int i = 4; i < 36; i++)
                {
                    hi.Add(new DynamicClass
                    {
                        b = worksheet.Cells[i, 2].Value.ToString(),
                        c = worksheet.Cells[i, 3].Value.ToString(),
                        d = worksheet.Cells[i, 4].Value.ToString(),

                    });
                }

                var npl = new List<dynamic>();

                for (int i = 3; i < 35; i++)
                {
                    npl.Add(new DynamicClass
                    {
                        f = worksheet.Cells[i, 6].Value.ToString(),
                        g = worksheet.Cells[i, 7].Value.ToString(),
                    });
                }

                var stats = new List<string>();

                stats.Add(worksheet.Cells[5, 9].Value.ToString());
                stats.Add(worksheet.Cells[5, 10].Value.ToString());
                stats.Add(worksheet.Cells[7, 10].Value.ToString());
                stats.Add(worksheet.Cells[8, 10].Value.ToString());

                var si = new List<dynamic>();
                for (int i = 39; i < 44; i++)
                {
                    si.Add(new DynamicClass
                    {
                        c = worksheet.Cells[i, 3].Value.ToString(),
                        d = worksheet.Cells[i, 4].Value.ToString(),
                        e = worksheet.Cells[i, 5].Value.ToString(),

                    });
                }

                var bemp = new List<dynamic>();
                for (int i = 47; i < 81; i++)
                {
                    bemp.Add(new DynamicClass
                    {
                        b = worksheet.Cells[i, 2].Value.ToString(),
                        c = worksheet.Cells[i, 3].Value.ToString(),
                        d = worksheet.Cells[i, 4].Value.ToString(),
                        e = worksheet.Cells[i, 5].Value.ToString(),

                    });
                }

                var omp = new List<dynamic>();
                for (int i = 84; i < 118; i++)
                {
                    omp.Add(new DynamicClass
                    {
                        b = worksheet.Cells[i, 2].Value.ToString(),
                        c = worksheet.Cells[i, 3].Value.ToString(),
                        d = worksheet.Cells[i, 4].Value.ToString(),
                        e = worksheet.Cells[i, 5].Value.ToString(),

                    });
                }

                var dmp = new List<dynamic>();
                for (int i = 121; i < 155; i++)
                {
                    dmp.Add(new DynamicClass
                    {
                        b = worksheet.Cells[i, 2].Value.ToString(),
                        c = worksheet.Cells[i, 3].Value.ToString(),
                        d = worksheet.Cells[i, 4].Value.ToString(),
                        e = worksheet.Cells[i, 5].Value.ToString(),

                    });
                }

                PopulatePDMacro(hi, npl, stats, si, bemp, omp, dmp);
            }
        }

        public void PopulateEAD(List<string> ead)
        {
            using (var package = new ExcelPackage(new FileInfo(output)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                for (int i = 3; i < 8; i++)
                {
                    worksheet.Cells[i, 2].Value = ead[i - 3];
                }
                package.Save();
            }
        }

        public void PopulateImpairment(List<string> result, List<string> rank)
        {
            using (var package = new ExcelPackage(new FileInfo(output)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[5];
                var rows = new[] { 2, 3, 4, 5, 10, 11, 13, 14, 17, 18, 19, 20, 21, 22, 23, 24, 26, 27 };

                for (int i = 2; i < rows.Length + 2; i++)
                {
                    worksheet.Cells[rows[i - 2], 2].Value = result[i - 2];
                }

                for (int i = 31; i < 51; i++)
                {
                    worksheet.Cells[i, 2].Value = rank[i - 31];
                }

                package.Save();
            }
        }
        public void PopulateLGD(List<dynamic> unsecuredRec, List<string> pdAssumptions, List<dynamic> costOfRecovery, List<dynamic> mes, List<string> coltype, List<dynamic> haircuts, List<dynamic> bes, List<dynamic> oes, List<dynamic> des)
        {
            using (var package = new ExcelPackage(new FileInfo(output)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[1];
                for (int i = 3; i < 16; i++)
                {
                    worksheet.Cells[i, 2].Value = unsecuredRec[i - 3].c;
                    worksheet.Cells[i, 3].Value = unsecuredRec[i - 3].d;
                    worksheet.Cells[i, 4].Value = unsecuredRec[i - 3].e;
                    worksheet.Cells[i, 5].Value = unsecuredRec[i - 3].f;
                    worksheet.Cells[i, 6].Value = unsecuredRec[i - 3].g;
                    worksheet.Cells[i, 7].Value = unsecuredRec[i - 3].h;
                    worksheet.Cells[i, 8].Value = unsecuredRec[i - 3].i;
                    worksheet.Cells[i, 9].Value = unsecuredRec[i - 3].j;
                    worksheet.Cells[i, 10].Value = unsecuredRec[i - 3].k;
                    worksheet.Cells[i, 11].Value = unsecuredRec[i - 3].l;
                    worksheet.Cells[i, 12].Value = unsecuredRec[i - 3].m;
                }

                for (int i = 20; i < 35; i++)
                {
                    worksheet.Cells[i, 2].Value = pdAssumptions[i - 20];
                }

                for (int i = 39; i < 41; i++)
                {
                    worksheet.Cells[i, 2].Value = costOfRecovery[i - 39].c;
                    worksheet.Cells[i, 3].Value = costOfRecovery[i - 39].d;
                    worksheet.Cells[i, 4].Value = costOfRecovery[i - 39].e;
                    worksheet.Cells[i, 5].Value = costOfRecovery[i - 39].f;
                    worksheet.Cells[i, 6].Value = costOfRecovery[i - 39].g;
                    worksheet.Cells[i, 7].Value = costOfRecovery[i - 39].h;
                    worksheet.Cells[i, 8].Value = costOfRecovery[i - 39].i;
                    worksheet.Cells[i, 9].Value = costOfRecovery[i - 39].j;
                    worksheet.Cells[i, 10].Value = costOfRecovery[i - 39].k;
                    worksheet.Cells[i, 11].Value = costOfRecovery[i - 39].l;
                }

                for (int i = 48; i < 51; i++)
                {
                    worksheet.Cells[i, 2].Value = mes[i - 48].c;
                    worksheet.Cells[i, 3].Value = mes[i - 48].d;
                    worksheet.Cells[i, 4].Value = mes[i - 48].e;
                    worksheet.Cells[i, 5].Value = mes[i - 48].f;
                    worksheet.Cells[i, 6].Value = mes[i - 48].g;
                    worksheet.Cells[i, 7].Value = mes[i - 48].h;
                    worksheet.Cells[i, 8].Value = mes[i - 48].i;
                    worksheet.Cells[i, 9].Value = mes[i - 48].j;
                    worksheet.Cells[i, 10].Value = mes[i - 48].k;
                    worksheet.Cells[i, 11].Value = mes[i - 48].l;
                }

                for (int i = 61; i < 70; i++)
                {
                    worksheet.Cells[i, 2].Value = coltype[i - 61];
                }

                worksheet.Cells[74, 1].Value = haircuts[0].b;
                worksheet.Cells[74, 2].Value = haircuts[0].c;
                worksheet.Cells[74, 3].Value = haircuts[0].d;
                worksheet.Cells[74, 4].Value = haircuts[0].e;
                worksheet.Cells[74, 5].Value = haircuts[0].f;
                worksheet.Cells[74, 6].Value = haircuts[0].g;
                worksheet.Cells[74, 7].Value = haircuts[0].h;
                worksheet.Cells[74, 8].Value = haircuts[0].i;
                worksheet.Cells[74, 9].Value = haircuts[0].j;


                worksheet.Cells[81, 1].Value = bes[0].b;
                worksheet.Cells[81, 2].Value = bes[0].c;
                worksheet.Cells[81, 3].Value = bes[0].d;
                worksheet.Cells[81, 4].Value = bes[0].e;
                worksheet.Cells[81, 5].Value = bes[0].f;
                worksheet.Cells[81, 6].Value = bes[0].g;
                worksheet.Cells[81, 7].Value = bes[0].h;
                worksheet.Cells[81, 8].Value = bes[0].i;
                worksheet.Cells[81, 9].Value = bes[0].j;
                worksheet.Cells[81, 10].Value = bes[0].k;

                worksheet.Cells[88, 1].Value = oes[0].b;
                worksheet.Cells[88, 2].Value = oes[0].c;
                worksheet.Cells[88, 3].Value = oes[0].d;
                worksheet.Cells[88, 4].Value = oes[0].e;
                worksheet.Cells[88, 5].Value = oes[0].f;
                worksheet.Cells[88, 6].Value = oes[0].g;
                worksheet.Cells[88, 7].Value = oes[0].h;
                worksheet.Cells[88, 8].Value = oes[0].i;
                worksheet.Cells[88, 9].Value = oes[0].j;
                worksheet.Cells[88, 10].Value = oes[0].k;

                worksheet.Cells[95, 1].Value = des[0].b;
                worksheet.Cells[95, 2].Value = des[0].c;
                worksheet.Cells[95, 3].Value = des[0].d;
                worksheet.Cells[95, 4].Value = des[0].e;
                worksheet.Cells[95, 5].Value = des[0].f;
                worksheet.Cells[95, 6].Value = des[0].g;
                worksheet.Cells[95, 7].Value = des[0].h;
                worksheet.Cells[95, 8].Value = des[0].i;
                worksheet.Cells[95, 9].Value = des[0].j;
                worksheet.Cells[95, 10].Value = des[0].k;


                package.Save();
            }
        }

        public void PopulatePD(List<string> result, List<dynamic> twelveMPD, List<dynamic> spCum)
        {
            using (var package = new ExcelPackage(new FileInfo(output)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[2];
                for (int i = 2; i < 4; i++)
                {
                    worksheet.Cells[i, 3].Value = result[i - 2];
                }

                for (int i = 8; i < 18; i++)
                {
                    worksheet.Cells[i, 2].Value = twelveMPD[i - 8].b;
                    worksheet.Cells[i, 3].Value = twelveMPD[i - 8].c;
                    worksheet.Cells[i, 4].Value = twelveMPD[i - 8].d;
                    worksheet.Cells[i, 5].Value = twelveMPD[i - 8].e;
                }

                for (int i = 23; i < 30; i++)
                {
                    worksheet.Cells[i, 1].Value = spCum[i - 23].g;
                    worksheet.Cells[i, 2].Value = spCum[i - 23].h;
                    worksheet.Cells[i, 3].Value = spCum[i - 23].i;
                    worksheet.Cells[i, 4].Value = spCum[i - 23].j; 
                    worksheet.Cells[i, 5].Value = spCum[i - 23].k;
                    worksheet.Cells[i, 6].Value = spCum[i - 23].l;
                    worksheet.Cells[i, 7].Value = spCum[i - 23].m;
                    worksheet.Cells[i, 8].Value = spCum[i - 23].n; 
                    worksheet.Cells[i, 9].Value = spCum[i - 23].o;
                    worksheet.Cells[i, 10].Value = spCum[i - 23].p;
                    worksheet.Cells[i, 11].Value = spCum[i - 23].q;
                    worksheet.Cells[i, 12].Value = spCum[i - 23].r;
                    worksheet.Cells[i, 13].Value = spCum[i - 23].s;
                    worksheet.Cells[i, 14].Value = spCum[i - 23].t;
                    worksheet.Cells[i, 15].Value = spCum[i - 23].u;
                    worksheet.Cells[i, 16].Value = spCum[i - 23].v;
                    worksheet.Cells[i, 17].Value = spCum[i - 23].w;
                }

                package.Save();
            }
        }

        public void PopulatePDMagDef(List<dynamic> result)
        {
            using (var package = new ExcelPackage(new FileInfo(output)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[3];
                for (int i = 4; i < 244; i++)
                {
                    worksheet.Cells[i, 2].Value = result[i - 4].c;
                    worksheet.Cells[i, 3].Value = result[i - 4].d;
                    worksheet.Cells[i, 4].Value = result[i - 4].e;
                    worksheet.Cells[i, 5].Value = result[i - 4].f;
                }
                package.Save();
            }
        }

        public void PopulatePDMacro(List<dynamic> hi, List<dynamic> npl, List<string> stats, List<dynamic> si, List<dynamic> bemp, List<dynamic> omp, List<dynamic> dmp)
        {
            using (var package = new ExcelPackage(new FileInfo(output)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[4];
                for (int i = 3; i < 35; i++)
                {
                    worksheet.Cells[i, 1].Value = hi[i - 3].b;
                    worksheet.Cells[i, 2].Value = hi[i - 3].c;
                    worksheet.Cells[i, 3].Value = hi[i - 3].d;
                }

                for (int i = 3; i < 35; i++)
                {
                    worksheet.Cells[i, 13].Value = npl[i - 3].f;
                    worksheet.Cells[i, 14].Value = npl[i - 3].g;
                }
               
                    worksheet.Cells[38, 2].Value = stats[0];
                    worksheet.Cells[39, 2].Value = stats[1];
                    worksheet.Cells[40, 3].Value = stats[2];
                    worksheet.Cells[41, 3].Value = stats[3];
               


                for (int i = 46; i < 51; i++)
                {
                    worksheet.Cells[i, 2].Value = si[i - 46].c;
                    worksheet.Cells[i, 3].Value = si[i - 46].d;
                    worksheet.Cells[i, 4].Value = si[i - 46].e;
                }


                for (int i = 57; i < 91; i++)
                {
                    worksheet.Cells[i, 1].Value = bemp[i - 57].b;
                    worksheet.Cells[i, 2].Value = bemp[i - 57].c;
                    worksheet.Cells[i, 3].Value = bemp[i - 57].d;
                    worksheet.Cells[i, 4].Value = bemp[i - 57].e;
                }


                for (int i = 96; i < 130; i++)
                {
                    worksheet.Cells[i, 1].Value = omp[i - 96].b;
                    worksheet.Cells[i, 2].Value = omp[i - 96].c;
                    worksheet.Cells[i, 3].Value = omp[i - 96].d;
                    worksheet.Cells[i, 4].Value = omp[i - 96].e;
                }

                for (int i = 135; i < 169; i++)
                {
                    worksheet.Cells[i, 1].Value = dmp[i - 135].b;
                    worksheet.Cells[i, 2].Value = dmp[i - 135].c;
                    worksheet.Cells[i, 3].Value = dmp[i - 135].d;
                    worksheet.Cells[i, 4].Value = dmp[i - 135].e;
                }

                package.Save();
            }
        }
    }
}