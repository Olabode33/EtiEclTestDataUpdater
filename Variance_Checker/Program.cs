﻿using System;

namespace Variance_Checker
{
    class Program
    {
        static void Main(string[] args)
        {
            VarianceChecker varianceChecker = new VarianceChecker();
            varianceChecker.ExtractEADFromExcel();
            varianceChecker.ExtractImpairmentFromExcel();
            varianceChecker.ExtractLGDFromExcel();
            varianceChecker.ExtractPDFromExcel();
            varianceChecker.ExtractPDMacroFromExcel();
            varianceChecker.ExtractPDMagDefFromExcel();

        }
    }
}
