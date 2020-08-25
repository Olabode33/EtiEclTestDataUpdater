using EtiEclTestDataUpdater.Processor;
using Microsoft.Extensions.Configuration;
using System;
using System.IO;

namespace EtiEclTestDataUpdater
{
    class Program
    {
        static void Main(string[] args)
        {
            //Console.WriteLine("Override Calibration and Assumption? \nEnter 'Y' for Yes and 'N' for No:");
            //string coreResponse = Console.ReadLine();
            //if (coreResponse == "Y")
            //{
            //    Console.WriteLine("Starting Calibration and Assumption Override...");
            //    CoreProcessor coreProcessor = new CoreProcessor();
            //    coreProcessor.Run();
            //    Console.WriteLine("Completed Calibration and Assumption Override.");
            //}

            //Console.WriteLine("\nExtract Calibration Result and ECL Run Assumption? \nEnter 'Y' for Yes and 'N' for N:");
            //string varianceResponse = Console.ReadLine();
            //if (varianceResponse == "Y")
            //{
                Console.WriteLine("Starting Calibration and Assumption Extraction...");
                Console.WriteLine("Enter Affiliate Id:");
                long affiliateId = Convert.ToInt64(Console.ReadLine());
                VarianceProcessor varianceProcessor = new VarianceProcessor(affiliateId);
                varianceProcessor.Run();
                Console.WriteLine("Completed Calibration and Assumption Extraction.");
            //}
        }
    }
}
