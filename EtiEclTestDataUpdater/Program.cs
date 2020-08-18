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

            CoreProcessor coreProcessor = new CoreProcessor();
            coreProcessor.Run();
        }
    }
}
