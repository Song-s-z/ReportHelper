using System;
using System.Configuration;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ConfigurationManager = System.Configuration.ConfigurationManager;


namespace Nlog
{
    class Program
    {
        static void Main(string[] args)
        {
            var path = ConfigurationManager.AppSettings["modelPath"];

            try
            {
                ExcelHelper.GenerateReport(path);
                Console.WriteLine( "Fale enerate sucefully!");
                Console.ReadKey();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                Console.ReadKey();
            }

            Console.ReadKey();
        }
    }
}
