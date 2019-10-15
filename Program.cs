using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PrestaCapExercise
{
    class Program
    {
        static void Main(string[] args)
        {
            if(args.Length != 2)
            {
                Console.WriteLine("Incorrect parameters.");
                return;
            }

            string pathToInputFile = args[0];
            string pathToOutputFile = args[1];

            Run(pathToInputFile, pathToOutputFile);
        }

        static public void Run(string input, string output)
        {
            try
            {
                HotelCollection hotelCollection = JsonConvert.DeserializeObject<HotelCollection>(File.ReadAllText(input));

                IReporter reporter = new ExcelReporter();

                byte[] stream = reporter.CreateReport(hotelCollection.HotelList);

                if (stream != null)
                    reporter.SaveReport(stream, output);
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}
