using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PrestaCapExercise
{
    public class ExcelReporter : IReporter
    {
        /// <summary>
        /// creates a report based on json input
        /// </summary>
        /// <param name="hotels"></param>
        /// <returns>byte array of report</returns>
        public byte[] CreateReport(List<HotelEntry> hotels)
        {
            using (ExcelBuilder builder = new ExcelBuilder())
            {
                foreach (HotelEntry hotel in hotels)
                {
                    builder.AddSheet(hotel.hotel.name);
                    var sheet = builder.GetSheet(hotel.hotel.name);

                    sheet.AddHeader(new List<string>() { "ARRIVAL_DATE", "DEPARTURE_DATE", "PRICE", "CURRENCY", "RATE_NAME", "ADULTS", "BREAKFAST_INCLUDED" });

                    foreach (HotelRate rating in hotel.hotelRates)
                    {
                        List<object> values = new List<object>()
                        {
                            rating.targetDay,
                            rating.targetDay.AddDays(rating.los),
                            rating.price.numericFloat,
                            rating.price.currency,
                            rating.rateName,
                            rating.adults,
                            rating.rateTags.FirstOrDefault(tag => tag.name.Equals("breakfast"))?.shape
                        };

                        sheet.AddRow(values);
                    }
                }

                return builder.GetByteArray();
            }
        }

        /// <summary>
        /// specify how to save report
        /// </summary>
        /// <param name="stream"></param>
        /// <param name="outputPath"></param>
        public void SaveReport(byte[] stream, string outputPath)
        {
            // save byte stream as excel file
            File.WriteAllBytes(outputPath, stream);
        }
    }
}
