using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newtonsoft.Json;
using OfficeOpenXml;
using PrestaCapExercise;
using System;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;

namespace UnitTests
{
    [TestClass]
    public class ExcelBuilderTests
    {
        [TestMethod]
        public void Verify_Adding_Sheets()
        {
            ExcelBuilder builder = new ExcelBuilder();
            builder.AddSheet("sheet 1");

            var sheet = builder.GetSheet("sheet 1");

            Assert.IsNotNull(sheet);
        }

        [TestMethod]
        public void Verify_Adding_Header()
        {
            ExcelBuilder builder = new ExcelBuilder();
            builder.AddSheet("sheet 1");
            var sheet = builder.GetSheet("sheet 1");

            sheet.AddHeader(new System.Collections.Generic.List<string>() { "HEADER 1", "HEADER 2", "HEADER 3" });

            // get cells
            var cells = sheet.GetCells();
            
            Assert.AreEqual(cells[1, 1].Value, "HEADER 1");
            Assert.AreEqual(cells[1, 2].Value, "HEADER 2");
            Assert.AreEqual(cells[1, 3].Value, "HEADER 3");
        }

        [TestMethod]
        public void Verify_Adding_Row()
        {
            ExcelBuilder builder = new ExcelBuilder();
            builder.AddSheet("sheet 1");
            var sheet = builder.GetSheet("sheet 1");

            sheet.AddRow(new System.Collections.Generic.List<object>() { "A", "B", "C" });

            // get cells
            var cells = sheet.GetCells();

            Assert.AreEqual(cells[1, 1].Value, "A");
            Assert.AreEqual(cells[1, 2].Value, "B");
            Assert.AreEqual(cells[1, 3].Value, "C");
        }

        [TestMethod]
        public void Verify_Adding_Header_And_Several_Rows()
        {
            ExcelBuilder builder = new ExcelBuilder();
            builder.AddSheet("sheet 1");
            var sheet = builder.GetSheet("sheet 1");

            sheet.AddHeader(new System.Collections.Generic.List<string>() { "HEADER 1", "HEADER 2", "HEADER 3" });

            sheet.AddRow(new System.Collections.Generic.List<object>() { "A", "B", "C" });
            sheet.AddRow(new System.Collections.Generic.List<object>() { "D", "E", "F" });
            sheet.AddRow(new System.Collections.Generic.List<object>() { "G", "H", "I" });

            // get cells
            var cells = sheet.GetCells();
            
            Assert.AreEqual(cells[1, 1].Value, "HEADER 1");
            Assert.AreEqual(cells[1, 2].Value, "HEADER 2");
            Assert.AreEqual(cells[1, 3].Value, "HEADER 3");

            Assert.AreEqual(cells[2, 1].Value, "A");
            Assert.AreEqual(cells[2, 2].Value, "B");
            Assert.AreEqual(cells[2, 3].Value, "C");

            Assert.AreEqual(cells[3, 1].Value, "D");
            Assert.AreEqual(cells[3, 2].Value, "E");
            Assert.AreEqual(cells[3, 3].Value, "F");

            Assert.AreEqual(cells[4, 1].Value, "G");
            Assert.AreEqual(cells[4, 2].Value, "H");
            Assert.AreEqual(cells[4, 3].Value, "I");
        }

        /// <summary>
        /// Use a dummy json file to create a report
        /// </summary>
        [TestMethod]
        [DeploymentItem(@"input1.json")]
        public void Verify_Report_Output_One_Hotel_With_One_Rating()
        {
            IReporter reporter = new ExcelReporter();

            var input = JsonConvert.DeserializeObject<HotelCollection>(File.ReadAllText("input1.json", Encoding.Default));

            // get dummy input 1
            byte[] output = reporter.CreateReport(input.HotelList);

            // verify package
            using (ExcelPackage p = new ExcelPackage(new MemoryStream(output)))
            {
                var sheet = p.Workbook.Worksheets.First();

                Assert.AreEqual(sheet.Name, "Hotel 1");
                Assert.AreEqual(DateTime.FromOADate((double) sheet.Cells[2, 1].Value).ToString(), @"14/03/2016 23:00:00");
                Assert.AreEqual(DateTime.FromOADate((double) sheet.Cells[2, 2].Value).ToString(), @"15/03/2016 23:00:00");
                Assert.AreEqual(sheet.Cells[2, 3].Value, 116.1);
                Assert.AreEqual(sheet.Cells[2, 4].Value, "EUR");
                Assert.AreEqual(@"Name 1", (String)sheet.Cells[2, 5].Value);
                Assert.AreEqual(2.0, sheet.Cells[2, 6].Value);
                Assert.AreEqual(sheet.Cells[2, 7].Value, false);
            }
        }
        
        [TestMethod]
        [DeploymentItem(@"input2.json")]
        public void Verify_Report_Output_Two_Hotels_With_Multiple_Ratings()
        {
            IReporter reporter = new ExcelReporter();

            var input = JsonConvert.DeserializeObject<HotelCollection>(File.ReadAllText("input2.json", Encoding.Default));

            // get dummy input 2
            byte[] output = reporter.CreateReport(input.HotelList);

            // verify package
            using (ExcelPackage p = new ExcelPackage(new MemoryStream(output)))
            {
                var sheet = p.Workbook.Worksheets.FirstOrDefault(s => s.Name.Equals("Hotel 1"));

                Assert.IsNotNull(sheet);

                Assert.AreEqual(sheet.Name, "Hotel 1");
                Assert.AreEqual(DateTime.FromOADate((double)sheet.Cells[2, 1].Value).ToString(), @"14/03/2016 23:00:00");
                Assert.AreEqual(DateTime.FromOADate((double)sheet.Cells[2, 2].Value).ToString(), @"15/03/2016 23:00:00");
                Assert.AreEqual(sheet.Cells[2, 3].Value, 116.1);
                Assert.AreEqual(sheet.Cells[2, 4].Value, "EUR");
                Assert.AreEqual(@"Name 1", (String)sheet.Cells[2, 5].Value);
                Assert.AreEqual(2.0, sheet.Cells[2, 6].Value);
                Assert.AreEqual(sheet.Cells[2, 7].Value, false);

                sheet = p.Workbook.Worksheets.FirstOrDefault(s => s.Name.Equals("Hotel 2"));

                Assert.IsNotNull(sheet);

                Assert.AreEqual(sheet.Name, "Hotel 2");

                Assert.AreEqual(DateTime.FromOADate((double)sheet.Cells[2, 1].Value).ToString(), @"14/03/2016 23:00:00");
                Assert.AreEqual(DateTime.FromOADate((double)sheet.Cells[2, 2].Value).ToString(), @"15/03/2016 23:00:00");
                Assert.AreEqual(sheet.Cells[2, 3].Value, 78.5);
                Assert.AreEqual(sheet.Cells[2, 4].Value, "EUR");
                Assert.AreEqual(@"Name 2", (String)sheet.Cells[2, 5].Value);
                Assert.AreEqual(2.0, sheet.Cells[2, 6].Value);
                Assert.AreEqual(sheet.Cells[2, 7].Value, true);

                Assert.AreEqual(DateTime.FromOADate((double)sheet.Cells[3, 1].Value).ToString(), @"14/03/2016 23:00:00");
                Assert.AreEqual(DateTime.FromOADate((double)sheet.Cells[3, 2].Value).ToString(), @"15/03/2016 23:00:00");
                Assert.AreEqual(sheet.Cells[3, 3].Value, 116.1);
                Assert.AreEqual(sheet.Cells[3, 4].Value, "EUR");
                Assert.AreEqual(@"Name 3", (String)sheet.Cells[3, 5].Value);
                Assert.AreEqual(1.0, sheet.Cells[3, 6].Value);
                Assert.AreEqual(sheet.Cells[3, 7].Value, false);
            }
        }
    }
}
