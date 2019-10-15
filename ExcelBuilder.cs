using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PrestaCapExercise
{    
    public class ExcelBuilder : IDisposable
    {
        public class Sheet
        {
            /// <summary>
            /// reference to worksheet
            /// </summary>
            private ExcelWorksheet Worksheet { get; set; }

            /// <summary>
            /// used to keep track of current row
            /// </summary>
            private int Row { get; set; }

            public Sheet(ExcelWorksheet worksheet)
            {
                this.Worksheet = worksheet;
                this.Row = 1;
            }

            /// <summary>
            /// Add header for work sheet
            /// </summary>
            /// <param name="ws"></param>
            public void AddHeader(List<string> headers)
            {
                for (int i = 1; i <= headers.Count; ++i)
                {
                    Worksheet.Cells[this.Row, i].Value = headers[i - 1];
                    Worksheet.Cells[this.Row, i].Style.Font.Bold = true;
                    // make columns sortable
                    Worksheet.Cells[this.Row, i].AutoFilter = true;
                }
                
                ++this.Row;
            }

            /// <summary>
            /// add a row to the spread sheet
            /// </summary>
            /// <param name="range"></param>
            /// <param name="row"></param>
            /// <param name="rating"></param>
            public void AddRow(List<object> values)
            {
                for(int i = 1; i <= values.Count; ++i)
                {
                    object value = values[i - 1];
                    var cell = Worksheet.Cells[this.Row, i];

                    // set the value of the excel cell
                    cell.Value = value;

                    if(value is DateTime)
                    {
                        cell.Style.Numberformat.Format = "yyyy/mm/dd hh-mm";
                    }
                }

                if (this.Row % 2 == 0)
                {
                    // color entire row
                    var range = Worksheet.Cells[this.Row, 1, this.Row, values.Count];
                    range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightBlue);
                }

                ++this.Row;
            }

            /// <summary>
            /// Used for testing
            /// </summary>
            /// <returns></returns>
            public ExcelRange GetCells()
            {
                return Worksheet.Cells;
            }
        }
        
        private ExcelPackage Package { get; set; }

        private Dictionary<string, Sheet> Sheets { get; set; }

        public ExcelBuilder()
        {
            this.Package = new ExcelPackage();
            this.Sheets = new Dictionary<string, Sheet>();
        }
        
        /// <summary>
        /// Add a new work sheet
        /// </summary>
        /// <param name="p"></param>
        /// <param name="hotel"></param>
        public void AddSheet(string name)
        {
            this.Sheets.Add(name, new Sheet(Package.Workbook.Worksheets.Add(name)));
        }

        /// <summary>
        /// get an existing sheet
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        public Sheet GetSheet(string name)
        {
            return this.Sheets[name];
        }

        /// <summary>
        /// Returns all sheets in excel package
        /// </summary>
        /// <returns></returns>
        public List<Sheet> GetSheets()
        {
            return this.Sheets.Values.ToList();
        }

        /// <summary>
        /// Gets the byte array representing the excel file
        /// </summary>
        /// <param name="path"></param>
        public byte[] GetByteArray()
        {
            return this.Package.GetAsByteArray();
        }

        /// <summary>
        /// close package if still open
        /// </summary>
        public void Dispose()
        {
            this.Package.Dispose();
        }
    }
}
