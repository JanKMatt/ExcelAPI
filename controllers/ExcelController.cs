using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using System.IO;
using OfficeOpenXml.Style;
using System.Drawing;

namespace ExcelApi.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ExcelController : ControllerBase
    {
        [HttpGet("generate")]
        public IActionResult GenerateExcel()
        {
            // Sample 2D array of data (strings)
            string[,] data = new string[,]
            {
                { "Header1", "Header2", "Header3" },
                { "Row1Col1", "Row1Col2", "123.456" },
                { "Row2Col1", "Row2Col2", "789.012" },
                { "Row3Col1", "Row3Col2", "345.678" }
            };

            // Create an Excel package in memory
            using (ExcelPackage package = new ExcelPackage())
            {
                // Add a worksheet to the workbook
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Sheet1");

                // Populate the worksheet with the array data
                for (int row = 0; row < data.GetLength(0); row++)
                {
                    for (int col = 0; col < data.GetLength(1); col++)
                    {
                        if (col == data.GetLength(1) - 1 && row > 0) // Last column and not the header row
                        {
                            // Parse the string to a number and set the value
                            if (decimal.TryParse(data[row, col], out decimal numericValue))
                            {
                                worksheet.Cells[row + 1, col + 1].Value = numericValue; // Set as numeric
                            }
                            else
                            {
                                worksheet.Cells[row + 1, col + 1].Value = data[row, col]; // Fallback to string if parsing fails
                            }
                        }
                        else
                        {
                            worksheet.Cells[row + 1, col + 1].Value = data[row, col]; // Set as string for other cells
                        }
                    }
                }

                // Format the first row (header row) to be bold and red
                using (var range = worksheet.Cells[1, 1, 1, data.GetLength(1)])  // First row range
                {
                    range.Style.Font.Bold = true;
                    range.Style.Font.Color.SetColor(Color.Red);
                }

                // Format the first column (leftmost column) to be italic and blue
                using (var range = worksheet.Cells[1, 1, data.GetLength(0), 1])  // First column range
                {
                    range.Style.Font.Italic = true;
                    range.Style.Font.Color.SetColor(Color.Blue);
                }

                // Format the last column as numbers with two decimal places
                int lastColumnIndex = data.GetLength(1);  // Get the last column index
                using (var range = worksheet.Cells[2, lastColumnIndex, data.GetLength(0), lastColumnIndex])
                {
                    range.Style.Numberformat.Format = "#,##0.00";  // Format as number with two decimals
                }

                // Convert the Excel package to a memory stream
                var stream = new MemoryStream();
                package.SaveAs(stream);
                stream.Position = 0; // Reset stream position to the beginning

                // Return the Excel file as a downloadable file
                string excelFileName = "generated_excel.xlsx";
                string contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

                return File(stream, contentType, excelFileName);
            }
        }
    }
}
