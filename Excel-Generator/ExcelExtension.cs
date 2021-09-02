using ClosedXML.Excel;
using Excel_Generator;
using System.IO;
using System.Linq;

namespace Pecege.MoveHumaniza.Domain.Extensions
{
    public static class ExcelExtension
    {
        public static MemoryStream GenerateExcelFile(this Excel excelData)
        {
            var headerColumns = excelData.Rows.FirstOrDefault();
            char[] alphabet = "abcdefghijklmnopqrstuvwxyz".ToCharArray();
            var firstColumn = alphabet[0];
            var lastColumn = alphabet[headerColumns.Length - 1];
            var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add(excelData.WorkSheetName);
            var cellRow = 1;
            var headerData = worksheet.Range($"{firstColumn}{cellRow}", $"{lastColumn}{cellRow}");

            SetGeralStyles(worksheet);
            SetHeader(headerData, $"{excelData.HeaderName}");
            cellRow++;

            var header = worksheet.Range($"{firstColumn}{cellRow}", $"{lastColumn}{cellRow}");
            header.Style = headerData.Style;

            for (var i = 0; i < headerColumns.Length; i++)
            {
                SetColumn(worksheet, alphabet[i].ToString(), cellRow, headerColumns[i]);
            }

            for (var i = 1; i < excelData.Rows.Count; i++)
            {
                cellRow++;

                for (var j = 0; j < headerColumns.Length; j++)
                {
                    SetColumn(worksheet, alphabet[j].ToString(), cellRow, excelData.Rows[i][j]);
                }
            }

            worksheet.Columns().AdjustToContents();
            worksheet.Rows().AdjustToContents();

            return SaveWorkbookToMemoryStream(workbook);
        }

        private static MemoryStream SaveWorkbookToMemoryStream(XLWorkbook workbook)
        {
            MemoryStream stream = new MemoryStream();
            workbook.SaveAs(stream, new SaveOptions { EvaluateFormulasBeforeSaving = false, GenerateCalculationChain = false, ValidatePackage = false });

            return stream;
        }

        private static void SetHeader(IXLRange xLRange, string rowValue)
        {
            xLRange.Value = rowValue;
            xLRange.Merge();
            xLRange.Style.Font.SetBold();
            xLRange.Style.Font.FontColor = XLColor.White;
            xLRange.Style.Fill.BackgroundColor = XLColor.FromArgb(155, 103, 229);
            xLRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
            xLRange.Style.Border.TopBorder = XLBorderStyleValues.Thin;
            xLRange.Style.Border.LeftBorder = XLBorderStyleValues.Thin;
            xLRange.Style.Border.RightBorder = XLBorderStyleValues.Thin;
            xLRange.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
        }

        private static void SetColumn(IXLWorksheet xLWorksheet, string letter, int cellRow, string rowValue, bool setBold = false, XLColor xLColor = null)
        {
            var cell = xLWorksheet.Cell($"{letter}{cellRow}");
            cell.Value = rowValue;
            cell.Style.Border.TopBorder = XLBorderStyleValues.Thin;
            cell.Style.Border.LeftBorder = XLBorderStyleValues.Thin;
            cell.Style.Border.RightBorder = XLBorderStyleValues.Thin;
            cell.Style.Border.BottomBorder = XLBorderStyleValues.Thin;

            if (xLColor != null)
                cell.Style.Font.FontColor = xLColor;

            if (setBold)
                cell.Style.Font.SetBold();
        }

        private static void SetGeralStyles(IXLWorksheet xLWorksheet)
        {
            xLWorksheet.ShowGridLines = true;
            xLWorksheet.Style.Font.FontSize = 14;
            xLWorksheet.Style.Font.SetFontName("Arial, Helvetica, sans-serif");
        }
    }
}
