using OfficeOpenXml; // package: OfficeOpenXml.Core.ExcelPackage
using OfficeOpenXml.Style; // package: EEPlus
using System.Drawing;

namespace ISMExcelHelper
{
    public static class ExcelHelper
    {
        public static ExcelRange GetRange(ExcelWorksheet worksheet, string text, int rowStart, int columnStart, int rowEnd, int columnEnd
      , float fontSize = 16
      , bool style = true
      , bool wrapText = false)
        {
            using (ExcelRange Rng = worksheet.Cells[rowStart, columnStart, rowEnd, columnEnd]) // satır başlangıç,sütun başlangıç,satır bitiş,sütun bitiş
            {
                Rng.Value = text;
                Rng.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                Rng.Merge = true;
                Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                Rng.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                Rng.Style.Border.Top.Color.SetColor(Color.Black);
                Rng.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                Rng.Style.Border.Left.Color.SetColor(Color.Black);
                Rng.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                Rng.Style.Border.Right.Color.SetColor(Color.Black);
                Rng.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                Rng.Style.Border.Bottom.Color.SetColor(Color.Black);
                Rng.Style.WrapText = wrapText;

                if (style)
                {

                    Rng.Style.Font.Color.SetColor(Color.Black);
                    Rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    Rng.Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                    Rng.Style.Font.Size = fontSize;
                    Rng.Style.Font.Bold = true;
                    Rng.Style.Font.Italic = true;
                }

                return Rng;
            }
        }


        public static ExcelRange GetRange1(ExcelWorksheet worksheet, string text, int rowStart, int columnStart, int rowEnd, int columnEnd
        , Color? fillColor = null
        , int fontSize = 16)
        {

            using (ExcelRange Rng = worksheet.Cells[rowStart, columnStart, rowEnd, columnEnd]) // satır başlangıç,sütun başlangıç,satır bitiş,sütun bitiş
            {
                Rng.Value = text;
                Rng.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                Rng.Style.Font.Size = fontSize;
                Rng.Style.Font.Bold = true;
                Rng.Style.Font.Italic = true;
                Rng.Merge = true;
                Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                Rng.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                Rng.Style.Border.Top.Color.SetColor(Color.Black);
                Rng.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                Rng.Style.Border.Left.Color.SetColor(Color.Black);
                Rng.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                Rng.Style.Border.Right.Color.SetColor(Color.Black);
                Rng.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                Rng.Style.Border.Bottom.Color.SetColor(Color.Black);
                Rng.Style.Font.Color.SetColor(Color.Black);
                Rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                Rng.Style.Fill.BackgroundColor.SetColor(Color.LightGray);

                if (fillColor.HasValue)
                    Rng.Style.Fill.BackgroundColor.SetColor(fillColor.Value);

                return Rng;
            }
        }


        public static void ExcelCell(ExcelWorksheet worksheet, string text, int row, int column, bool bold, ExcelHorizontalAlignment excelHorizontalAlignment,
     ExcelVerticalAlignment excelVerticalAlignment, ExcelFillStyle excelFillStyle, Color backgroundColor, int width)
        {
            worksheet.Cells[row, column].Value = text;
            worksheet.Cells[row, column].Style.Font.Bold = bold;
            worksheet.Cells[row, column].Style.HorizontalAlignment = excelHorizontalAlignment;
            worksheet.Cells[row, column].Style.VerticalAlignment = excelVerticalAlignment;
            worksheet.Cells[row, column].Style.Fill.PatternType = excelFillStyle;
            worksheet.Cells[row, column].Style.Fill.BackgroundColor.SetColor(backgroundColor);
            worksheet.Column(column).Width = width;
        }


        public static void ExcelCell(ExcelWorksheet worksheet, object value, int row, int column, Color? color = null)
        {
            worksheet.Cells[row, column].Value = value;
            worksheet.Cells[row, column].Style.Font.Bold = false;
            worksheet.Cells[row, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            worksheet.Cells[row, column].Style.VerticalAlignment = ExcelVerticalAlignment.Bottom;
            worksheet.Cells[row, column].Style.Fill.PatternType = ExcelFillStyle.None;

            if (color.HasValue)
                worksheet.Cells[row, column].Style.Font.Color.SetColor(color.Value);
        }
    }
}
