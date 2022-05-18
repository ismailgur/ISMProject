using ISMExtensions.Extensions;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Reflection;

namespace ISMExcelHelper
{
    public static class Excel_Export_Generic_List<T>
    {
        public static Stream Export(List<T> model, string workSheetName)
        {
            MemoryStream stream;
            using (ExcelPackage package = new ExcelPackage())
            {
                package.Workbook.Worksheets.Add(workSheetName);

                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                try
                {
                    int clmnWidth = 25;
                    int rowIndex = 2;

                    int i = 0;

                    foreach (T t in model)
                    {
                        if (i == 0)
                        {
                            int clmnCnt = 1;
                            foreach (PropertyInfo info in typeof(T).GetProperties())
                            {
                                var clmnDisplayName = info.GetCustomAttributes(typeof(DisplayNameAttribute), true);
                                worksheet.Cells[1, clmnCnt].Value = clmnDisplayName.Length == 0 ? info.Name : (clmnDisplayName[0] as DisplayNameAttribute).DisplayName;
                                worksheet.Column(clmnCnt).Width = clmnWidth;
                                //worksheet.Column(clmnIndex).Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                                //worksheet.Column(clmnIndex).Style.Font.Color.SetColor(Color.Black);
                                clmnCnt++;
                            }

                            //worksheet.Column(1).Width = 0; // id alanını gizlemek için
                        }

                        int k = 1;
                        foreach (PropertyInfo info in typeof(T).GetProperties())
                        {
                            var value = info.GetValue(t, null);

                            if (value != null)
                            {
                                if (info.PropertyType.IsEnum || info.PropertyType.IsNullableEnum())
                                {
                                    value = info.PropertyType.GetDisplayForEnum(value);
                                    worksheet.Cells[rowIndex, k].Value = value;
                                }
                                else
                                {
                                    if (value is string)
                                        worksheet.Cells[rowIndex, k].Value = value.ToString().GetPlainTextFromHtml();
                                    else
                                    {
                                        worksheet.Cells[rowIndex, k].Value = value;

                                        if (value is DateTime)
                                            worksheet.Cells[rowIndex, k].Style.Numberformat.Format = "mm-dd-yy HH:MM";
                                    }
                                }
                            }


                            k++;
                        }
                        rowIndex++;

                        i++;
                    }

                    stream = new MemoryStream(package.GetAsByteArray());
                }
                catch (Exception)
                {
                    throw;
                }
            }

            return stream;
        }
    }
}