using ExcelDataReader; // package: ExcelDataReader
using ISMExtensions.Extensions;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.IO;
using System.Linq;
using System.Reflection;

namespace ISMExcelHelper
{
    public static class Excel_Import_Generic_List<T> where T : new()
    {
        public static List<T> GetData(string fileUrl, string workSheetName = null, bool matchPropertyDisplayNameAttribute = false, bool matchEnumDisplayAttribute = false)
        {
            var stream = System.IO.File.Open(fileUrl, FileMode.Open, FileAccess.Read);
            return GetData(stream, workSheetName, matchPropertyDisplayNameAttribute, matchEnumDisplayAttribute);
        }


        public static List<T> GetData(Stream stream, string workSheetName = null, bool matchPropertyDisplayNameAttribute = false, bool matchEnumDisplayAttribute = false)
        {
            List<T> data = new List<T>();
            dynamic m = new System.Dynamic.ExpandoObject();

            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            using (stream)
            {
                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var i = 0;
                    Dictionary<string, int> columns = new Dictionary<string, int>();
                    do
                    {
                        while (reader.Read())
                        {
                            if (workSheetName != null && reader.Name != workSheetName)
                                break;

                            if (i >= 0)
                            {
                                try
                                {
                                    if (i == 0)
                                    {
                                        try
                                        {
                                            for (int k = 0; k < reader.FieldCount; k++)
                                            {
                                                columns.Add(reader[k].ToString(), k);
                                            }
                                        }
                                        catch (Exception)
                                        {
                                        }
                                    }
                                    else
                                    {
                                        var a = new T();

                                        foreach (PropertyInfo prop in a.GetType().GetProperties())
                                        {
                                            var q = columns.Where(x => x.Key == prop.Name);

                                            if (matchPropertyDisplayNameAttribute)
                                            {
                                                var attr = prop.CustomAttributes.Where(x => x.AttributeType.FullName == "System.ComponentModel.DisplayNameAttribute").SingleOrDefault();
                                                var attrValue = attr.ConstructorArguments[0].Value;

                                                q = columns.Where(x => x.Key == attrValue.ToString());
                                            }


                                            if (q.Count() == 1)
                                            {

                                                object value = null;
                                                try
                                                {
                                                    value = reader[q.Single().Value] is string ? reader[q.Single().Value].ToString().GetPlainTextFromHtml() : reader[q.Single().Value];

                                                    Type t = Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType;

                                                    var isNull = IsNull(value);
                                                    var type = prop.PropertyType;


                                                    if (isNull)
                                                    {
                                                        if (!type.IsNullAssignable())
                                                        {
                                                            throw new InvalidCastException($"Cannot cast null to {type}");
                                                        }
                                                        continue;
                                                    }

                                                    var nonNullableType = type.AsNonNullable();

                                                    if (nonNullableType.IsEnum)
                                                    {
                                                        if (matchEnumDisplayAttribute)
                                                        {
                                                            foreach (var field in nonNullableType.GetFields())
                                                            {
                                                                DisplayAttribute attribute = Attribute.GetCustomAttribute(field, typeof(DisplayAttribute)) as DisplayAttribute;
                                                                if (attribute == null)
                                                                    continue;

                                                                if (attribute.Name == value?.ToString())
                                                                {
                                                                    value = field.GetValue(null);
                                                                    break;
                                                                }
                                                            }
                                                        }
                                                        else
                                                        {
                                                            var enumUnderlyingType = Enum.GetUnderlyingType(nonNullableType);
                                                            var enumUnderlyingValue = Convert.ChangeType(value, enumUnderlyingType);
                                                            value = Enum.ToObject(nonNullableType, enumUnderlyingValue);
                                                        }
                                                    }
                                                    else
                                                        value = Convert.ChangeType(value, nonNullableType);

                                                    prop.SetValue(a, value, null);
                                                }
                                                catch (Exception err)
                                                {
                                                }
                                            }
                                        }

                                        data.Add(a);
                                    }
                                }
                                catch (Exception err)
                                {
                                    //return null;
                                }
                            }
                            i++;
                        }
                    } while (reader.NextResult());
                }
            }

            return data;
        }

        
        private static bool IsNull(object value)
        {
            // - value == null uses the type's equality operator (usefull for Nullable)
            // - ReferenceEquals checks for actual null references
            // - DBNull is a special null value

            return value == null
                || ReferenceEquals(null, value)
                || value is DBNull;
        }
    }
}
