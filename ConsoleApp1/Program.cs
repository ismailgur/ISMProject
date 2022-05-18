using System;
using System.Collections.Generic;
using System.IO;

namespace ConsoleApp1
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");

            var list = new List<testmodel>();
            list.Add(new testmodel
            {
                Id = 1,
                FullName = "İsmail GÜr"
            });

            var excelStream = ISMExcelHelper.Excel_Export_Generic_List<testmodel>.Export(list, "test");
            string path = @"c:\users\ismail\desktop\test.xlsx";

            using (var fileStream = File.Create(path))
            {
                excelStream.Seek(0, SeekOrigin.Begin);
                excelStream.CopyTo(fileStream);
            }
        }
    }

    class testmodel
    {
        public long Id { get; set; }

        public string FullName { get; set; }
    }
}
