using Npoi.Mapper;
using Npoi.Mapper.Attributes;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.IO;

namespace FindFiles
{
    public class Program
    {
        private static void Main(string[] args)
        {
            string excelFile = @"D:\ExcelFile\ExcelImage.xlsx";
            string directory = @"D:\SajiloImage";


            IWorkbook workbook;

            using (FileStream file = new FileStream(excelFile, FileMode.Open, FileAccess.Read))
            {
                workbook = WorkbookFactory.Create(file);
            }
            var importer = new Mapper(workbook);
            var items = importer.Take<ExcelFormat>(0, 0);


            foreach (var item in items)
            {
                var row = item.Value;

                if(Directory.Exists(directory))
                {
                    ProcessDirectory(directory, row.Imagename);

                }
                else
                {
                    Console.WriteLine("{0} is not a valid directory", directory);
                }

            }

            Console.ReadLine();
        }


        public static void ProcessDirectory(string targetDirectory, string imageName)

        {
            string[] fileEntries = Directory.GetFiles(targetDirectory, imageName);

            foreach (string fileName in fileEntries)
            {

                if(fileEntries != null)
                {
                    string destination = @"D:\ExcelFile\Destination";

                    if(!Directory.Exists(destination))
                    {
                        Directory.CreateDirectory(destination);
                    }
                    string destinationFileName = Path.Combine(destination, imageName);
                    File.Move(fileName, destinationFileName);

                }
                ProcessFile(fileName);
            }

            string[] subdirectoryEntries = Directory.GetDirectories(targetDirectory);
            foreach (string subdirectory in subdirectoryEntries)
            {
                ProcessDirectory(subdirectory, imageName);
            }
        }

        public static void  ProcessFile(string path)
        {

            Console.WriteLine("Processed file '{0}'.", path);

        }

        private class ExcelFormat
        {
            [Column("Imagename")]
            public string Imagename { get; set; }

            //[Column(3)] By Column
        }
    }
}