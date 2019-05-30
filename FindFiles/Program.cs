using Npoi.Mapper;
using Npoi.Mapper.Attributes;
using NPOI.SS.UserModel;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;

namespace FindFiles
{
    public class Program
    {
        public static List<string> fileList = new List<string>();
        public static string logFile = @"D:\SearchImage\Log.txt";

        private static void Main(string[] args)
        {
            Stopwatch watch = Stopwatch.StartNew();
            string directory = @"D:\SearchImage\Source";
            string excelFile = @"D:\SearchImage\ExcelImage.xlsx";

            using (StreamWriter file = new StreamWriter(logFile, true))
            {
                file.WriteLine("###################################");

                file.WriteLine("Start Of Program {0}", DateTime.Now);

            }

            if (!File.Exists(excelFile))
            {
                Console.WriteLine("Excel File Not Found");
                using (StreamWriter file = new StreamWriter(logFile, true))
                {
                    file.WriteLine("Excel File Not Found");
                }
                Console.ReadLine();
                return;
            }


            if (Directory.Exists(directory))
            {
                FindAllFiles(directory);
            }

            ListAllFiles();


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

                if (Directory.Exists(directory))
                {
                    List<string> imageNames = new List<string>();

                    foreach (var file in fileList)
                    {
                        if (Path.GetFileName(file) == row.Imagename)
                        {
                            ProcessDirectory(directory, row.Imagename);
                            imageNames.Add(Path.GetFileName(file));
                        }
                    }

                    if (!imageNames.Contains(row.Imagename))
                    {
                        Console.WriteLine("Image name '{0}' was not found", row.Imagename);
                        using (StreamWriter file = new StreamWriter(logFile, true))
                        {
                            file.WriteLine("Image name '{0}' was not found", row.Imagename);
                        }
                    }
                }
                else
                {
                    Console.WriteLine("{0} is not a valid directory", directory);
                    using (StreamWriter file = new StreamWriter(logFile, true))
                    {
                        file.WriteLine("{0} is not a valid directory", directory);
                    }
                }
            }
            Console.WriteLine(watch.Elapsed);
            watch.Stop();
            Console.ReadLine();
        }

        public static void ListAllFiles()
        {
            string listExcel = @"D:\SearchImage\ListExcel.xlsx";

            IWorkbook workbook;

            using (FileStream file = new FileStream(listExcel, FileMode.Open, FileAccess.Read))
            {
                workbook = WorkbookFactory.Create(file);
            }

            FileInfo file1 = new FileInfo(listExcel);

            using (ExcelPackage excelPackage = new ExcelPackage(file1))
            {
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets[1];


                for (int i = 1; i <= fileList.Count; i++)
                {
                    //worksheet.Cells[1, 1].Value = "Encoded String";

                    worksheet.Cells[i, 1].Value = Path.GetFileName(fileList[i - 1]);

                }

                //int i = 1;
                //foreach (var item in fileList)
                //{
                //     worksheet.Cells[i, 1].Value = Path.GetFileName(fileList[i]);
                //    i++;
                //}
                excelPackage.Save();
            }


        }
        public static void ProcessDirectory(string targetDirectory, string imageName)

        {
            var fileEntries = Directory.EnumerateFiles(targetDirectory, imageName);
            //string[] fileEntries = Directory.GetFiles(targetDirectory, imageName);

            foreach (string fileName in fileEntries)
            {
                if (fileEntries != null)
                {
                    string destination = @"D:\SearchImage\Destination\";

                    if (!Directory.Exists(destination))
                    {
                        Directory.CreateDirectory(destination);
                    }

                    string destinationFileName = Path.Combine(destination, imageName);

                    string name = imageName.Substring(imageName.LastIndexOf('_') + 1, 10);
                    string[] path = name.Split('-');
                    var directory = destination + path[0] + @"\" + path[1] + @"\" + path[2] + @"\";
                    var imgFullPath = destination + path[0] + @"\" + path[1] + @"\" + path[2] + @"\" + imageName;

                    if (!Directory.Exists(directory))
                    {
                        Directory.CreateDirectory(directory);
                    }

                    //Console.WriteLine("Moving File '{0}' to '{1}'", fileName, destination);
                    Console.WriteLine("Moving File '{0}' to '{1}'", fileName, directory);

                    using (StreamWriter file = new StreamWriter(logFile, true))
                    {
                        file.WriteLine("Moving File '{0}' to '{1}'", fileName, directory);
                    }

                    if (!File.Exists(imgFullPath))
                    {
                        //Task.Run(() =>
                        //{
                        //    File.Copy(fileName, destinationFileName);

                        //});
                        File.Copy(fileName, imgFullPath);
                        Console.WriteLine("Moved File '{0}' to '{1}", fileName, directory);
                        using (StreamWriter file = new StreamWriter(logFile, true))
                        {
                            file.WriteLine("Moved File '{0}' to '{1}", fileName, directory);
                        }
                    }
                    else
                    {
                        Console.WriteLine("Filename '{0}' Already Exists in destination", fileName);
                        using (StreamWriter file = new StreamWriter(logFile, true))
                        {
                            file.WriteLine("Filename '{0}' Already Exists in destination", fileName);
                        }
                    }
                }
            }

            var subdirectoryEntries = Directory.EnumerateDirectories(targetDirectory);
            //string[] subdirectoryEntries = Directory.GetDirectories(targetDirectory);

            foreach (string subdirectory in subdirectoryEntries)
            {
                ProcessDirectory(subdirectory, imageName);
            }
        }

        public static List<string> FindAllFiles(string targetDirectory)
        {
            //, "*eval*" SearchPattern
            var fileEntries = Directory.EnumerateFiles(targetDirectory, "*merch*");
            //string[] fileEntries = Directory.GetFiles(targetDirectory);

            foreach (string fileName in fileEntries)
            {
                if (fileEntries != null)
                {
                    Console.WriteLine("Found file '{0}'.", fileName);
                    using (StreamWriter file = new StreamWriter(logFile, true))
                    {
                        file.WriteLine("Found file '{0}'.", fileName);
                    }

                    fileList.Add(fileName);
                }
            }

            var subdirectoryEntries = Directory.EnumerateDirectories(targetDirectory);
            //string[] subdirectoryEntries = Directory.GetDirectories(targetDirectory);

            foreach (string subdirectory in subdirectoryEntries)
            {
                FindAllFiles(subdirectory);
            }

            return fileList;
        }

        private class ExcelFormat
        {

            [Column("Imagename")]
            public string Imagename { get; set; }


        }
    }
}