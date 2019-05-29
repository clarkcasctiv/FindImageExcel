using Npoi.Mapper;
using Npoi.Mapper.Attributes;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;

namespace FindFiles
{
    public class Program
    {
        public static List<string> fileList = new List<string>();
        public static string logFile = @"D:\ExcelFile\Log.txt";

        private static void Main(string[] args)
        {
            string excelFile = @"D:\ExcelFile\ExcelImage.xlsx";
            string directory = @"D:\SajiloImage";


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


            if(Directory.Exists(directory))
            {
                FindAllFiles(directory);
            }


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

                    if(!imageNames.Contains(row.Imagename))
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

            Console.ReadLine();
        }



        public static void ProcessDirectory(string targetDirectory, string imageName)

        {
            string[] fileEntries = Directory.GetFiles(targetDirectory, imageName);

            foreach (string fileName in fileEntries)
            {

                if (fileEntries != null)
                {
                    string destination = @"D:\ExcelFile\Destination";

                    if (!Directory.Exists(destination))
                    {
                        Directory.CreateDirectory(destination);
                    }
                    string destinationFileName = Path.Combine(destination, imageName);

                    Console.WriteLine("Moving File '{0}' to '{1}'", fileName, destination);
                    using (StreamWriter file = new StreamWriter(logFile, true))
                    {
                        file.WriteLine("Moving File '{0}' to '{1}'", fileName, destination);
                    }

                    if (!File.Exists(destinationFileName))
                    {
                        //Task.Run(() =>
                        //{
                        //    File.Copy(fileName, destinationFileName);

                        //});
                        File.Copy(fileName, destinationFileName);
                        Console.WriteLine("Moved File '{0}' to '{1}", fileName, destination);
                        using (StreamWriter file = new StreamWriter(logFile, true))
                        {
                            file.WriteLine("Moved File '{0}' to '{1}", fileName, destination);
                        }

                    }
                    else
                    {
                        Console.WriteLine("Filename '{0}' Already Exists in destination",fileName);
                        using (StreamWriter file = new StreamWriter(logFile, true))
                        {
                            file.WriteLine("Filename '{0}' Already Exists in destination",fileName);
                        }
                    }

                }

            }

            string[] subdirectoryEntries = Directory.GetDirectories(targetDirectory);
            foreach (string subdirectory in subdirectoryEntries)
            {
                ProcessDirectory(subdirectory, imageName);
            }
        }



        public static List<string> FindAllFiles(string targetDirectory)
        {
            //, "*eval*" SearchPattern

            string[] fileEntries = Directory.GetFiles(targetDirectory);
            foreach (string fileName in fileEntries)
            {
                if(fileEntries !=null)
                {
                    Console.WriteLine("Found file '{0}'.", fileName);
                    using( StreamWriter file = new StreamWriter(logFile, true))
                    {

                        file.WriteLine("Found file '{0}'.", fileName);
                    }

                    fileList.Add(fileName);
                }

            }

            string[] subdirectoryEntries = Directory.GetDirectories(targetDirectory);
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