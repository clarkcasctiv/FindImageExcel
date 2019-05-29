﻿using Npoi.Mapper;
using Npoi.Mapper.Attributes;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.IO;

namespace FindFiles
{
    public class Program
    {
        public static List<string> fileList = new List<string>();

        private static void Main(string[] args)
        {
            string excelFile = @"D:\ExcelFile\ExcelImage.xlsx";
            string directory = @"D:\SajiloImage";


            if(!File.Exists(excelFile))
            {
                Console.WriteLine("Excel File Not Found");
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
                    }
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

                if (fileEntries != null)
                {
                    string destination = @"D:\ExcelFile\Destination";

                    if (!Directory.Exists(destination))
                    {
                        Directory.CreateDirectory(destination);
                    }
                    string destinationFileName = Path.Combine(destination, imageName);

                    Console.WriteLine("Moving File '{0}' to '{1}'", fileName, destination);

                    if(!File.Exists(destinationFileName))
                    {
                        File.Copy(fileName, destinationFileName);
                        Console.WriteLine("Moved File '{0}' to '{1}", fileName, destination);


                    }
                    else
                    {
                        Console.WriteLine("Filename '{0}' Already Exists in destination",fileName);
                    }

                }

                //ProcessFile(fileName);

            }

            string[] subdirectoryEntries = Directory.GetDirectories(targetDirectory);
            foreach (string subdirectory in subdirectoryEntries)
            {
                ProcessDirectory(subdirectory, imageName);
            }
        }

        public static void ProcessFile(string path)
        {
            //string destination = @"D:\ExcelFile\Destination";

            ////Console.WriteLine("Processed file '{0}'.", path);

            //if(File.Exists(path))
            //{
            //    Console.WriteLine("Moved File '{0}' to '{1}", path, destination);

            //}

        }


        public static List<string> FindAllFiles(string targetDirectory)
        {

            string[] fileEntries = Directory.GetFiles(targetDirectory);
            foreach (string fileName in fileEntries)
            {
                if(fileEntries !=null)
                {
                    Console.WriteLine("Found file '{0}'.", fileName);

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

            //[Column(3)] By Column
        }
    }
}