// See https://aka.ms/new-console-template for more information

using ClosedXML.Excel;

namespace _ExportCataloguePhotosFromExcel
{
    internal class Program
    {
        private static string catalogueDirectory = "";
        private static string outPutPath = "";
        private static int complications = 0;
        private static string WorkSheetPath = "";
        private static int WorkSheetNumber = 1;
        private static string selectedColumn = "";
        private static XLWorkbook workbook;
        private static bool removeUnderscore = false;

        static void Main(string[] args)
        {
            PrintIntro();
            Setup();

            Console.Out.WriteLine("Scanning all files now!");
            var worksheet = workbook.Worksheets.Worksheet(WorkSheetNumber);
            foreach (var row in worksheet.Rows())
            {
                var cell = row.Cell("" + selectedColumn);
                try
                {
                    if (cell.Value.ToString().StartsWith("G"))
                    {
                        var picture = FindFileInCatalog(cell.Value);
                        if (picture.Equals("")) continue;

                        string copyPath = "";
                        if (removeUnderscore)
                        {
                            copyPath = (outPutPath + "\\" + picture + ".JPG").Replace("_", "");
                        }
                        else
                        {
                            copyPath = (outPutPath + "\\" + picture + ".JPG");
                        }

                        File.Copy(
                            catalogueDirectory + "\\" + cell.Value.ToString().Substring(0, 4) + "\\" + picture + ".JPG", copyPath
                            );
                    }
                    else
                    {
                        Console.WriteLine("File is invalid");
                        continue;
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.ToString());
                }
                
            }
            
        }

        private static string FindFileInCatalog(XLCellValue cellValue)
        {
            string dir = catalogueDirectory + "\\" + cellValue.ToString().Substring(0, 4);
            if (Directory.Exists(dir))
            {
                string[] pictures = Directory.GetFiles(dir);
                foreach (var picture in pictures)
                {
                    FileInfo fileInfo = new FileInfo(picture);
                    string pictureName = fileInfo.Name.Replace(fileInfo.Extension, "");
                    if (pictureName.Equals(cellValue.ToString()))
                    {
                        return pictureName;
                    }
                }
            }

            Console.WriteLine("Could not find file! " + cellValue);
            return "";
        }


        private static void PrintIntro()
        {
            Console.WriteLine("EXPORT CATALOGUE PHOTOS FROM EXCEL APPLICATION");

        }

        private static void Setup()
        {
            while (outPutPath == "")
            {
                Console.Write("\nChoose your output path: ");
                string? s = Console.ReadLine();
                if (s != null)
                {
                    if (Directory.Exists(s))
                    {
                        string[] files = Directory.GetFiles(s);
                        if (files.Length > 0)
                        {
                            Console.WriteLine("Directory is not empty!");
                            continue;
                        }
                        else
                        {
                            outPutPath = s;
                        }
                    }
                    else
                    {
                        Directory.CreateDirectory(s);
                        outPutPath = s;
                    }
                }
                else
                {
                    Console.WriteLine("Invalid input!");
                }
                
            }
            
            
            while (catalogueDirectory == "")
            {
                Console.Write("\nChoose your catalog path: ");
                string? s = Console.ReadLine();
                if (s != null)
                {
                    if (Directory.Exists(s))
                    {
                        catalogueDirectory = s;
                    }
                    else
                    {
                        Console.WriteLine("Directory doesn't exist!");
                    }
                }
                else
                {
                    Console.WriteLine("Invalid input!");
                }
                
            }

            while (WorkSheetPath == "")
            {
                Console.Write("\nChoose your excel file path: ");
                string? s = Console.ReadLine();
                if (s != null)
                {
                    if (File.Exists(s))
                    {
                        FileInfo fileInfo = new FileInfo(s);
                        if (fileInfo.Extension.Equals(".xlsx"))
                        {
                            WorkSheetPath = s;
                            workbook = new XLWorkbook(s);
                        }
                        else
                        {
                            Console.WriteLine("File is not an .xlsx file");
                        }
                    }
                    else
                    {
                        Console.WriteLine("File doesn't exist");
                    }
                }
                else
                {
                    Console.Out.WriteLine("Invalid Input!");
                }
            }

            if (workbook.Worksheets.Count > 1)
            {
                Console.WriteLine("The xlsx file has more than one work sheet!");
                Console.Write("\nWhich one do you choose? ");
                Console.Write("(");
                for(int i = 1; i < workbook.Worksheets.Count+1; i++)
                {
                    Console.Write(i + ",");
                }

                Console.Write("): ");
                WorkSheetNumber = Int32.Parse(Console.ReadLine());
            }

            while (selectedColumn == "")
            {
                Console.Write("\nChoose column to scan(Remember capital letters): ");
                string? s = Console.ReadLine();
                if (s != null)
                {
                    selectedColumn = s;
                }
                else
                {
                    Console.Out.WriteLine("Invalid Input!");
                }
                
            }
            
            Console.Write("\nRemove underscored from output files? ");
            string? b = Console.ReadLine();
            if (b.Equals("y"))
            {
                removeUnderscore = true;
            }

        }
    }
}