using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace GailisRunTests
{
    static class TestManager
    {
        private static readonly string mainDirectoryPath = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
        private static string MainDirectoryPath => mainDirectoryPath;
        static void Main(string[] args)
        {
            /*Find and prepare master excel test case file*/
            string MainExcelPath = MainDirectoryPath + "\\gailisTest.xlsx";
            if (File.Exists(MainExcelPath))
            {
                DirectoryInfo di = new DirectoryInfo(MainDirectoryPath);
                Excel main_excel = new Excel(MainExcelPath, 1);
                //Clear master test case file result fields
                List<string> tests = new List<string>();
                Console.WriteLine("Generate test cases from master? y/n:");
                if (Console.ReadLine() == "y")
                {
                    foreach (Worksheet sheet in main_excel.Wb.Worksheets)
                    {
                        string testPath = MainDirectoryPath + "\\" + sheet.Name + ".txt";
                        main_excel.ChangeSheet(sheet.Index, false);
                        Cell input = new Cell(2, 0);
                        using (StreamWriter outputFile = new StreamWriter(testPath))
                        {
                            while (main_excel.ReadCell(input) != "")
                            {
                                outputFile.WriteLine(main_excel.ReadCell(input));
                                input.Down();
                            }
                        }
                        tests.Add(testPath);
                       
                    }
                }
                else
                {
                    foreach (Worksheet sheet in main_excel.Wb.Worksheets)
                    {
                        string testPath = MainDirectoryPath + "\\" + sheet.Name + ".txt";
                        if (File.Exists(testPath))
                        {
                            tests.Add(testPath);
                        }
                        else
                        {
                            Console.WriteLine("Couldn't find test case for {0}. Generating...", sheet.Name);
                            Cell input = new Cell(2, 0);
                            using (StreamWriter outputFile = new StreamWriter(testPath))
                            {
                                while (main_excel.ReadCell(input) != "")
                                {
                                    outputFile.WriteLine(main_excel.ReadCell(input));
                                    input.Down();
                                }
                            }
                            tests.Add(testPath);
                        }
                    }
                }
                main_excel.Save();
                /*Paste master test case file into test directories*/
                List<(string, string)> exe_excel_map = new List<(string, string)>();
                //Test files are marked by a timestamp
                DateTime local = DateTime.Now;
                string time_signature = local.ToString("ddMMyyyy") + "_" + local.ToString("HHmmss");
                //Try pasting master test case file and add exe - excel pairs to mapping if successful
                foreach (DirectoryInfo dir in di.GetDirectories())
                {
                    //Name new files gailisTest_DIR_TIMESTAMP
                    string fullPath = Path.GetFullPath(dir.FullName).TrimEnd(Path.DirectorySeparatorChar);
                    string dirName = fullPath.Split(Path.DirectorySeparatorChar).Last();
                    string testFilePath = fullPath + "\\gailisTest_" + dirName + "_" + time_signature + ".xlsx";
                    string exeFilePath = fullPath + "\\gailis_" + dirName + ".exe";
                    if (File.Exists(exeFilePath))
                    {
                        try
                        {
                            if (main_excel.SaveAs(testFilePath))
                            {
                                Console.WriteLine("Saved to {0} successfully", testFilePath);
                                exe_excel_map.Add((fullPath, dirName));
                            }
                            else
                            {
                                Console.WriteLine("Failed to save to {0}", testFilePath);
                            }
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine("Exception thrown while saving to {0}", testFilePath);
                            main_excel.Close();
                            throw;
                        }
                    }
                    else
                    {
                        Console.WriteLine("Failed to find exe at {0} while saving test xlsx copies", exeFilePath);
                    }
                }
                main_excel.Close();
                /*Run test cases for each successful exe - excel mapping*/
                foreach ((string, string) mapping in exe_excel_map)
                {
                    string inputFilePath = mapping.Item1 + "\\gailis.in";
                    string outputFilePath = mapping.Item1 + "\\gailis.out";

                    string testFilePath = mapping.Item1 + "\\gailisTest_" + mapping.Item2 + "_" + time_signature + ".xlsx";
                    string exeFilePath = mapping.Item1 + "\\gailis_" + mapping.Item2 + ".exe";

                    if (File.Exists(exeFilePath) && File.Exists(testFilePath))
                    {
                        //Define test exe launch options
                        var startInfo = new ProcessStartInfo
                        {
                            WorkingDirectory = mapping.Item1,
                            FileName = exeFilePath,
                            UseShellExecute = false,
                            RedirectStandardError = true
                        };
                        //Test case excel workbook
                        Excel ex = new Excel(testFilePath, 1);
                        //Prepare input/output files
                        if (!File.Exists(inputFilePath))
                        {
                            FileStream f = File.Create(inputFilePath);
                            f.Dispose();
                        }

                        if (!File.Exists(outputFilePath))
                        {
                            FileStream f = File.Create(outputFilePath);
                            f.Dispose();
                        }
                        /*Execute test cases*/


                        //Each worksheet symbolizes a test case
                        foreach (Worksheet sheet in ex.Wb.Worksheets)
                        {
                            string test = tests.Find(x => Path.GetFileNameWithoutExtension(x) == sheet.Name);
                            if (test != null)
                            {

                                //Select current sheet
                                ex.ChangeSheet(sheet.Index, false);
                                File.Copy(test, inputFilePath, true);
                                //Clear output file
                                using (StreamWriter writer = new StreamWriter(outputFilePath, false))
                                {
                                    writer.Close();
                                }

                                //Initialize cell trackers
                                Cell expectedOut = new Cell(2, 1);
                                Cell actualOut = new Cell(2, 2);
                                Cell result = new Cell(2, 3);
                                Cell exception = new Cell(2, 4);
                                //Write test case input field to input file

                                /*Run process with test case input*/
                                try
                                {
                                    using (Process p = new Process())
                                    {
                                        p.StartInfo = startInfo;
                                        p.Start();
                                        //p.BeginOutputReadLine();
                                        StreamReader sr = p.StandardError;
                                        string errors = sr.ReadLine();
                                        if (!p.WaitForExit(30000))//30 second timeout
                                        {
                                            ex.AlterCell(exception, "Timed Out");
                                            exception.Down();
                                            p.Kill();
                                        }
                                        ex.AlterCell(exception, errors);

                                    }
                                }
                                catch (Exception e)
                                {
                                    Console.WriteLine(mapping.Item1 + " failed test case: " + sheet.Name + " with error: " + e.Message);
                                }
                                /*Write output and test result to file*/
                                if (File.Exists(outputFilePath))
                                {
                                    using (StreamReader file = new StreamReader(outputFilePath))
                                    {
                                        string outputLine;
                                        outputLine = file.ReadLine();
                                        string expectedLine;
                                        expectedLine = ex.ReadCell(expectedOut);
                                        while (!string.IsNullOrEmpty(outputLine) || !string.IsNullOrEmpty(expectedLine))
                                        {
                                            ex.AlterCell(actualOut, outputLine);
                                            if (expectedLine == outputLine)
                                            {
                                                ex.AlterCell(result, "PASS");
                                            }
                                            else
                                            {
                                                ex.AlterCell(result, "FAIL");
                                            }
                                            expectedOut.Down();
                                            actualOut.Down();
                                            result.Down();
                                            outputLine = file.ReadLine();
                                            expectedLine = ex.ReadCell(expectedOut);
                                        }
                                    }
                                }
                                else
                                {
                                    ex.AlterCell(exception, "Output not found");
                                }
                                ex.Save();
                            }
                            else
                            {
                                Console.WriteLine("Couldn't find file for test case: {0}", sheet.Name);
                            }
                        }
                        ex.Close();
                    }
                    else
                    {
                        Console.WriteLine("exe not found at: " + mapping.Item1);
                    }
                }
            }
            else
            {
                Console.WriteLine("Main excel file not found");
            }
        }
    }
}
