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
        public static string MainDirectoryPath => mainDirectoryPath;

        static void Main(string[] args)
        {

            //Regex ext = new Regex(@"^i\d\d*$");
            //get containing directory of exe
            DirectoryInfo di = new DirectoryInfo(MainDirectoryPath);
            //get subdirs of containing directory
            var directories = di.GetDirectories();
            if (directories.Count() > 0)
            {
                foreach (var directory in directories)
                {
                    //fullpath - path to subdirectory
                    string fullPath = Path.GetFullPath(directory.FullName).TrimEnd(Path.DirectorySeparatorChar);

                    //projectName - name of subdirectory
                    string projectName = fullPath.Split(Path.DirectorySeparatorChar).Last();

                    //exeName - gailis_[projectName]
                    string exeFilePath = fullPath + "\\gailis_" + projectName + ".exe";

                    //testBookName - gailisTest_[projectName]
                    string testFilePath = fullPath + "\\gailisTest_" + projectName + ".xlsx";

                    //If subdir name compliant exe and testcase xlsx exist


                    if (File.Exists(exeFilePath) && File.Exists(testFilePath))
                    {
                        var startInfo = new ProcessStartInfo
                        {
                            WorkingDirectory = fullPath,
                            FileName = exeFilePath
                        };

                        //test case excel workbook
                        Excel ex = new Excel(testFilePath, 1);
                        string inputFilePath = fullPath + "\\gailis.in";

                        //create input file if it doesn't exist; this isn't subdir name compliant

                        if (!File.Exists(inputFilePath))
                        {
                            var f = File.Create(inputFilePath);
                            f.Close();
                        }

                        //Each worksheet symbolizes a test case
                        foreach (Worksheet sheet in ex.Wb.Worksheets)
                        {
                            //Select current sheet
                            ex.ChangeSheet(sheet.Index, false);

                            //Prepare clear input file
                            File.WriteAllText(inputFilePath, string.Empty);

                            //Initialize cell trackers
                            //move each tracker down, as first cell in each column is reserved for its title

                            //TEST CASE
                            Cell input = new Cell(0, 0);
                            input.Down();

                            Cell expectedOut = new Cell(0, 1);
                            expectedOut.Down();

                            //RESULT 
                            Cell actualOut = new Cell(0, 2);
                            actualOut.Down();

                            Cell result = new Cell(0, 3);
                            result.Down();

                            Cell exception = new Cell(0, 4);
                            exception.Down();

                            string temp1;
                            string temp2;
                            string temp3;

                            while ((temp1 = ex.ReadCell(actualOut)) != "" || (temp2 = ex.ReadCell(result)) != "" || (temp3 = ex.ReadCell(exception)) != "")
                            {
                                ex.AlterCell(actualOut, string.Empty);
                                ex.AlterCell(result, string.Empty);
                                ex.AlterCell(exception, string.Empty);

                                actualOut.Down();
                                result.Down();
                                exception.Down();
                            }

                            actualOut.ZeroIndexSet(1, 2);
                            result.ZeroIndexSet(1, 3);
                            exception.ZeroIndexSet(1, 4);

                            using (System.IO.StreamWriter file = new System.IO.StreamWriter(inputFilePath))
                            {
                                string inputLine = ex.ReadCell(input);
                                while (inputLine != "")
                                {
                                    file.WriteLine(inputLine);
                                    input.Down();
                                    inputLine = ex.ReadCell(input);
                                }

                            }
                            try
                            {
                                Process.Start(startInfo).WaitForExit(5000);
                            }
                            catch (Exception e)
                            {
                                ex.AlterCell(exception, e.Message);
                                exception.Down();
                            }

                            string outputFilePath = fullPath + "\\gailis.out";

                            if (File.Exists(outputFilePath))
                            {
                                using (StreamReader file = new StreamReader(outputFilePath))
                                {
                                    string outputLine;
                                    string expectedLine;
                                    while ((outputLine = file.ReadLine()) != null)
                                    {
                                        ex.AlterCell(actualOut, outputLine);
                                        if ((expectedLine = ex.ReadCell(expectedOut)) == outputLine)
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
                                    }
                                }
                            }
                            else
                            {
                                ex.AlterCell(exception, "Output not found");
                            }
                            ex.Save();
                        }
                        ex.Close();
                    }
                    else Console.WriteLine("exe not found at: " + fullPath);
                }
            }
            else
            {
                Console.WriteLine("No test subdirectories specified in: " + di.FullName);
            }
        }
    }
}
