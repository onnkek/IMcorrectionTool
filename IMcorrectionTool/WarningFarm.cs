using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using System.Drawing;
using System.Windows.Controls;
using System.Text.RegularExpressions;

namespace IMcorrectionTool
{
    class WarningFarm
    {



        public static List<Warming> GetWarningListFromCduFormat(string PathToCSV)
        {
            List<Warming> warnings = new List<Warming>();
            const Int32 BufferSize = 128;
            Encoding win1251 = Encoding.GetEncoding(1251);

            using (FileStream fstream = File.OpenRead(PathToCSV))
            {
                using (var streamReader = new StreamReader(fstream, win1251, true, BufferSize))
                {
                    String line;
                    while ((line = streamReader.ReadLine()) != null)
                    {
                        var splittedLine = line.Split(';');
                        if (splittedLine[0] == "ОДУ Урала")
                        {
                            warnings.Add(new Warming(splittedLine[0], splittedLine[1], splittedLine[2], splittedLine[3], splittedLine[4], splittedLine[5], splittedLine[6], splittedLine[7]));
                        }
                    }
                }

            }

            return warnings;
        }
        public static List<Warming> GetWarningListFromCduFormatExcel(string PathToExcel)
        {
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(PathToExcel);
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;


            List<Warming> warnings = new List<Warming>();

            // Find the last real row
            var lastUsedRow = xlWorksheet.Cells.Find("*", System.Reflection.Missing.Value,
                                           System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                                           Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious,
                                           false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;

            // Find the last real column
            var lastUsedColumn = xlWorksheet.Cells.Find("*", System.Reflection.Missing.Value,
                                           System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                                           Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlPrevious,
                                           false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Column;

            Excel.Range range = xlWorksheet.get_Range("A1", "H" + lastUsedRow.ToString());
            System.Array myvalues = (System.Array)range.Cells.Value;
            //string[] strArray = ConvertToStringArray(myvalues);

            for (int i = 1; i <= lastUsedRow; i++)
            {
                MainWindow.dispatcher.Invoke(MainWindow.updProgress, new object[] { ProgressBar.ValueProperty, ++MainWindow.value });
                if (myvalues.GetValue(i, 1) != null && myvalues.GetValue(i, 1).ToString() == "ОДУ Урала")
                {
                    string[] splittedLine = new string[8];
                    for (int j = 1; j <= lastUsedColumn; j++)
                    {
                        if (myvalues.GetValue(i, j) == null)
                        {
                            splittedLine[j - 1] = "";
                        }
                        else splittedLine[j - 1] = myvalues.GetValue(i, j).ToString();
                    }
                    warnings.Add(new Warming(splittedLine[0], splittedLine[1], splittedLine[2], splittedLine[3], splittedLine[4], splittedLine[5], splittedLine[6], splittedLine[7]));
                }
            }



            //for (int i = 1; i <= lastUsedRow; i++)
            //{
            //    string[] splittedLine = new string[8];
            //    if (xlRange.Cells[i, 1] != null && xlRange.Cells[i, 1].Value2 != null && xlRange.Cells[i, 1].Value2 == "ОДУ Урала")
            //    {
            //        for (int j = 1; j <= lastUsedColumn; j++)
            //        {
            //            //new line


            //            //write the value to the console
            //            if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
            //                splittedLine[j - 1] = xlRange.Cells[i, j].Value2.ToString();

            //            //add useful things here!   
            //        }
            //    }
            //    if (splittedLine[0] == "ОДУ Урала")
            //            {
            //                warnings.Add(new Warming(splittedLine[0], splittedLine[1], splittedLine[2], splittedLine[3], splittedLine[4], splittedLine[5], splittedLine[6], splittedLine[7]));
            //            }
            //}

            //const Int32 BufferSize = 128;
            //Encoding win1251 = Encoding.GetEncoding(1251);

            //using (FileStream fstream = File.OpenRead(PathToCSV))
            //{
            //    using (var streamReader = new StreamReader(fstream, win1251, true, BufferSize))
            //    {
            //        String line;
            //        while ((line = streamReader.ReadLine()) != null)
            //        {
            //            var splittedLine = line.Split(';');
            //            if (splittedLine[0] == "ОДУ Урала")
            //            {
            //                warnings.Add(new Warming(splittedLine[0], splittedLine[1], splittedLine[2], splittedLine[3], splittedLine[4], splittedLine[5], splittedLine[6], splittedLine[7]));
            //            }
            //        }
            //    }

            //}


            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //rule of thumb for releasing com objects:
            //  never use two dots, all COM objects must be referenced and released individually
            //  ex: [somthing].[something].[something] is bad

            //release com objects to fully kill excel process from running in the background
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            //close and release
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            //quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);



            return warnings;
        }

        static List<string[]> ConvertToStrinпList(System.Array values)
        {
            List<string[]> stringList = new List<string[]>();
            string[] theArray = new string[values.Length - 1];
            return stringList;
        }
        static string[] ConvertToStringArray(System.Array values)
        {

            // create a new string array
            string[] theArray = new string[values.Length];

            // loop through the 2-D System.Array and populate the 1-D String Array
            for (int i = 1; i <= values.Length; i++)
            {
                if (values.GetValue(1, i) == null)
                    theArray[i - 1] = "";
                else
                    theArray[i - 1] = (string)values.GetValue(1, i).ToString();
            }

            return theArray;
        }
        public static List<Warming> GetWarningListFromCK11Format(string PathToCSV)
        {
            List<Warming> warnings = new List<Warming>();
            const Int32 BufferSize = 128;
            Encoding win1251 = Encoding.GetEncoding(1251);
            using (FileStream fstream = File.OpenRead(PathToCSV))
            {
                using (var streamReader = new StreamReader(fstream, win1251, true, BufferSize))
                {
                    String line;
                    while ((line = streamReader.ReadLine()) != null)
                    {
                        MainWindow.dispatcher.Invoke(MainWindow.updProgress, new object[] { ProgressBar.ValueProperty, ++MainWindow.value });
                        var splittedLine = line.Split(';');
                        if (splittedLine[9] == "ОДУ Урала" || splittedLine[9] == "Тюменское РДУ" || splittedLine[9] == "Челябинское РДУ" || splittedLine[9] == "Свердловское РДУ" || splittedLine[9] == "Пермское РДУ" || splittedLine[9] == "Оренбургское РДУ" || splittedLine[9] == "Башкирское РДУ")
                        {
                            warnings.Add(new Warming("ОДУ Урала", splittedLine[9], splittedLine[2], splittedLine[3], splittedLine[5], splittedLine[6], splittedLine[7]));
                        }
                    }
                }
            }
            return warnings;
        }
        public static void SaveToExcelBasedOnCurrentMonth(string PathToCurrentMonth, string PathToSave, List<Warming> WarningList)
        {
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(PathToCurrentMonth);
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;


            List<Warming> warnings = new List<Warming>();

            // Find the last real row
            var lastUsedRow = xlWorksheet.Cells.Find("*", System.Reflection.Missing.Value,
                                           System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                                           Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious,
                                           false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;

            // Find the last real column
            var lastUsedColumn = xlWorksheet.Cells.Find("*", System.Reflection.Missing.Value,
                                           System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                                           Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlPrevious,
                                           false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Column;

            Excel.Range range = xlWorksheet.get_Range("A1", "H" + lastUsedRow.ToString());

            Excel.Range writeRange = xlWorksheet.get_Range("H1", "H" + lastUsedRow.ToString());
            string[,] writeValues = new string[lastUsedRow, 1];

            System.Array myvalues = (System.Array)range.Cells.Value;
            //string[] strArray = ConvertToStringArray(myvalues);

            for (int i = 1; i <= lastUsedRow; i++)
            {
                MainWindow.dispatcher.Invoke(MainWindow.updProgress, new object[] { ProgressBar.ValueProperty, ++MainWindow.value });
                if (myvalues.GetValue(i, 1) != null && myvalues.GetValue(i, 1).ToString() == "ОДУ Урала")
                {
                    string resultText = myvalues.GetValue(i, 7).ToString();
                    // Выборка строк, в которых есть id. Необходимо для ограничения количества
                    // строк, прогоняемых через регулярное выражение, ибо оно работает очень медленно.
                    if (myvalues.GetValue(i, 7).ToString().Contains("Id") || 
                        myvalues.GetValue(i, 7).ToString().Contains("id") || 
                        myvalues.GetValue(i, 7).ToString().Contains("ID"))
                        resultText = Regex.Replace(myvalues.GetValue(i, 7).ToString(), @"\W*id\W*[0-9]*\W*", "Id=", RegexOptions.IgnoreCase).Trim();
                    var id = myvalues.GetValue(i, 4).ToString() + resultText;
                    var wrn = WarningList.FirstOrDefault(x => x.ID == id);
                    if (wrn != null)
                    {
                        writeValues[i - 1, 0] = wrn.PreviousComment;
                        if (!string.IsNullOrEmpty(wrn.Comment))
                            writeValues[i - 1, 0] = wrn.Comment;
                        if (wrn.Comment != "Устранено" && wrn.IsNewInMonth == false)
                            xlWorksheet.get_Range("A" + (i).ToString(), "H" + (i).ToString()).Interior.Color = ColorTranslator.ToOle(Color.Plum);
                    }
                }
            }
            writeRange.Value2 = writeValues;
            // writeRange.Value = 
            var newWarnings = WarningList.Where(x => x.IsNewInKGID).ToList();

            for (int i = 1; i <= newWarnings.Count(); i++)
            {
                MainWindow.dispatcher.Invoke(MainWindow.updProgress, new object[] { ProgressBar.ValueProperty, ++MainWindow.value });
                xlWorksheet.Cells[i + lastUsedRow, 1].Value2 = newWarnings[i - 1].ODU;
                xlWorksheet.Cells[i + lastUsedRow, 2].Value2 = newWarnings[i - 1].ModelingAuthoritySet;
                xlWorksheet.Cells[i + lastUsedRow, 3].Value2 = newWarnings[i - 1].RuleID;
                xlWorksheet.Cells[i + lastUsedRow, 4].Value2 = newWarnings[i - 1].ObjectUID;
                xlWorksheet.Cells[i + lastUsedRow, 5].Value2 = newWarnings[i - 1].ObjectName;
                xlWorksheet.Cells[i + lastUsedRow, 6].Value2 = newWarnings[i - 1].ObjectName;
                xlWorksheet.Cells[i + lastUsedRow, 7].Value2 = newWarnings[i - 1].WarningText;
                xlWorksheet.Cells[i + lastUsedRow, 8].Value2 = newWarnings[i - 1].PreviousComment;
                if (!string.IsNullOrEmpty(newWarnings[i - 1].Comment))
                {
                    xlWorksheet.Cells[i + lastUsedRow, 8].Value2 = newWarnings[i - 1].Comment;
                }


                xlWorksheet.get_Range("A" + (i + lastUsedRow).ToString(), "H" + (i + lastUsedRow).ToString()).Interior.Color = ColorTranslator.ToOle(Color.Yellow);
            }

            xlWorksheet.get_Range("A1", "H" + (lastUsedRow + newWarnings.Count).ToString()).Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            //for (int i = 1; i <= lastUsedRow; i++)
            //{
            //    string[] splittedLine = new string[8];
            //    if (xlRange.Cells[i, 1] != null && xlRange.Cells[i, 1].Value2 != null && xlRange.Cells[i, 1].Value2 == "ОДУ Урала")
            //    {
            //        for (int j = 1; j <= lastUsedColumn; j++)
            //        {
            //            //new line


            //            //write the value to the console
            //            if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
            //                splittedLine[j - 1] = xlRange.Cells[i, j].Value2.ToString();

            //            //add useful things here!   
            //        }
            //    }
            //    if (splittedLine[0] == "ОДУ Урала")
            //            {
            //                warnings.Add(new Warming(splittedLine[0], splittedLine[1], splittedLine[2], splittedLine[3], splittedLine[4], splittedLine[5], splittedLine[6], splittedLine[7]));
            //            }
            //}

            //const Int32 BufferSize = 128;
            //Encoding win1251 = Encoding.GetEncoding(1251);

            //using (FileStream fstream = File.OpenRead(PathToCSV))
            //{
            //    using (var streamReader = new StreamReader(fstream, win1251, true, BufferSize))
            //    {
            //        String line;
            //        while ((line = streamReader.ReadLine()) != null)
            //        {
            //            var splittedLine = line.Split(';');
            //            if (splittedLine[0] == "ОДУ Урала")
            //            {
            //                warnings.Add(new Warming(splittedLine[0], splittedLine[1], splittedLine[2], splittedLine[3], splittedLine[4], splittedLine[5], splittedLine[6], splittedLine[7]));
            //            }
            //        }
            //    }

            //}
            //Save
            if (!System.IO.File.Exists(PathToSave))
            {
                xlWorkbook.SaveAs(PathToSave);
            }
            else
            {
                xlWorkbook.SaveAs(PathToSave + "Copy.xlsx");
            }

            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //rule of thumb for releasing com objects:
            //  never use two dots, all COM objects must be referenced and released individually
            //  ex: [somthing].[something].[something] is bad

            //release com objects to fully kill excel process from running in the background
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            //close and release
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            //quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);


            

        }
    }
}
