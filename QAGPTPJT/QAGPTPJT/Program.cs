using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Timers;
using System.Diagnostics;
using System.Threading;
using System.Drawing;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing;
using ViDi2;
using ViDi2.Local;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

// Note : JK modify code 20023.05.04 : processing Red Tool(HDM, FSu, FUn)  and Saving result and Chart in Excel

namespace QAGPTPJT
{
    static class Constants
    {
        public const int RepeatProcess = 1000;
    }

    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine($"*** QA-Get Processing Time:" + DateTime.Now.ToString("yyyy-MM-dd") + " ***\n");
            Console.WriteLine($"\nStep 1. Start getting processing time");
            using (ViDi2.Runtime.Local.Control control = new ViDi2.Runtime.Local.Control(GpuMode.Deferred))
            {
                control.InitializeComputeDevices(GpuMode.SingleDevicePerTool, new List<int>() { });
                Stopwatch stopWatch = new Stopwatch();
                // Red HDM - Start
                List<string> RedHDMTimeList = new List<string>();
                string pathRuntime_Red_HDM = "..\\..\\..\\..\\..\\TestResource\\Runtime\\1_RED_HighDetailMode.vrws";
                Console.WriteLine(" - Runtime Path: {0}", pathRuntime_Red_HDM);// Index: control.ComputeDevices[0].Index.ToString()
                ViDi2.Runtime.IWorkspace workspaceRedHDM = control.Workspaces.Add("workspace0", pathRuntime_Red_HDM);
                IStream streamRedHDM = workspaceRedHDM.Streams["default"];
                ITool RedHDMTool = streamRedHDM.Tools["Analyze"];
                var RedHDMParam = RedHDMTool.ParametersBase as ViDi2.Runtime.IToolParametersHighDetail;
                //RedHDMParam.ProcessTensorRT = true or false;                
                string pathRedImagesRedHDM = "..\\..\\..\\..\\..\\TestResource\\Images";
                var extRedHDM = new System.Collections.Generic.List<string> { ".jpg", ".bmp", ".png" };
                var myImagesFilesRedHDM = Directory.GetFiles(pathRedImagesRedHDM, "*.*", SearchOption.TopDirectoryOnly).Where(s => extRedHDM.Any(e => s.EndsWith(e)));
                Console.WriteLine("First Image info. : " + myImagesFilesRedHDM.ElementAt(0));
                long sumRedHDM = 0;
                int countRedHDM = 0;
                var fileRedHDM = myImagesFilesRedHDM.ElementAt(0);
                for (int repeatcnt = 0; repeatcnt < Constants.RepeatProcess; repeatcnt++)
                {
                    countRedHDM++;
                    using (IImage image = new LibraryImage(fileRedHDM))
                    {
                        using (ISample sample = streamRedHDM.CreateSample(image))
                        {
                            stopWatch.Start();
                            sample.Process(RedHDMTool);
                            stopWatch.Stop();
                            sumRedHDM += stopWatch.ElapsedMilliseconds;
                            RedHDMTimeList.Add(stopWatch.ElapsedMilliseconds.ToString());
                            stopWatch.Reset();
                        }
                    }
                }
                double avgRedHDM = sumRedHDM / (double)countRedHDM;
                Console.WriteLine(" - Processing Time Average({0} images): {1} [msec]", (int)countRedHDM, avgRedHDM);
                // Red HDM - End

                // Red Focused Supervised - Start                
                List<string> RedFSuTimeList = new List<string>();
                string pathRuntime_Red_FSu = "..\\..\\..\\..\\..\\TestResource\\Runtime\\2_RED_FocusedSupervised.vrws";
                Console.WriteLine(" - Runtime Path: {0}", pathRuntime_Red_FSu);
                ViDi2.Runtime.IWorkspace workspaceRedFSu = control.Workspaces.Add("workspace1", pathRuntime_Red_FSu);
                IStream streamRedFSu = workspaceRedFSu.Streams["default"];
                ITool RedFSuTool = streamRedFSu.Tools["Analyze"];
                var RedFSuParam = RedFSuTool.ParametersBase as ViDi2.Runtime.IRedTool;
                string pathRedImagesRedFSu = "..\\..\\..\\..\\..\\TestResource\\Images";
                var extRedFSu = new System.Collections.Generic.List<string> { ".jpg", ".bmp", ".png" };
                var myImagesFilesRedFSu = Directory.GetFiles(pathRedImagesRedFSu, "*.*", SearchOption.TopDirectoryOnly).Where(s => extRedFSu.Any(e => s.EndsWith(e)));
                Console.WriteLine("First Image info. : " + myImagesFilesRedFSu.ElementAt(0));
                long sumRedFSu = 0;
                int countRedFSu = 0;
                var fileRedFSu = myImagesFilesRedFSu.ElementAt(0);
                for (int repeatcnt = 0; repeatcnt < Constants.RepeatProcess; repeatcnt++)
                {
                    countRedFSu++;
                    using (IImage image = new LibraryImage(fileRedFSu))
                    {
                        using (ISample sample = streamRedFSu.CreateSample(image))
                        {
                            stopWatch.Start();
                            sample.Process(RedFSuTool);
                            stopWatch.Stop();
                            sumRedFSu += stopWatch.ElapsedMilliseconds;
                            RedFSuTimeList.Add(stopWatch.ElapsedMilliseconds.ToString());
                            stopWatch.Reset();
                        }
                    }
                }
                double avgRedFSu = sumRedFSu / (double)countRedFSu;
                Console.WriteLine(" - Processing Time Average({0} images): {1} [msec]", (int)countRedFSu, avgRedFSu);
                // Red Focused Supervised - End

                // Red Focused Unsupervised - Start
                List<string> RedFUnTimeList = new List<string>();
                string pathRuntime_Red_FUn = "..\\..\\..\\..\\..\\TestResource\\Runtime\\3_RED_FocusedUnsupervised.vrws";
                Console.WriteLine(" - Runtime Path: {0}", pathRuntime_Red_FUn);
                ViDi2.Runtime.IWorkspace workspaceRedFUn = control.Workspaces.Add("workspace2", pathRuntime_Red_FUn);
                IStream streamRedFUn = workspaceRedFUn.Streams["default"];
                ITool RedFUnTool = streamRedFUn.Tools["Analyze"];
                var RedFUnParam = RedFUnTool.ParametersBase as ViDi2.Runtime.IRedTool;
                string pathRedImagesRedFUn = "..\\..\\..\\..\\..\\TestResource\\Images";
                var extRedFUn = new System.Collections.Generic.List<string> { ".jpg", ".bmp", ".png" };
                var myImagesFilesRedFUn = Directory.GetFiles(pathRedImagesRedFUn, "*.*", SearchOption.TopDirectoryOnly).Where(s => extRedFUn.Any(e => s.EndsWith(e)));
                Console.WriteLine("First Image info. : " + myImagesFilesRedFUn.ElementAt(0));
                long sumRedFUn = 0;
                int countRedFUn = 0;
                var fileRedFUn = myImagesFilesRedFUn.ElementAt(0);
                for (int repeatcnt = 0; repeatcnt < Constants.RepeatProcess; repeatcnt++)
                {
                    countRedFUn++;
                    using (IImage image = new LibraryImage(fileRedFUn))
                    {
                        using (ISample sample = streamRedFUn.CreateSample(image))
                        {
                            stopWatch.Start();
                            sample.Process(RedFUnTool);
                            stopWatch.Stop();
                            sumRedFUn += stopWatch.ElapsedMilliseconds;
                            RedFUnTimeList.Add(stopWatch.ElapsedMilliseconds.ToString());
                            stopWatch.Reset();
                        }
                    }
                }
                double avgRedFUn = sumRedFUn / (double)countRedFUn;
                Console.WriteLine(" - Processing Time Average({0} images): {1} [msec]", (int)countRedFUn, avgRedFUn);

                // Step 3. Finish the getting processing time ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                Console.WriteLine($"\nStep 2. Finish the getting processing time");
                string strDateGetResult = DateTime.Now.ToString("yyyy-MM-dd");
                string csvFileName = "GetProcessingTime_" + strDateGetResult + ".csv";

                // EPPlus Excel - 20230426                //var getResultList = new List<string>();
                var getResultListRedHDM = new List<string>();
                var getResultListRedFSu = new List<string>();
                var getResultListRedFUn = new List<string>();

                using (System.IO.StreamWriter resultFile = new System.IO.StreamWriter(@"..\..\..\..\..\TestResultCSV\" + csvFileName, true, System.Text.Encoding.GetEncoding("utf-8")))
                {
                    resultFile.WriteLine("ImagePath, SpendingTime");
                    for (int indexcnt = 0; indexcnt < Constants.RepeatProcess; indexcnt++)
                    {
                        resultFile.WriteLine("{0}, {1}, {2}, {3}", myImagesFilesRedHDM.ElementAt(0), RedHDMTimeList[indexcnt].ToString(), RedFSuTimeList[indexcnt].ToString(), RedFUnTimeList[indexcnt].ToString()); //resultFile.WriteLine("{0}, {1}", myImagesFiles, resultTime.ToString());
                        getResultListRedHDM.Add(RedHDMTimeList[indexcnt].ToString());
                        getResultListRedFSu.Add(RedFSuTimeList[indexcnt].ToString());
                        getResultListRedFUn.Add(RedFUnTimeList[indexcnt].ToString());
                    }
                }
                Console.WriteLine(" - Result CSV File: {0}", csvFileName);

                Console.WriteLine("\nStep 3. Save resultin Excel file");
                string getDateInfo = DateTime.Now.ToString("yyyy-MM-dd"); // refer to //www.delftstack.com/ko/howto/csharp/how-to-get-the-current-date-without-time-in-csharp/
                string strExcelFileName = "QAGetProcessingTime_" + getDateInfo + ".xlsx";
                string strExcelFileDirectory = Path.GetFullPath(@"..\..\..\..\..\TestResultCSV\") + strExcelFileName;   // Refer to - Processing file path name in using C# : //myoung-min.tistory.com/45
                Console.WriteLine(strExcelFileDirectory);

                ExcelDataEPPlusRedTools(getResultListRedHDM, getResultListRedFSu, getResultListRedFUn, strExcelFileDirectory); // create epplus excel  - 20230504
                Console.WriteLine("\nStep 4. Complete QA Test - Get Processing time of Red Tool");
            }
        }

        private static void ExcelDataEPPlusRedTools(List<string> GetPTimesRedHDM, List<string> GetPTimesRedFSu, List<string> GetPTimesRedFUn, string savePath)
        {
            Console.WriteLine("JK Test 1. Create Excel File");
            ExcelPackage ExcelPkg = new ExcelPackage();
            ExcelWorksheet wsSheet1 = ExcelPkg.Workbook.Worksheets.Add("Sheet1");
            using (ExcelRange Rng = wsSheet1.Cells[1, 1, 1, 1])  // 1x1
            {
                //Rng.Value = "QA JK's Task : Get processing time in teamcity!";
                Rng.Value = "Repeat";
                Rng.Style.Font.Size = 11; //16;
                Rng.Style.Font.Bold = true;
                Rng.Style.Font.Italic = true;
            }
            using (ExcelRange Rng = wsSheet1.Cells[1, 2, 1, 2])
            {
                //Rng.Value = "QA JK's Task : Get processing time in teamcity!";
                Rng.Value = "Red HDM";
                Rng.Style.Font.Size = 11; //16;
                Rng.Style.Font.Bold = true;
                Rng.Style.Font.Italic = true;
            }
            using (ExcelRange Rng = wsSheet1.Cells[1, 3, 1, 3])
            {
                //Rng.Value = "QA JK's Task : Get processing time in teamcity!";
                Rng.Value = "Red FSu";
                Rng.Style.Font.Size = 11; //16;
                Rng.Style.Font.Bold = true;
                Rng.Style.Font.Italic = true;
            }
            using (ExcelRange Rng = wsSheet1.Cells[1, 4, 1, 4])
            {
                //Rng.Value = "QA JK's Task : Get processing time in teamcity!";
                Rng.Value = "Red FUn";
                Rng.Style.Font.Size = 11; //16;
                Rng.Style.Font.Bold = true;
                Rng.Style.Font.Italic = true;
            }
            wsSheet1.Protection.IsProtected = false;
            wsSheet1.Protection.AllowSelectLockedCells = false;
            ExcelPkg.SaveAs(new FileInfo(@savePath));
            Console.WriteLine(" - Complete the creating excel file!");

            Console.WriteLine("JK Test 2. Adding Chart after loading the created excel.");

            string pathExcelFile = savePath;
            Console.WriteLine(" - Load ExcelInfo: {0}", pathExcelFile);

            FileInfo existingFile = new FileInfo(pathExcelFile);
            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets["Sheet1"];
                int column = 1;
                for (int row = 2; row < (Constants.RepeatProcess + 2); row++) // using repeat 100 times = 102
                    worksheet.Cells[row, column].Value = row - 1;
                int col = 2;    // Red HDM
                for (int row = 2; row < (Constants.RepeatProcess + 2); row++) // using repeat 100 times = 102
                {
                    worksheet.Cells[row, col].Value = int.Parse(GetPTimesRedHDM[row - 2]);
                }
                col = 3;        // Red Focused Supervised
                for (int row = 2; row < (Constants.RepeatProcess + 2); row++) // using repeat 100 times = 102
                {
                    worksheet.Cells[row, col].Value = int.Parse(GetPTimesRedFSu[row - 2]);
                }
                col = 4;        // Red Focused Unsupervised
                for (int row = 2; row < (Constants.RepeatProcess + 2); row++)   // using repeat 100 times = 102
                {
                    worksheet.Cells[row, col].Value = int.Parse(GetPTimesRedFUn[row - 2]);
                }

                var chart = worksheet.Drawings.AddChart("Chart1", eChartType.Line);
                chart.Title.Text = "Processing Time Red Tool(HDM/FSu/FUn)"; ////From row 1 colum 5 with five pixels offset                
                chart.Title.Font.Size = 14;
                chart.Title.Font.Bold = true;
                chart.Title.Font.Italic = true;
                chart.SetPosition(1, 1, 6, 6); // Start point to dispale of Chart  ex) 0,0,5,5 : Draw a chart from F1 Cell vs 1,1,6,6 : Draw a chart from G2 Cell
                chart.SetSize(900, 600);

                ExcelAddress valueAddress_Data1 = new ExcelAddress(2, 2, (Constants.RepeatProcess + 1), 2); // using repeat 100 times
                ExcelAddress RepeatAddress_Data1 = new ExcelAddress(2, 1, (Constants.RepeatProcess + 1), 1);
                var ser1 = (chart.Series.Add(valueAddress_Data1.Address, RepeatAddress_Data1.Address) as ExcelLineChartSerie); // using repeat 100 time
                ser1.Header = "Red HDM";

                ExcelAddress valueAddress_Data2 = new ExcelAddress(2, 3, (Constants.RepeatProcess + 1), 3); // using repeat 100 times
                ExcelAddress RepeatAddress_Data2 = new ExcelAddress(2, 1, (Constants.RepeatProcess + 1), 1);
                var ser2 = (chart.Series.Add(valueAddress_Data2.Address, RepeatAddress_Data2.Address) as ExcelLineChartSerie); // using repeat 100 time
                ser2.Header = "Red FSu";

                ExcelAddress valueAddress_Data3 = new ExcelAddress(2, 4, (Constants.RepeatProcess + 1), 4); // using repeat 100 times
                ExcelAddress RepeatAddress_Data3 = new ExcelAddress(2, 1, (Constants.RepeatProcess + 1), 1);
                var ser3 = (chart.Series.Add(valueAddress_Data3.Address, RepeatAddress_Data3.Address) as ExcelLineChartSerie); // using repeat 100 times
                ser3.Header = "Red FUn";

                chart.Legend.Border.LineStyle = eLineStyle.Solid;
                chart.Legend.Border.Fill.Style = eFillStyle.SolidFill;
                chart.Legend.Border.Fill.Color = Color.DarkBlue;
                package.Save();
            }
            Console.WriteLine("Complete - adding chart with using EPPlus.4.5.3.3");
            Console.WriteLine();
        }
    }
}

