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
// Note : JK modify code 20023.05.08 : Adding Green Tools(HDM, Focused, HDM Quick), This data save in Excel file as above.

namespace QAGPTPJT
{
    static class Constants
    {
        public const int RepeatProcess = 5000;
        //public const int RepeatProcess = 3;
    }

    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine($"*** QA-Get Processing Time:" + DateTime.Now.ToString("yyyy-MM-dd") + " ***\n");
            Console.WriteLine($"\nStep 1. Start getting processing time");
            using (ViDi2.Runtime.Local.Control control = new ViDi2.Runtime.Local.Control(GpuMode.Deferred))
            {
                Console.WriteLine($"\n - initialize GPU Device");
                control.InitializeComputeDevices(GpuMode.SingleDevicePerTool, new List<int>() { });
                Stopwatch stopWatch = new Stopwatch();

                // Blue Locate - Start // BlueLocate
                Console.WriteLine($"\n - Blue Locate - Start");
                List<string> BlueLocateTimeList = new List<string>();
                string pathRuntime_BlueLocate = "..\\..\\..\\..\\..\\TestResource\\Runtime\\6_BlueLocate.vrws";
                Console.WriteLine(" - Runtime Path: {0}", pathRuntime_BlueLocate);// Index: control.ComputeDevices[0].Index.ToString()
                ViDi2.Runtime.IWorkspace workspaceBlueLocate = control.Workspaces.Add("workspaceBlueLocate", pathRuntime_BlueLocate);
                IStream streamBlueLocate = workspaceBlueLocate.Streams["default"];
                ITool BlueLocateTool = streamBlueLocate.Tools["Locate"];
                var BlueLocateParam = BlueLocateTool.ParametersBase as ViDi2.Runtime.IBlueTool;
                string pathBlueImagesBlueLocate = "..\\..\\..\\..\\..\\TestResource\\Images_BlueLocate";
                var extBlueLocate = new System.Collections.Generic.List<string> { ".jpg", ".bmp", ".png" };
                var myImagesFilesBlueLocate = Directory.GetFiles(pathBlueImagesBlueLocate, "*.*", SearchOption.TopDirectoryOnly).Where(s => extBlueLocate.Any(e => s.EndsWith(e)));
                Console.WriteLine("First Image info. : " + myImagesFilesBlueLocate.ElementAt(0));
                long sumBlueLocate = 0;
                int countBlueLocate = 0;
                var fileBlueLocate = myImagesFilesBlueLocate.ElementAt(0);
                for (int repeatcnt = 0; repeatcnt < Constants.RepeatProcess; repeatcnt++)
                {
                    countBlueLocate++;
                    using (IImage image = new LibraryImage(fileBlueLocate))
                    {
                        using (ISample sample = streamBlueLocate.CreateSample(image))
                        {
                            stopWatch.Start();
                            sample.Process(BlueLocateTool);
                            stopWatch.Stop();
                            sumBlueLocate += stopWatch.ElapsedMilliseconds;
                            BlueLocateTimeList.Add(stopWatch.ElapsedMilliseconds.ToString());
                            stopWatch.Reset();
                        }
                    }
                }
                double avgBlueLocate = sumBlueLocate / (double)countBlueLocate;
                Console.WriteLine(" - Processing Time Average({0} images): {1} [msec]", (int)countBlueLocate, avgBlueLocate);
                // Blue Locate - End

                // Blue Read - Start // BlueRead
                Console.WriteLine($"\n - Blue Read - Start");
                List<string> BlueReadTimeList = new List<string>();
                string pathRuntime_BlueRead = "..\\..\\..\\..\\..\\TestResource\\Runtime\\7_BlueRead.vrws";
                Console.WriteLine(" - Runtime Path: {0}", pathRuntime_BlueRead);// Index: control.ComputeDevices[0].Index.ToString()
                ViDi2.Runtime.IWorkspace workspaceBlueRead = control.Workspaces.Add("workspaceBlueRead", pathRuntime_BlueRead);
                IStream streamBlueRead = workspaceBlueRead.Streams["default"];
                ITool BlueReadTool = streamBlueRead.Tools["Read"];
                var BlueReadParam = BlueReadTool.ParametersBase as ViDi2.Runtime.IBlueTool;
                string pathBlueImagesBlueRead = "..\\..\\..\\..\\..\\TestResource\\Images_BlueRead";
                var extBlueRead = new System.Collections.Generic.List<string> { ".jpg", ".bmp", ".png" };
                var myImagesFilesBlueRead = Directory.GetFiles(pathBlueImagesBlueRead, "*.*", SearchOption.TopDirectoryOnly).Where(s => extBlueRead.Any(e => s.EndsWith(e)));
                Console.WriteLine("First Image info. : " + myImagesFilesBlueRead.ElementAt(0));
                long sumBlueRead = 0;
                int countBlueRead = 0;
                var fileBlueRead = myImagesFilesBlueRead.ElementAt(0);
                for (int repeatcnt = 0; repeatcnt < Constants.RepeatProcess; repeatcnt++)
                {
                    countBlueRead++;
                    using (IImage image = new LibraryImage(fileBlueRead))
                    {
                        using (ISample sample = streamBlueRead.CreateSample(image))
                        {
                            stopWatch.Start();
                            sample.Process(BlueReadTool);
                            stopWatch.Stop();
                            sumBlueRead += stopWatch.ElapsedMilliseconds;
                            BlueReadTimeList.Add(stopWatch.ElapsedMilliseconds.ToString());
                            stopWatch.Reset();
                        }
                    }
                }
                double avgBlueRead = sumBlueRead / (double)countBlueRead;
                Console.WriteLine(" - Processing Time Average({0} images): {1} [msec]", (int)countBlueRead, avgBlueRead);

                // Blue Read - End

                // Green HDM - Start
                Console.WriteLine($"\n - Green HDM - Start");
                List<string> GreenHDMTimeList = new List<string>();
                string pathRuntime_Greem_HDM = "..\\..\\..\\..\\..\\TestResource\\Runtime\\1_Green_HighDetailMode.vrws";
                Console.WriteLine(" - Runtime Path: {0}", pathRuntime_Greem_HDM);// Index: control.ComputeDevices[0].Index.ToString()
                ViDi2.Runtime.IWorkspace workspaceGreenHDM = control.Workspaces.Add("workspaceGreenHDM", pathRuntime_Greem_HDM);
                IStream streamGreenHDM = workspaceGreenHDM.Streams["default"];
                ITool GreenHDMTool = streamGreenHDM.Tools["Classify"];
                //var GreenHDMParam = GreenHDMTool.ParametersBase as ViDi2.Runtime.IToolParametersHighDetail; // 기존 실험 적용 코드 - 2023.05.08
                var GreenHDMParam = GreenHDMTool.ParametersBase as ViDi2.Runtime.IGreenHighDetailParameters;

                //RedHDMParam.ProcessTensorRT = true or false;                
                string pathGreenImagesGreenHDM = "..\\..\\..\\..\\..\\TestResource\\Images_Green";
                var extGreenHDM = new System.Collections.Generic.List<string> { ".jpg", ".bmp", ".png" };
                var myImagesFilesGreenHDM = Directory.GetFiles(pathGreenImagesGreenHDM, "*.*", SearchOption.TopDirectoryOnly).Where(s => extGreenHDM.Any(e => s.EndsWith(e)));
                Console.WriteLine("First Image info. : " + myImagesFilesGreenHDM.ElementAt(0));
                long sumGreenHDM = 0;
                int countGreenHDM = 0;
                var fileGreenHDM = myImagesFilesGreenHDM.ElementAt(0);
                for (int repeatcnt = 0; repeatcnt < Constants.RepeatProcess; repeatcnt++)
                {
                    countGreenHDM++;
                    using (IImage image = new LibraryImage(fileGreenHDM))
                    {
                        using (ISample sample = streamGreenHDM.CreateSample(image))
                        {
                            stopWatch.Start();
                            sample.Process(GreenHDMTool);
                            stopWatch.Stop();
                            sumGreenHDM += stopWatch.ElapsedMilliseconds;
                            GreenHDMTimeList.Add(stopWatch.ElapsedMilliseconds.ToString());
                            stopWatch.Reset();
                        }
                    }
                }
                double avgGreenHDM = sumGreenHDM / (double)countGreenHDM;
                Console.WriteLine(" - Processing Time Average({0} images): {1} [msec]", (int)countGreenHDM, avgGreenHDM);
                // Green HDM - End

                // Green Focused - Start GreenFocused
                Console.WriteLine($"\n - Green Focused - Start");
                List<string> GreenFocusedTimeList = new List<string>();
                string pathRuntime_Greem_Focused = "..\\..\\..\\..\\..\\TestResource\\Runtime\\2_Green_FocusedMode.vrws";
                Console.WriteLine(" - Runtime Path: {0}", pathRuntime_Greem_Focused);// Index: control.ComputeDevices[0].Index.ToString()
                ViDi2.Runtime.IWorkspace workspaceGreenFocused = control.Workspaces.Add("workspaceGreenFocused", pathRuntime_Greem_Focused);
                IStream streamGreenFocused = workspaceGreenFocused.Streams["default"];
                ITool GreenFocusedTool = streamGreenFocused.Tools["Classify"];
                //var GreenFocusedParam = GreenFocusedTool.ParametersBase as ViDi2.Runtime.IToolParametersHighDetail; // 기존 실험 적용 코드 1- 2023.05.08
                var GreenFocusedParam = GreenFocusedTool.ParametersBase as ViDi2.Runtime.IGreenTool;
                //var GreenFocusedParam = GreenFocusedTool.ParametersBase as ViDi2.Runtime.ITool;// 기존 실험 적용 코드 2- 2023.05.08

                //RedHDMParam.ProcessTensorRT = true or false;                
                string pathGreenImagesGreenFocused = "..\\..\\..\\..\\..\\TestResource\\Images_Green";
                var extGreenFocused = new System.Collections.Generic.List<string> { ".jpg", ".bmp", ".png" };
                var myImagesFilesGreenFocused = Directory.GetFiles(pathGreenImagesGreenFocused, "*.*", SearchOption.TopDirectoryOnly).Where(s => extGreenFocused.Any(e => s.EndsWith(e)));
                Console.WriteLine("First Image info. : " + myImagesFilesGreenFocused.ElementAt(0));
                long sumGreenFocused = 0;
                int countGreenFocused = 0;
                var fileGreenFocused = myImagesFilesGreenFocused.ElementAt(0);
                for (int repeatcnt = 0; repeatcnt < Constants.RepeatProcess; repeatcnt++)
                {
                    countGreenFocused++;
                    using (IImage image = new LibraryImage(fileGreenFocused))
                    {
                        using (ISample sample = streamGreenFocused.CreateSample(image))
                        {
                            stopWatch.Start();
                            sample.Process(GreenFocusedTool);
                            stopWatch.Stop();
                            sumGreenFocused += stopWatch.ElapsedMilliseconds;
                            GreenFocusedTimeList.Add(stopWatch.ElapsedMilliseconds.ToString());
                            stopWatch.Reset();
                        }
                    }
                }
                double avgGreenFocused = sumGreenFocused / (double)countGreenFocused;
                Console.WriteLine(" - Processing Time Average({0} images): {1} [msec]", (int)countGreenFocused, avgGreenFocused);
                // Green Focused - End

                // Green HDM Qucik - Start
                Console.WriteLine($"\n - Green HDM Quick - Start");
                List<string> GreenHDMQTimeList = new List<string>();
                string pathRuntime_Greem_HDMQ = "..\\..\\..\\..\\..\\TestResource\\Runtime\\3_Green_HighDetailModeQuick.vrws";
                Console.WriteLine(" - Runtime Path: {0}", pathRuntime_Greem_HDMQ);// Index: control.ComputeDevices[0].Index.ToString()
                ViDi2.Runtime.IWorkspace workspaceGreenHDMQ = control.Workspaces.Add("workspaceGreenHDMQ", pathRuntime_Greem_HDMQ);
                IStream streamGreenHDMQ = workspaceGreenHDMQ.Streams["default"];
                ITool GreenHDMQTool = streamGreenHDMQ.Tools["Classify"];
                var GreenHDMQParam = GreenHDMQTool.ParametersBase as ViDi2.Runtime.IToolParametersHighDetail;
                //RedHDMParam.ProcessTensorRT = true or false;                
                string pathGreenImagesGreenHDMQ = "..\\..\\..\\..\\..\\TestResource\\Images_Green";
                var extGreenHDMQ = new System.Collections.Generic.List<string> { ".jpg", ".bmp", ".png" };
                var myImagesFilesGreenHDMQ = Directory.GetFiles(pathGreenImagesGreenHDMQ, "*.*", SearchOption.TopDirectoryOnly).Where(s => extGreenHDMQ.Any(e => s.EndsWith(e)));
                Console.WriteLine("First Image info. : " + myImagesFilesGreenHDMQ.ElementAt(0));
                long sumGreenHDMQ = 0;
                int countGreenHDMQ = 0;
                var fileGreenHDMQ = myImagesFilesGreenHDMQ.ElementAt(0);
                for (int repeatcnt = 0; repeatcnt < Constants.RepeatProcess; repeatcnt++)
                {
                    countGreenHDMQ++;
                    using (IImage image = new LibraryImage(fileGreenHDMQ))
                    {
                        using (ISample sample = streamGreenHDMQ.CreateSample(image))
                        {
                            stopWatch.Start();
                            sample.Process(GreenHDMQTool);
                            stopWatch.Stop();
                            sumGreenHDMQ += stopWatch.ElapsedMilliseconds;
                            GreenHDMQTimeList.Add(stopWatch.ElapsedMilliseconds.ToString());
                            stopWatch.Reset();
                        }
                    }
                }
                double avgGreenHDMQ = sumGreenHDMQ / (double)countGreenHDMQ;
                Console.WriteLine(" - Processing Time Average({0} images): {1} [msec]", (int)countGreenHDMQ, avgGreenHDMQ);
                // Green HDM Quick - End

                // Red HDM - Start
                Console.WriteLine($"\n - Red HDM ");
                List<string> RedHDMTimeList = new List<string>();
                string pathRuntime_Red_HDM = "..\\..\\..\\..\\..\\TestResource\\Runtime\\1_RED_HighDetailMode.vrws";
                Console.WriteLine(" - Runtime Path: {0}", pathRuntime_Red_HDM);// Index: control.ComputeDevices[0].Index.ToString()
                ViDi2.Runtime.IWorkspace workspaceRedHDM = control.Workspaces.Add("workspaceRedHDM", pathRuntime_Red_HDM);
                IStream streamRedHDM = workspaceRedHDM.Streams["default"];
                ITool RedHDMTool = streamRedHDM.Tools["Analyze"];
                var RedHDMParam = RedHDMTool.ParametersBase as ViDi2.Runtime.IToolParametersHighDetail;
                //RedHDMParam.ProcessTensorRT = true or false;                
                string pathRedImagesRedHDM = "..\\..\\..\\..\\..\\TestResource\\Images_Red";
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
                Console.WriteLine($"\n - Red Focused Supervised ");
                List<string> RedFSuTimeList = new List<string>();
                string pathRuntime_Red_FSu = "..\\..\\..\\..\\..\\TestResource\\Runtime\\2_RED_FocusedSupervised.vrws";
                Console.WriteLine(" - Runtime Path: {0}", pathRuntime_Red_FSu);
                ViDi2.Runtime.IWorkspace workspaceRedFSu = control.Workspaces.Add("workspaceRedFSu", pathRuntime_Red_FSu);
                IStream streamRedFSu = workspaceRedFSu.Streams["default"];
                ITool RedFSuTool = streamRedFSu.Tools["Analyze"];
                //var RedFSuParam = RedFSuTool.ParametersBase as ViDi2.Runtime.IRedTool; // 기존에 적용했던 코드 2023.05.08
                var RedFSuParam = RedFSuTool.ParametersBase as ViDi2.Runtime.IRedTool;
                string pathRedImagesRedFSu = "..\\..\\..\\..\\..\\TestResource\\Images_Red";
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
                Console.WriteLine($"\n - Red Focused Unsupervised ");
                List<string> RedFUnTimeList = new List<string>();
                string pathRuntime_Red_FUn = "..\\..\\..\\..\\..\\TestResource\\Runtime\\3_RED_FocusedUnsupervised.vrws";
                Console.WriteLine(" - Runtime Path: {0}", pathRuntime_Red_FUn);
                ViDi2.Runtime.IWorkspace workspaceRedFUn = control.Workspaces.Add("workspaceRedFUn", pathRuntime_Red_FUn);
                IStream streamRedFUn = workspaceRedFUn.Streams["default"];
                ITool RedFUnTool = streamRedFUn.Tools["Analyze"];
                var RedFUnParam = RedFUnTool.ParametersBase as ViDi2.Runtime.IRedTool;
                string pathRedImagesRedFUn = "..\\..\\..\\..\\..\\TestResource\\Images_Red";
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

                var getResultListGreenHDM = new List<string>();
                var getResultListGreenFocused = new List<string>();
                var getResultListGreenHDMQucik = new List<string>();

                var getResultListBlueLocate = new List<string>();

                var getResultListBlueRead = new List<string>();


                using (System.IO.StreamWriter resultFile = new System.IO.StreamWriter(@"..\..\..\..\..\TestResultCSV\" + csvFileName, true, System.Text.Encoding.GetEncoding("utf-8")))
                {
                    resultFile.WriteLine("Red Image, RedHDM, RedFSu, RedFUn, Green Image, GreenHDM, GreenFcs, GreenHDMQ, BlueLocate Image, BlueLocate, BlueRead Image, BlueRead ");
                    for (int indexcnt = 0; indexcnt < Constants.RepeatProcess; indexcnt++)
                    {
                        // Adding Green HDM Tool's getting process time.
                        // 0. Red Image, 1. RedHDM, 2. RedFSu, 3. RedFUn, 4. Green Image, 5. GreenHDM, 6. GreenFcs, 7. GreenHDMQ, 8. BlueLocate Image, 9. BlueLocate 10. BlueRead Image, 11. BlueRead
                        resultFile.WriteLine("{0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11}", myImagesFilesRedHDM.ElementAt(0), RedHDMTimeList[indexcnt].ToString(), RedFSuTimeList[indexcnt].ToString(), RedFUnTimeList[indexcnt].ToString(), myImagesFilesGreenHDM.ElementAt(0), GreenHDMTimeList[indexcnt].ToString(), GreenFocusedTimeList[indexcnt].ToString(), GreenHDMQTimeList[indexcnt].ToString(), myImagesFilesBlueLocate.ElementAt(0), BlueLocateTimeList[indexcnt].ToString(), myImagesFilesBlueRead.ElementAt(0), BlueReadTimeList[indexcnt].ToString());

                        getResultListRedHDM.Add(RedHDMTimeList[indexcnt].ToString());
                        getResultListRedFSu.Add(RedFSuTimeList[indexcnt].ToString());
                        getResultListRedFUn.Add(RedFUnTimeList[indexcnt].ToString());

                        getResultListGreenHDM.Add(GreenHDMTimeList[indexcnt].ToString());
                        getResultListGreenFocused.Add(GreenFocusedTimeList[indexcnt].ToString());
                        getResultListGreenHDMQucik.Add(GreenHDMQTimeList[indexcnt].ToString());

                        getResultListBlueLocate.Add(BlueLocateTimeList[indexcnt].ToString());

                        getResultListBlueRead.Add(BlueReadTimeList[indexcnt].ToString());


                    }
                }
                Console.WriteLine(" - Result CSV File: {0}", csvFileName);

                Console.WriteLine("\nStep 3. Save resultin Excel file");
                string getDateInfo = DateTime.Now.ToString("yyyy-MM-dd"); // refer to //www.delftstack.com/ko/howto/csharp/how-to-get-the-current-date-without-time-in-csharp/
                string strExcelFileName = "QAGetProcessingTime_" + getDateInfo + ".xlsx";
                string strExcelFileDirectory = Path.GetFullPath(@"..\..\..\..\..\TestResultCSV\") + strExcelFileName;   // Refer to - Processing file path name in using C# : //myoung-min.tistory.com/45
                Console.WriteLine(strExcelFileDirectory);

                //ExcelDataEPPlusRedTools(getResultListRedHDM, getResultListRedFSu, getResultListRedFUn, strExcelFileDirectory); // create epplus excel  - 20230504
                //ExcelDataEPPlusRedTools(getResultListRedHDM, getResultListRedFSu, getResultListRedFUn, getResultListGreenHDM, getResultListGreenFocused, getResultListGreenHDMQucik, getResultListBlueLocate, strExcelFileDirectory); // Adding Green HDM Tool in the create epplus excel  - 20230508
                ExcelDataEPPlusRedTools(getResultListRedHDM, getResultListRedFSu, getResultListRedFUn, getResultListGreenHDM, getResultListGreenFocused, getResultListGreenHDMQucik, getResultListBlueLocate, getResultListBlueRead, strExcelFileDirectory); // Adding Green HDM Tool in the create epplus excel  - 20230508
                Console.WriteLine("\nStep 4. Complete QA Test - Get Processing time of Red Tool");
            }
        }

        private static void ExcelDataEPPlusRedTools(List<string> GetPTimesRedHDM, List<string> GetPTimesRedFSu, List<string> GetPTimesRedFUn, List<string> GetPTimesGreenHDM, List<string> GetPTimesGreenFocused, List<string> GetPTimesGreenHDMQuick, List<string> GetPTimesBlueLocate, List<string> GetPTimesBlueRead, string savePath)
        {
            Console.WriteLine("JK Test 1. Create Excel File");
            ExcelPackage ExcelPkg = new ExcelPackage();

            // Red - Start
            ExcelWorksheet wsSheetRed = ExcelPkg.Workbook.Worksheets.Add("RedTools"); // Red Tools
            using (ExcelRange Rng = wsSheetRed.Cells[1, 1, 1, 1])  // 1x1
            {
                //Rng.Value = "QA JK's Task : Get processing time in teamcity!";
                Rng.Value = "Repeat";
                Rng.Style.Font.Size = 11; //16;
                Rng.Style.Font.Bold = true;
                Rng.Style.Font.Italic = true;
            }
            using (ExcelRange Rng = wsSheetRed.Cells[1, 2, 1, 2])
            {
                //Rng.Value = "QA JK's Task : Get processing time in teamcity!";
                Rng.Value = "Red HDM";
                Rng.Style.Font.Size = 11; //16;
                Rng.Style.Font.Bold = true;
                Rng.Style.Font.Italic = true;
            }
            using (ExcelRange Rng = wsSheetRed.Cells[1, 3, 1, 3])
            {
                //Rng.Value = "QA JK's Task : Get processing time in teamcity!";
                Rng.Value = "Red FSu";
                Rng.Style.Font.Size = 11; //16;
                Rng.Style.Font.Bold = true;
                Rng.Style.Font.Italic = true;
            }
            using (ExcelRange Rng = wsSheetRed.Cells[1, 4, 1, 4])
            {
                //Rng.Value = "QA JK's Task : Get processing time in teamcity!";
                Rng.Value = "Red FUn";
                Rng.Style.Font.Size = 11; //16;
                Rng.Style.Font.Bold = true;
                Rng.Style.Font.Italic = true;
            }
            wsSheetRed.Protection.IsProtected = false;
            wsSheetRed.Protection.AllowSelectLockedCells = false;

            // Red - End

            // Green - Start
            ExcelWorksheet wsSheetGreen = ExcelPkg.Workbook.Worksheets.Add("GreenTools"); // Green Tools

            using (ExcelRange Rng = wsSheetGreen.Cells[1, 1, 1, 1])  // 1x1
            {
                //Rng.Value = "QA JK's Task : Get processing time in teamcity!";
                Rng.Value = "Repeat";
                Rng.Style.Font.Size = 11; //16;
                Rng.Style.Font.Bold = true;
                Rng.Style.Font.Italic = true;
            }
            using (ExcelRange Rng = wsSheetGreen.Cells[1, 2, 1, 2])
            {
                //Rng.Value = "QA JK's Task : Get processing time in teamcity!";
                Rng.Value = "Green HDM";
                Rng.Style.Font.Size = 11; //16;
                Rng.Style.Font.Bold = true;
                Rng.Style.Font.Italic = true;
            }
            using (ExcelRange Rng = wsSheetGreen.Cells[1, 3, 1, 3])
            {
                //Rng.Value = "QA JK's Task : Get processing time in teamcity!";
                Rng.Value = "Green Fcs";
                Rng.Style.Font.Size = 11; //16;
                Rng.Style.Font.Bold = true;
                Rng.Style.Font.Italic = true;
            }
            using (ExcelRange Rng = wsSheetGreen.Cells[1, 4, 1, 4])
            {
                //Rng.Value = "QA JK's Task : Get processing time in teamcity!";
                Rng.Value = "Green HDMQ";
                Rng.Style.Font.Size = 11; //16;
                Rng.Style.Font.Bold = true;
                Rng.Style.Font.Italic = true;
            }
            wsSheetGreen.Protection.IsProtected = false;
            wsSheetGreen.Protection.AllowSelectLockedCells = false;
            // Green - End


            // Blue Locate - Start //GetPTimesBlueLocate
            ExcelWorksheet wsSheetBlueL = ExcelPkg.Workbook.Worksheets.Add("BlueLocateTool"); // Green Tools

            using (ExcelRange Rng = wsSheetBlueL.Cells[1, 1, 1, 1])  // 1x1
            {
                //Rng.Value = "QA JK's Task : Get processing time in teamcity!";
                Rng.Value = "Repeat";
                Rng.Style.Font.Size = 11; //16;
                Rng.Style.Font.Bold = true;
                Rng.Style.Font.Italic = true;
            }
            using (ExcelRange Rng = wsSheetBlueL.Cells[1, 2, 1, 2])
            {
                //Rng.Value = "QA JK's Task : Get processing time in teamcity!";
                Rng.Value = "Blue Locate";
                Rng.Style.Font.Size = 11; //16;
                Rng.Style.Font.Bold = true;
                Rng.Style.Font.Italic = true;
            }
            wsSheetBlueL.Protection.IsProtected = false;
            wsSheetBlueL.Protection.AllowSelectLockedCells = false;
            // Blue Locate - End

            // Blue Read - Start
            ExcelWorksheet wsSheetBlueR = ExcelPkg.Workbook.Worksheets.Add("BlueReadTool"); // Green Tools

            using (ExcelRange Rng = wsSheetBlueR.Cells[1, 1, 1, 1])  // 1x1
            {
                //Rng.Value = "QA JK's Task : Get processing time in teamcity!";
                Rng.Value = "Repeat";
                Rng.Style.Font.Size = 11; //16;
                Rng.Style.Font.Bold = true;
                Rng.Style.Font.Italic = true;
            }
            using (ExcelRange Rng = wsSheetBlueR.Cells[1, 2, 1, 2])
            {
                //Rng.Value = "QA JK's Task : Get processing time in teamcity!";
                Rng.Value = "Blue Read";
                Rng.Style.Font.Size = 11; //16;
                Rng.Style.Font.Bold = true;
                Rng.Style.Font.Italic = true;
            }
            wsSheetBlueR.Protection.IsProtected = false;
            wsSheetBlueR.Protection.AllowSelectLockedCells = false;
            // Blue Read - End

            ExcelPkg.SaveAs(new FileInfo(@savePath));
            Console.WriteLine(" - Complete the creating excel file!");

            Console.WriteLine("JK Test 2. Adding Chart after loading the created excel.");
            string pathExcelFile = savePath;
            Console.WriteLine(" - Load ExcelInfo: {0}", pathExcelFile);

            FileInfo existingFile = new FileInfo(pathExcelFile);
            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                // RedTools Chart - Start
                ExcelWorksheet worksheetRedTools = package.Workbook.Worksheets["RedTools"];
                int columnRedTools = 1;
                for (int row = 2; row < (Constants.RepeatProcess + 2); row++) // using repeat 100 times = 102
                    worksheetRedTools.Cells[row, columnRedTools].Value = row - 1;
                int colRedTools = 2;    // Red HDM
                for (int row = 2; row < (Constants.RepeatProcess + 2); row++) // using repeat 100 times = 102
                {
                    worksheetRedTools.Cells[row, colRedTools].Value = int.Parse(GetPTimesRedHDM[row - 2]);
                }
                colRedTools = 3;        // Red Focused Supervised
                for (int row = 2; row < (Constants.RepeatProcess + 2); row++) // using repeat 100 times = 102
                {
                    worksheetRedTools.Cells[row, colRedTools].Value = int.Parse(GetPTimesRedFSu[row - 2]);
                }
                colRedTools = 4;        // Red Focused Unsupervised
                for (int row = 2; row < (Constants.RepeatProcess + 2); row++)   // using repeat 100 times = 102
                {
                    worksheetRedTools.Cells[row, colRedTools].Value = int.Parse(GetPTimesRedFUn[row - 2]);
                }

                var chartRedTools = worksheetRedTools.Drawings.AddChart("Chart_Red", eChartType.Line);
                chartRedTools.Title.Text = "Processing Time Red Tool(HDM/FSu/FUn)[ms]"; ////From row 1 colum 5 with five pixels offset                
                //chartRedTools.Title.Font.Color = Color.FromArgb(238, 46, 34);
                chartRedTools.Title.Font.Size = 14;
                chartRedTools.Title.Font.Bold = true;
                chartRedTools.Title.Font.Italic = true;
                chartRedTools.SetPosition(1, 1, 6, 6); // Start point to dispale of Chart  ex) 0,0,5,5 : Draw a chart from F1 Cell vs 1,1,6,6 : Draw a chart from G2 Cell
                chartRedTools.SetSize(800, 600);

                ExcelAddress valueAddress_Data1_RedTools = new ExcelAddress(2, 2, (Constants.RepeatProcess + 1), 2); // using repeat 100 times
                ExcelAddress RepeatAddress_Data1_RedTools = new ExcelAddress(2, 1, (Constants.RepeatProcess + 1), 1);
                var ser1_RedTools = (chartRedTools.Series.Add(valueAddress_Data1_RedTools.Address, RepeatAddress_Data1_RedTools.Address) as ExcelLineChartSerie); // using repeat 100 time
                ser1_RedTools.Header = "Red HDM";

                ExcelAddress valueAddress_Data2_RedTools = new ExcelAddress(2, 3, (Constants.RepeatProcess + 1), 3); // using repeat 100 times
                ExcelAddress RepeatAddress_Data2_RedTools = new ExcelAddress(2, 1, (Constants.RepeatProcess + 1), 1);
                var ser2_RedTools = (chartRedTools.Series.Add(valueAddress_Data2_RedTools.Address, RepeatAddress_Data2_RedTools.Address) as ExcelLineChartSerie); // using repeat 100 time
                ser2_RedTools.Header = "Red FSu";

                ExcelAddress valueAddress_Data3_RedTools = new ExcelAddress(2, 4, (Constants.RepeatProcess + 1), 4); // using repeat 100 times
                ExcelAddress RepeatAddress_Data3_RedTools = new ExcelAddress(2, 1, (Constants.RepeatProcess + 1), 1);
                var ser3_RedTools = (chartRedTools.Series.Add(valueAddress_Data3_RedTools.Address, RepeatAddress_Data3_RedTools.Address) as ExcelLineChartSerie); // using repeat 100 times
                ser3_RedTools.Header = "Red FUn";

                chartRedTools.Legend.Border.LineStyle = eLineStyle.Solid;
                chartRedTools.Legend.Border.Fill.Style = eFillStyle.SolidFill;
                chartRedTools.Legend.Border.Fill.Color = Color.DarkRed;
                //chartRedTools.Border.Width = 1;
                chartRedTools.Border.Fill.Color = Color.DarkRed; // Color.FromArgb(238, 46, 34);


                // RedTools Chart - End

                // GreenTools Chart - Start
                ExcelWorksheet worksheetGreenTools = package.Workbook.Worksheets["GreenTools"];
                int columnGreenTools = 1;
                for (int row = 2; row < (Constants.RepeatProcess + 2); row++) // using repeat 100 times = 102
                    worksheetGreenTools.Cells[row, columnGreenTools].Value = row - 1;
                int colGreenTools = 2;    // Red HDM
                for (int row = 2; row < (Constants.RepeatProcess + 2); row++) // using repeat 100 times = 102
                {
                    worksheetGreenTools.Cells[row, colGreenTools].Value = int.Parse(GetPTimesGreenHDM[row - 2]);
                }
                colGreenTools = 3;        // Red Focused Supervised
                for (int row = 2; row < (Constants.RepeatProcess + 2); row++) // using repeat 100 times = 102
                {
                    worksheetGreenTools.Cells[row, colGreenTools].Value = int.Parse(GetPTimesGreenFocused[row - 2]);
                }
                colGreenTools = 4;        // Red Focused Unsupervised
                for (int row = 2; row < (Constants.RepeatProcess + 2); row++)   // using repeat 100 times = 102
                {
                    worksheetGreenTools.Cells[row, colGreenTools].Value = int.Parse(GetPTimesGreenHDMQuick[row - 2]);
                }

                var chartGreenTools = worksheetGreenTools.Drawings.AddChart("Chart_Green", eChartType.Line);
                chartGreenTools.Title.Text = "Processing Time Green Tool(HDM/Focused/HDMQuick)[ms]"; ////From row 1 colum 5 with five pixels offset                
                //chartGreenTools.Title.Font.Color = Color.FromArgb(16, 203, 34);
                chartGreenTools.Title.Font.Size = 14;
                chartGreenTools.Title.Font.Bold = true;
                chartGreenTools.Title.Font.Italic = true;
                chartGreenTools.SetPosition(1, 1, 6, 6); // Start point to dispale of Chart  ex) 0,0,5,5 : Draw a chart from F1 Cell vs 1,1,6,6 : Draw a chart from G2 Cell
                chartGreenTools.SetSize(800, 600);

                ExcelAddress valueAddress_Data1_GreenTools = new ExcelAddress(2, 2, (Constants.RepeatProcess + 1), 2); // using repeat 100 times
                ExcelAddress RepeatAddress_Data1_GreenTools = new ExcelAddress(2, 1, (Constants.RepeatProcess + 1), 1);
                var ser1_GreenTools = (chartGreenTools.Series.Add(valueAddress_Data1_GreenTools.Address, RepeatAddress_Data1_GreenTools.Address) as ExcelLineChartSerie); // using repeat 100 time
                ser1_GreenTools.Header = "Green HDM";

                ExcelAddress valueAddress_Data2_GreenTools = new ExcelAddress(2, 3, (Constants.RepeatProcess + 1), 3); // using repeat 100 times
                ExcelAddress RepeatAddress_Data2_GreenTools = new ExcelAddress(2, 1, (Constants.RepeatProcess + 1), 1);
                var ser2_GreenTools = (chartGreenTools.Series.Add(valueAddress_Data2_GreenTools.Address, RepeatAddress_Data2_GreenTools.Address) as ExcelLineChartSerie); // using repeat 100 time
                ser2_GreenTools.Header = "Green Focused";

                ExcelAddress valueAddress_Data3_GreenTools = new ExcelAddress(2, 4, (Constants.RepeatProcess + 1), 4); // using repeat 100 times
                ExcelAddress RepeatAddress_Data3_GreenTools = new ExcelAddress(2, 1, (Constants.RepeatProcess + 1), 1);
                var ser3_GreenTools = (chartGreenTools.Series.Add(valueAddress_Data3_GreenTools.Address, RepeatAddress_Data3_GreenTools.Address) as ExcelLineChartSerie); // using repeat 100 times
                ser3_GreenTools.Header = "Green HDMQuick";

                chartGreenTools.Legend.Border.LineStyle = eLineStyle.Solid;
                chartGreenTools.Legend.Border.Fill.Style = eFillStyle.SolidFill;
                chartGreenTools.Legend.Border.Fill.Color = Color.DarkGreen;
                //chartGreenTools.Border.Width = 1;
                chartGreenTools.Border.Fill.Color = Color.DarkGreen; //Color.FromArgb(16, 203, 34);
                // GreenTools Chart - End

                // BlueLocate Chart - Start
                ExcelWorksheet worksheetBlueLocateTool = package.Workbook.Worksheets["BlueLocateTool"];
                int columnBlueLocateTool = 1;
                for (int row = 2; row < (Constants.RepeatProcess + 2); row++) // using repeat 100 times = 102
                    worksheetBlueLocateTool.Cells[row, columnBlueLocateTool].Value = row - 1;
                int colBlueLocateTool = 2;    // Red HDM
                for (int row = 2; row < (Constants.RepeatProcess + 2); row++) // using repeat 100 times = 102
                {
                    worksheetBlueLocateTool.Cells[row, colBlueLocateTool].Value = int.Parse(GetPTimesBlueLocate[row - 2]);
                }

                var chartBlueLocateTool = worksheetBlueLocateTool.Drawings.AddChart("Chart_BlueLocate", eChartType.Line);
                chartBlueLocateTool.Title.Text = "Processing Time Blue Locate Tool[ms]"; //From row 1 colum 5 with five pixels offset
                //chartBlueLocateTool.Title.Font.Color = Color.FromArgb(0, 145, 255);
                chartBlueLocateTool.Title.Font.Size = 14;
                chartBlueLocateTool.Title.Font.Bold = true;
                chartBlueLocateTool.Title.Font.Italic = true;
                chartBlueLocateTool.SetPosition(1, 1, 6, 6); // Start point to dispale of Chart  ex) 0,0,5,5 : Draw a chart from F1 Cell vs 1,1,6,6 : Draw a chart from G2 Cell
                chartBlueLocateTool.SetSize(800, 600);

                ExcelAddress valueAddress_Data1_BlueLocateTool = new ExcelAddress(2, 2, (Constants.RepeatProcess + 1), 2); // using repeat 100 times
                ExcelAddress RepeatAddress_Data1_BlueLocateTool = new ExcelAddress(2, 1, (Constants.RepeatProcess + 1), 1);
                var ser1_BlueLocateTool = (chartBlueLocateTool.Series.Add(valueAddress_Data1_BlueLocateTool.Address, RepeatAddress_Data1_BlueLocateTool.Address) as ExcelLineChartSerie); // using repeat 100 time
                ser1_BlueLocateTool.Header = "Blue Locate";

                chartBlueLocateTool.Legend.Border.LineStyle = eLineStyle.Solid;
                chartBlueLocateTool.Legend.Border.Fill.Style = eFillStyle.SolidFill;
                chartBlueLocateTool.Legend.Border.Fill.Color = Color.DarkBlue;
                //chartBlueLocateTool.Border.Width = 1;
                chartBlueLocateTool.Border.Fill.Color = Color.DarkBlue;

                // BlueLocare Chart - End

                // BlueRead Chart - Start
                ExcelWorksheet worksheetBlueReadTool = package.Workbook.Worksheets["BlueReadTool"];
                int columnBlueRead = 1;
                for (int row = 2; row < (Constants.RepeatProcess + 2); row++) // using repeat 100 times = 102
                    worksheetBlueReadTool.Cells[row, columnBlueRead].Value = row - 1;
                int colBlueReadTool = 2;    // Red HDM
                for (int row = 2; row < (Constants.RepeatProcess + 2); row++) // using repeat 100 times = 102
                {
                    worksheetBlueReadTool.Cells[row, colBlueReadTool].Value = int.Parse(GetPTimesBlueRead[row - 2]);
                }

                var chartBlueReadTool = worksheetBlueReadTool.Drawings.AddChart("Chart_BlueRead", eChartType.Line);
                chartBlueReadTool.Title.Text = "Processing Time Blue Read Tool[ms]"; ////From row 1 colum 5 with five pixels offset                
                //chartBlueReadTool.Title.Font.Color = Color.FromArgb(0, 75, 163);
                chartBlueReadTool.Title.Font.Size = 14;
                chartBlueReadTool.Title.Font.Bold = true;
                chartBlueReadTool.Title.Font.Italic = true;
                chartBlueReadTool.SetPosition(1, 1, 6, 6); // Start point to dispale of Chart  ex) 0,0,5,5 : Draw a chart from F1 Cell vs 1,1,6,6 : Draw a chart from G2 Cell
                chartBlueReadTool.SetSize(800, 600);

                ExcelAddress valueAddress_Data1_BlueReadTool = new ExcelAddress(2, 2, (Constants.RepeatProcess + 1), 2); // using repeat 100 times
                ExcelAddress RepeatAddress_Data1_BlueReadTool = new ExcelAddress(2, 1, (Constants.RepeatProcess + 1), 1);
                var ser1_BlueReadTool = (chartBlueReadTool.Series.Add(valueAddress_Data1_BlueReadTool.Address, RepeatAddress_Data1_BlueReadTool.Address) as ExcelLineChartSerie); // using repeat 100 time
                ser1_BlueReadTool.Header = "Blue Read";

                chartBlueReadTool.Legend.Border.LineStyle = eLineStyle.Solid;
                chartBlueReadTool.Legend.Border.Fill.Style = eFillStyle.SolidFill;
                chartBlueReadTool.Legend.Border.Fill.Color = Color.DarkBlue;
                //chartBlueReadTool.Border.Width = 1;
                chartBlueReadTool.Border.Fill.Color = Color.DarkBlue;



                // BlueRead Chart - End


                package.Save();
            }
            Console.WriteLine("Complete - adding chart with using EPPlus.4.5.3.3");
            Console.WriteLine();
        }
    }
}

