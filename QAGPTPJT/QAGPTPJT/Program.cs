﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ViDi2;
using ViDi2.Local;
using System.Timers;
using System.Diagnostics;
using System.Threading;

using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;


namespace QAGPTPJT
{
    class Program
    {
        static void Main(string[] args)
        {
            //string dataInfo = DateTime.Now.ToString("yyyy-MM-dd"); // JK start a test to get the processing time of all tool on VPDL.
            Console.WriteLine($"*** QA-Get Processing Time:" + DateTime.Now.ToString("yyyy-MM-dd") + " ***\n");

            // Initializes the control, This initialization does not allocate any gpu ressources.
            using (ViDi2.Runtime.Local.Control control = new ViDi2.Runtime.Local.Control(GpuMode.Deferred))
            {
                Console.WriteLine($"\n01. Preparation of configuration - Initializes all CUDA devices.");
                control.InitializeComputeDevices(GpuMode.SingleDevicePerTool, new List<int>() { }); // Initializes all CUDA devices
                /* Getting configuration in system e.g., GPU model, Driver Version, OS etc - It's next task*/
                Console.WriteLine("[Configuration of the current agent in teamcity]");
                Console.WriteLine(" - VPDL Ver.: " + control.CLibraryVersion);
                Console.WriteLine(" - GPU Model: {0}", control.ComputeDevices[0].Name);// Index: control.ComputeDevices[0].Index.ToString()

                // Step 1. Load RedHDM-Runtime worksapce & the directory of images. /////////////////////////////////////////////////////////////////////////////////////////////////////////

                Console.WriteLine($"\nStep 1. Load RedHDM-Runtime worksapce & the directory of images.");
                // initialization for using Stopwatch and saving result of processing time.
                Stopwatch stopWatch = new Stopwatch();
                List<string> stimeList = new List<string>(); // JK start : to get the spending time of each image.                
                // Process the image by the tool. All upstream tools are also processed // Console.WriteLine($"img, processing time(ms)");

                // Set Runtime of Red Tool 
                string pathRuntime_Red_HDM = "..\\..\\..\\..\\..\\TestResource\\Runtime\\1_RED_HighDetailMode.vrws";
                Console.WriteLine(" - Runtime Path: {0}", pathRuntime_Red_HDM);// Index: control.ComputeDevices[0].Index.ToString()
                ViDi2.Runtime.IWorkspace workspace = control.Workspaces.Add("workspace", pathRuntime_Red_HDM);

                //string pathRuntime_Red_FSu = "..\\..\\..\\..\\..\\TestResource\\Runtime\\2_RED_FocusedSupervised.vrws";
                //Console.WriteLine(" - Runtime Path: {0}", pathRuntime_Red_FSu);// Index: control.ComputeDevices[0].Index.ToString()
                //ViDi2.Runtime.IWorkspace workspace = control.Workspaces.Add("workspace", pathRuntime_Red_FSu);

                //string pathRuntime_Red_FUn = "..\\..\\..\\..\\..\\TestResource\\Runtime\\3_RED_FocusedUnsupervised.vrws";
                //Console.WriteLine(" - Runtime Path: {0}", pathRuntime_Red_FUn);// Index: control.ComputeDevices[0].Index.ToString()
                //ViDi2.Runtime.IWorkspace workspace = control.Workspaces.Add("workspace", pathRuntime_Red_FUn);



                IStream stream = workspace.Streams["default"]; // Store a reference to the stream 'default'
                ITool redTool = stream.Tools["Analyze"];

                //var hdParam = redTool.ParametersBase as ViDi2.Runtime.IRedTool; // in case of usnign Red Focused Tool's runtime

                var hdParam = redTool.ParametersBase as ViDi2.Runtime.IToolParametersHighDetail; // in case of usnign Red HDM Tool's runtime
                //hdParam.ProcessTensorRT = true; // To use, Need to do prework which have done build by Example.Runtime.OptimizeHDTool.console - 20230419
                //hdParam.ProcessTensorRT = false; // This is case that runtime did not apply Optimized runtime.
                // This runtime workspace didn't apply the optimized runtime by tenserRT. So You can not use the setting dParam.ProcessTensorRT whether it's true or false as below that. In conclusion, this setting have to skip as below when execution file was run after building in teamcity.

                // Set Data of Red Tool
                string pathRedImages = "..\\..\\..\\..\\..\\TestResource\\Images";
                var ext = new System.Collections.Generic.List<string> { ".jpg", ".bmp", ".png" }; // Load an image from file                		                               
                var myImagesFiles = Directory.GetFiles(pathRedImages, "*.*", SearchOption.TopDirectoryOnly).Where(s => ext.Any(e => s.EndsWith(e)));
                // Check a status whether load first image or not
                Console.WriteLine("First Image info. : " + myImagesFiles.ElementAt(0));

                Console.WriteLine($"\nStep 2. Start the getting processing time");

                long sum = 0;
                int count = 0;
                foreach (var file in myImagesFiles)
                {
                    count++;
                    using (IImage image = new LibraryImage(file))
                    {
                        using (ISample sample = stream.CreateSample(image))
                        {
                            stopWatch.Start();
                            sample.Process(redTool);
                            stopWatch.Stop();
                            //Console.WriteLine($"{file},{stopWatch.ElapsedMilliseconds}, [msec]");
                            sum += stopWatch.ElapsedMilliseconds;
                            stimeList.Add(stopWatch.ElapsedMilliseconds.ToString()); // JK start : to get the spending time of each image.                            
                            stopWatch.Reset();
                        }
                    }
                }
                double avg = sum / (double)count;
                Console.WriteLine(" - Processing Time Average({0} images): {1} [msec]", (int)count, avg);

                // Step 3. Finish the getting processing time ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                Console.WriteLine($"\nStep 3. Finish the getting processing time");
                string strDateGetResult = DateTime.Now.ToString("yyyy-MM-dd");
                string csvFileName = "GetProcessingTime_" + strDateGetResult + ".csv";

                // JK test to apply Excel - Start.               
                List<ImageProcessingTime> getResultList = new List<ImageProcessingTime>(); //var getResultList = new List<ImageProcessingTime>();
                // JK test to apply Excel - End.

                int indexcnt = 0;
                using (System.IO.StreamWriter resultFile = new System.IO.StreamWriter(@"..\..\..\..\..\TestResultCSV\" + csvFileName, true, System.Text.Encoding.GetEncoding("utf-8")))
                //using (System.IO.StreamWriter resultFile = new System.IO.StreamWriter(@"..\..\..\..\..\TestResultCSV\" + csvFileName, false, System.Text.Encoding.GetEncoding("utf-8"))) // false : overwrite
                {
                    resultFile.WriteLine("ImagePath, SpendingTime");    // 각 필드에 사용될 제목 정의   Refer to : bjy2.tistory.com/199
                    foreach (var resultTime in stimeList)               // Fill in value(processing time) in Cell(field)
                    {
                        resultFile.WriteLine("{0}, {1}", myImagesFiles.ElementAt(indexcnt), resultTime.ToString()); //resultFile.WriteLine("{0}, {1}", myImagesFiles, resultTime.ToString());
                        // JK test to apply Excel - Start.
                        getResultList.Add(new ImageProcessingTime() { ImagePath = myImagesFiles.ElementAt(indexcnt), ProcessingTime = resultTime }); // refer to learn.microsoft.com/ko-kr/dotnet/api/system.collections.generic.list-1.add?view=net-7.0
                        // JK test to apply Excel - End                        
                        indexcnt = indexcnt + 1;    //file.WriteLine("{0},{1}", el.name, el.age);                       


                    }
                }
                Console.WriteLine(" - Result CSV File: {0}", csvFileName);

                Console.WriteLine("\nStep 4. Complete QA Test");

                // hostramus.tistory.com/94 --> C# Excel how to use that
                // Microsoft.Office.Interop.Excel;

                // Create a list of accounts.
                //var bankAccounts = new List<Account>
                //{
                //    new Account
                //    {
                //        ID = 345678,
                //        Balance = 541.27
                //    },
                //    new Account
                //    {
                //        ID = 1230221,
                //        Balance = -127.44
                //    }
                //};           

                //// checking how to use
                //var getResultList = new List<ImageProcessingTime>();
                //ImageProcessingTime test = new ImageProcessingTime();
                //test.ImagePath = "J.png";
                //test.ProcessingTime = "123.1";
                //getResultList.Add(test);
                //var getResultList = new List<ImageProcessingTime>
                //{
                //    new ImageProcessingTime
                //    {
                //        ImagePath = "345678",
                //        ProcessingTime = "541.27"
                //    },
                //    new ImageProcessingTime
                //    {
                //        ImagePath = "1230221",
                //        ProcessingTime = "-127.44"
                //    }
                //};


                // Display the list in an Excel spreadsheet.
                //DisplayInExcel(bankAccounts);

                // refer to get tiem : //developer-talk.tistory.com/147
                string getDateInfo = DateTime.Now.ToString("yyyy-MM-dd"); // refer to //www.delftstack.com/ko/howto/csharp/how-to-get-the-current-date-without-time-in-csharp/
                string strExcelFileName = "QAGetProcessingTime_" + getDateInfo + ".xlsx";
                string strExcelFileDirectory = Path.GetFullPath(@"..\..\..\..\..\TestResultCSV\") + strExcelFileName;   // Refer to - Processing file path name in using C# : //myoung-min.tistory.com/45
                Console.WriteLine(strExcelFileDirectory);
                //DisplayInExcel(getResultList, strExcelFileDirectory);
                JKWriteExcelData(getResultList, strExcelFileDirectory);



                //var fullPath = Path.GetFullPath(@"..\..\..\..\..\TestResultCSV\JK_Test.xlsx");
                //Console.WriteLine(fullPath);
                //DisplayInExcel(getResultList, fullPath);


            }
        }


        private static void JKWriteExcelData(IEnumerable<ImageProcessingTime> accounts, string savePath)
        {
            // 최초 엑셀 파일을 만들고 내용을 기입하는 경우. 경로에 아무런 엑셀 파일이 없는 경우 의미함	
            Excel.Application excelApp = null;
            Excel.Workbook wb = null;
            Excel.Worksheet ws = null;

            excelApp = new Excel.Application();

            excelApp.DisplayAlerts = false; // refer to : //musma.github.io/2019/04/01/dotnet-undead-excel-process.html
            // refer to : //learn.microsoft.com/ko-kr/dotnet/api/microsoft.office.interop.excel._application.displayalerts?view=excel-pia
            // 매크로가 실행되는 동안 MS Excel 에서 특정 경고 및 메시지를 표시하는 경우 True임 --> When i was done the overwriting by SaveAs(), 
            // 통합 문서에서 SaveAs(Object, Object, Object, Object, Object, Object, XlSaveAsAccessMode, Object, Object, Object, Object, Object) 메서드를 사용하여 기존 파일을 덮어쓸 때 '덮어쓰기' 경고의 기본값은 '아니요'입니다.
            // DisplayAlerts 속성이 False 로 설정된 경우 Excel에서 '예' 응답을 선택합니다 .

            wb = excelApp.Workbooks.Add();    // Excel.Workbook wb = excelApp.Workbooks.Open(ExcelPath); // ex) Open(@"D:\test\test.xlsx");
            ws = wb.Worksheets.get_Item(1) as Excel.Worksheet;  // Case 1. Select first worksheet
            ws = wb.Worksheets.Item["Sheet1"];  // Case 2. Select first worksheet - //musma.github.io/2019/04/01/dotnet-undead-excel-process.html

            ws.Name = "RedTool"; // workbooks.Open()인 경우에만 ws.Save()  되는 것 같음....            

            //ws.Cells[1, "A"] = "Input image directory";
            //ws.Cells[1, "B"] = "Processing Time [msec]";


            ws.Cells[1, "A"] = " *** ";
            ws.Cells[2, "A"] = "Input image directory";
            ws.Cells[1, "B"] = "Processing Time [msec]";
            ws.Cells[2, "B"] = "Red_HDM";
            ws.Cells[2, "C"] = "Red_FSu";

            //var row = 1;
            var row = 2;
            foreach (var acct in accounts)
            {
                row++;
                ws.Cells[row, "A"] = acct.ImagePath;
                ws.Cells[row, "B"] = acct.ProcessingTime; // Red HDM
            }
            string strLastIndex = row.ToString();
            //string strLastLine = "B" + strLastIndex;
            string strLastLine = "C" + strLastIndex;
            ws.Columns[1].AutoFit();
            ws.Columns[2].AutoFit();

            ws.Columns[3].AutoFit();

            ws.Range["A1", strLastLine].AutoFormat(Excel.XlRangeAutoFormat.xlRangeAutoFormatClassic2); // xlRangeAutoFormatClassic2 (구분셀 색상) vs xlRangeAutoFormatClassic1 (기본)

            // C# adding chart in excal : //csharp.net-informations.com/excel/csharp-excel-chart.htm
            object misValue = System.Reflection.Missing.Value;
            Excel.Range chartRange;

            Excel.ChartObjects xlChart = (Excel.ChartObjects)ws.ChartObjects(Type.Missing);
            Excel.ChartObject myChart = (Excel.ChartObject)xlChart.Add(10, 80, 300, 250);
            Excel.Chart chartPage = myChart.Chart;

            //chartRange = ws.get_Range("B1", "B5");
            chartRange = ws.get_Range("B1", strLastLine);
            chartPage.SetSourceData(chartRange, misValue);
            //chartPage.ChartType = Excel.XlChartType.xlColumnClustered; // Value 51 : //learn.microsoft.com/en-us/office/vba/api/excel.xlcharttype
            chartPage.ChartType = Excel.XlChartType.xlLine; // Value 4




            //string fullRange = "A1:" + strLastLine;
            //ws.Range[fullRange].Copy();            

            wb.SaveAs(@savePath);//, ReadOnlyRecommended:false); // checked saving file in this directory
            wb.Close();
            excelApp.Quit();

            ReleaseObject(ws);
            ReleaseObject(wb);
            ReleaseObject(excelApp);
        }

        /// <summary>
        /// 액셀 객체 해제 메소드
        /// </summary>
        /// <param name="obj"></param>
        static void ReleaseObject(object obj)
        {
            try
            {
                if (obj != null)
                {
                    Marshal.ReleaseComObject(obj);  // 액셀 객체 해제
                    obj = null;
                }
            }
            catch (System.Exception ex)
            {
                obj = null;
                throw ex;
            }
            finally
            {
                GC.Collect();   // 가비지 수집
            }
        }


        // 참고 1. //blog.naver.com/PostView.nhn?blogId=freedman80&logNo=221533629306
        private static void WriteExcelData() // 참고 자료 출처 : //gigong.tistory.com/96
        {
            //string ExcelPath = @"..\..\..\..\..\TestResultCSV\JK_Test.xlsx";

            Excel.Application excelApp = null;
            Excel.Workbook wb = null;
            Excel.Worksheet ws = null;

            excelApp = new Excel.Application();

            //wb = excelApp.Workbooks.Open(ExcelPath);
            // 엑셀파일을 엽니다.
            // ExcelPath 대신 문자열도 가능합니다
            // 예. Open(@"D:\test\test.xlsx");

            ws = wb.Worksheets.get_Item(1) as Excel.Worksheet;
            // 첫번째 Worksheet를 선택합니다.

            ws.Cells[1, "A"] = "Input image directory"; //"ID Number";

            // 따로 저장하지 않는다면 지금 파일에 그대로 저장합니다.
            wb.Save();

            wb.Close();
            excelApp.Quit();

        }





        static void DisplayInExcel(IEnumerable<ImageProcessingTime> accounts, string savePath)
        {
            // 참고 코드 출처 - Office interop 객체에 액서스하는 방법 : //learn.microsoft.com/ko-kr/dotnet/csharp/advanced-topics/interop/how-to-access-office-interop-objects
            // 참고로 이 내용은 엑셀 화면 창을 띄고 할 수 이는 듯함.

            var excelApp = new Excel.Application();
            // Make the object visible.
            //excelApp.Visible = true;
            excelApp.Visible = false;

            // Create a new, empty workbook and add it to the collection returned
            // by property Workbooks. The new workbook becomes the active workbook.
            // Add has an optional parameter for specifying a particular template.
            // Because no argument is sent in this example, Add creates a new workbook.
            excelApp.Workbooks.Add();

            // This example uses a single workSheet.
            Excel._Worksheet workSheet = excelApp.ActiveSheet;

            // Earlier versions of C# require explicit casting.
            //Excel._Worksheet workSheet = (Excel.Worksheet)excelApp.ActiveSheet;

            // Establish column headings in cells A1 and B1.
            workSheet.Cells[1, "A"] = "Input image directory"; //"ID Number";
            workSheet.Cells[1, "B"] = "Processing Time [msec]";//"Current Balance";

            var row = 1;
            foreach (var acct in accounts)
            {
                row++;
                workSheet.Cells[row, "A"] = acct.ImagePath;
                workSheet.Cells[row, "B"] = acct.ProcessingTime;
            }
            string strLastIndex = row.ToString();

            string strLastLine = "B" + strLastIndex;

            workSheet.Columns[1].AutoFit();
            workSheet.Columns[2].AutoFit();

            // Call to AutoFormat in Visual C#. This statement replaces the
            // two calls to AutoFit.
            //workSheet.Range["A1", "B3"].AutoFormat(Excel.XlRangeAutoFormat.xlRangeAutoFormatClassic2);
            //workSheet.Range["A1", "B5"].AutoFormat(Excel.XlRangeAutoFormat.xlRangeAutoFormatClassic2); //xlRangeAutoFormatClassic1

            //workSheet.Range["A1", "B5"].AutoFormat(Excel.XlRangeAutoFormat.xlRangeAutoFormatClassic1); //xlRangeAutoFormatClassic1
            workSheet.Range["A1", strLastLine].AutoFormat(Excel.XlRangeAutoFormat.xlRangeAutoFormatClassic1); //xlRangeAutoFormatClassic1

            // Put the spreadsheet contents on the clipboard. The Copy method has one
            // optional parameter for specifying a destination. Because no argument
            // is sent, the destination is the Clipboard.
            //workSheet.Range["A1:B3"].Copy();
            //workSheet.Range["A1:B5"].Copy();

            string fullRange = "A1:" + strLastLine;
            workSheet.Range[fullRange].Copy();

            // refer to learn.microsoft.com/ko-kr/visualstudio/vsto/how-to-programmatically-save-workbooks?source=recommendations&view=vs-2022&tabs=csharp
            //workSheet.SaveAs(@"H:\123.xlsx"); // checked saving file in this directory
            //



            workSheet.SaveAs(@savePath, ReadOnlyRecommended: false); // checked saving file in this directory



            //excelApp.Quit();




            // 위와 같이 하면 저장하면서 기존파일 덮어쓰기 여부 문의하고 추가여 열면 읽기모드 할것 인지 문의함.
            //workSheet.SaveAs(@"..\..\..\..\..\TestResultCSV\123.xlsx"); //..\..\..\..\..\TestResultCSV\ --> 실패


        }

        //static void CreateIconInWordDoc()
        //{
        //    var wordApp = new Word.Application();
        //    wordApp.Visible = true;

        //    // The Add method has four reference parameters, all of which are
        //    // optional. Visual C# allows you to omit arguments for them if
        //    // the default values are what you want.
        //    wordApp.Documents.Add();

        //    // PasteSpecial has seven reference parameters, all of which are
        //    // optional. This example uses named arguments to specify values
        //    // for two of the parameters. Although these are reference
        //    // parameters, you do not need to use the ref keyword, or to create
        //    // variables to send in as arguments. You can send the values directly.
        //    wordApp.Selection.PasteSpecial(Link: true, DisplayAsIcon: true);
        //}
    }

    public class ImageProcessingTime //Account
    {

        public string ImagePath { get; set; }
        public string ProcessingTime { get; set; }

        //public int ID { get; set; }
        //public double Balance { get; set; }
    }

}



// Middle Report Code - Before 20230421, Teamcity agent #5 & #7 - Build No.123
//using System;
//using System.Collections.Generic;
//using System.IO;
//using System.Linq;
//using ViDi2;
//using ViDi2.Local;
//using System.Timers;
//using System.Diagnostics;
//using System.Threading;


//namespace QAGPTPJT
//{
//    class Program
//    {
//        static void Main(string[] args)
//        {
//            Console.WriteLine($"QA starts test - 20230419"); // Modify test code version from Console.WriteLine($"QA starts test - 20230306");

//            // Initializes the control, This initialization does not allocate any gpu ressources.
//            using (ViDi2.Runtime.Local.Control control = new ViDi2.Runtime.Local.Control(GpuMode.Deferred))
//            {
//                Console.WriteLine($"01. Preparation of configuration - Initializes all CUDA devices.");
//                control.InitializeComputeDevices(GpuMode.SingleDevicePerTool, new List<int>() { }); // Initializes all CUDA devices
//                /* Getting configuration in system e.g., GPU model, Driver Version, OS etc - It's next task*/

//                Console.WriteLine($"Step 1. Load RedHDM-Runtime worksapce & the directory of images.");              

//                // JK 20230419 - Start : Case 1. This under code line use when Red-HDM Tool applied in runtime // TestResource's path : QAGPT_Build_22_artifacts\target_directory\TestResource
//                ViDi2.Runtime.IWorkspace workspace = control.Workspaces.Add("workspace", "..\\..\\..\\..\\..\\TestResource\\Runtime\\2_REDHDM_S128x128.vrws"); // x64\release 사용으로 ..\ 추가됨.                
//                // JK 20230419 - End

//                // JK 20230419 - Start : Case 2. This under code line use when Red-Focused Supervised Tool applied in runtime
//                //ViDi2.Runtime.IWorkspace workspace = control.Workspaces.Add("workspace", "..\\..\\..\\..\\..\\TestResource\\Runtime\\1_REDFSUPER_S128x128.vrws"); // x64\release 사용으로 ..\ 추가됨.                
//                // JK 20230419 - End

//                IStream stream = workspace.Streams["default"]; // Store a reference to the stream 'default'
//                Stopwatch stopWatch = new Stopwatch();
//                var ext = new System.Collections.Generic.List<string> { ".jpg", ".bmp", ".png" }; // Load an image from file                		                               
//                var myImagesFiles = Directory.GetFiles($"..\\..\\..\\..\\..\\TestResource\\Images", "*.*", SearchOption.TopDirectoryOnly).Where(s => ext.Any(e => s.EndsWith(e)));
//                //var myImagesFiles = Directory.GetFiles($".\\TestResource\\Images", "*.*", SearchOption.TopDirectoryOnly).Where(s => ext.Any(e => s.EndsWith(e)));// 기존 코드
//                Console.WriteLine("First Image info. : " + myImagesFiles.ElementAt(0));

//                ITool redTool = stream.Tools["Analyze"];


//                // JK 20230419 - Start : Case 1. This under code line use when Red-HDM Tool applied in runtime
//                var hdParam = redTool.ParametersBase as ViDi2.Runtime.IToolParametersHighDetail;
//                // JK Notity 20230419
//                // This runtime workspace didn't apply the optimized runtime by tenserRT.
//                //So You can not use the setting dParam.ProcessTensorRT whether it's true or false as below that.
//                //In conclusion, this setting have to skip as below when execution file was run after building in teamcity.
//                //hdParam.ProcessTensorRT = true; // To use, Need to do prework which have done build by Example.Runtime.OptimizeHDTool.console - 20230419
//                //hdParam.ProcessTensorRT = false; // This is case that runtime did not apply Optimized runtime.
//                // JK 20230419 - End

//                // JK 20230419 - Start : Case 2. This under code line use when Red-Focused Supervised Tool applied in runtime
//                //var hdParam = redTool.ParametersBase as ViDi2.Runtime.IRedTool; // JK modify 20230418 becuase of using red focused super
//                // JK 20230419 - End

//                // Process the image by the tool. All upstream tools are also processed                
//                List<string> stimeList = new List<string>(); // JK start : to get the spending time of each image.                
//                //Console.WriteLine($"img, processing time(ms)");

//                Console.WriteLine($"Step 2. Start the getting processing time");

//                long sum = 0;
//                int count = 0;
//                foreach (var file in myImagesFiles)
//                {
//                    count++;
//                    using (IImage image = new LibraryImage(file))
//                    {
//                        using (ISample sample = stream.CreateSample(image))
//                        {
//                            stopWatch.Start();
//                            sample.Process(redTool);
//                            stopWatch.Stop();
//                            //Console.WriteLine($"{file},{stopWatch.ElapsedMilliseconds}, [msec]");
//                            sum += stopWatch.ElapsedMilliseconds;
//                            stimeList.Add(stopWatch.ElapsedMilliseconds.ToString()); // JK start : to get the spending time of each image.                            
//                            stopWatch.Reset();
//                        }
//                    }
//                }
//                double avg = sum / (double)count;
//                Console.WriteLine($"\nAverage: {avg} [msec]");

//                Console.WriteLine($"Step 3. Finish the getting processing time");

//                string csvFileName = "GetProcessingTime_ImageSize_128_test20230418-1.csv";
//                int indexcnt = 0;
//                //using (System.IO.StreamWriter resultFile = new System.IO.StreamWriter(@"..\..\..\..\" + csvFileName, false, System.Text.Encoding.GetEncoding("utf-8")))
//                using (System.IO.StreamWriter resultFile = new System.IO.StreamWriter(@"..\..\..\..\..\TestResultCSV\" + csvFileName, false, System.Text.Encoding.GetEncoding("utf-8")))
//                // 참고용 경로 : H:\20230410_\QAGPT_Build_22_artifacts\target_directory\TestResource\Result\GetProcessingTime_ImageSize_128_test20230220-1.csv

//                //H:\20230410_\QAGPT_Build_22_artifacts\target_directory\QAGPTPJT\QAGPTPJT\TestResource\Result\GetProcessingTime_ImageSize_128_test20230220 - 1.csv

//                //using (System.IO.StreamWriter resultFile = new System.IO.StreamWriter(@"H:\_JK_Task_2023Q1\TestCode_GetProcessingTime\" + csvFileName, false, System.Text.Encoding.GetEncoding("utf-8")))
//                {
//                    resultFile.WriteLine("ImagePath, SpendingTime"); // 각 필드에 사용될 제목 정의   Refer to : bjy2.tistory.com/199
//                                                                     // 
//                    foreach (var resultTime in stimeList) // 필드에 값을 채워줌
//                    {
//                        resultFile.WriteLine("{0}, {1}", myImagesFiles.ElementAt(indexcnt), resultTime.ToString()); //resultFile.WriteLine("{0}, {1}", myImagesFiles, resultTime.ToString());
//                        indexcnt = indexcnt + 1;                         //file.WriteLine("{0},{1}", el.name, el.age);
//                    }
//                }
//                Console.WriteLine($"Step 4. Complete Saving result in cvs file");
//            }
//        }
//    }
//}
