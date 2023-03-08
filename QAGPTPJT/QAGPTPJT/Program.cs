using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ViDi2;
using ViDi2.Local;
using System.Timers;
using System.Diagnostics;
using System.Threading;


namespace QAGPTPJT
{
    class Program
    {
        static void Main(string[] args)
        {


            Console.WriteLine($"QA starts test - 20230306");

            // Initializes the control, This initialization does not allocate any gpu ressources.
            using (ViDi2.Runtime.Local.Control control = new ViDi2.Runtime.Local.Control(GpuMode.Deferred))
            {
                control.InitializeComputeDevices(GpuMode.SingleDevicePerTool, new List<int>() { }); // Initializes all CUDA devices                

                Console.WriteLine($"Step 1. Load RedHDM-Runtime worksapce & the directory of images.");

                ViDi2.Runtime.IWorkspace workspace = control.Workspaces.Add("workspace", "..\\..\\..\\..\\TestResource\\Runtime\\1_HDMR_S128x128.vrws");
                IStream stream = workspace.Streams["default"]; // Store a reference to the stream 'default'
                Stopwatch stopWatch = new Stopwatch();
                var ext = new System.Collections.Generic.List<string> { ".jpg", ".bmp", ".png" }; // Load an image from file                		               
                var myImagesFiles = Directory.GetFiles($"..\\..\\..\\..\\TestResource\\Images", "*.*", SearchOption.TopDirectoryOnly).Where(s => ext.Any(e => s.EndsWith(e)));// 기존 코드
                //Console.WriteLine(myImagesFiles.ElementAt(0));

                ITool redTool = stream.Tools["Analyze"];
                var hdParam = redTool.ParametersBase as ViDi2.Runtime.IToolParametersHighDetail;
                hdParam.ProcessTensorRT = true;
                // Process the image by the tool. All upstream tools are also processed                
                List<string> stimeList = new List<string>(); // JK start : to get the spending time of each image.                
                //Console.WriteLine($"img, processing time(ms)");

                Console.WriteLine($"Step 2. Start the getting processing time");

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
                Console.WriteLine($"\nAverage: {avg} [msec]");

                Console.WriteLine($"Step 3. Finish the getting processing time");

                string csvFileName = "GetProcessingTime_ImageSize_128_test20230220-1.csv";
                int indexcnt = 0;
                using (System.IO.StreamWriter resultFile = new System.IO.StreamWriter(@"..\..\..\..\" + csvFileName, false, System.Text.Encoding.GetEncoding("utf-8")))
                //using (System.IO.StreamWriter resultFile = new System.IO.StreamWriter(@"H:\_JK_Task_2023Q1\TestCode_GetProcessingTime\" + csvFileName, false, System.Text.Encoding.GetEncoding("utf-8")))
                {
                    resultFile.WriteLine("ImagePath, SpendingTime"); // 각 필드에 사용될 제목 정의   Refer to : bjy2.tistory.com/199
                                                                     // 
                    foreach (var resultTime in stimeList) // 필드에 값을 채워줌
                    {
                        resultFile.WriteLine("{0}, {1}", myImagesFiles.ElementAt(indexcnt), resultTime.ToString()); //resultFile.WriteLine("{0}, {1}", myImagesFiles, resultTime.ToString());
                        indexcnt = indexcnt + 1;                         //file.WriteLine("{0},{1}", el.name, el.age);
                    }
                }
                Console.WriteLine($"Step 4. Complete Saving result in cvs file");
            }
        }
    }
}
