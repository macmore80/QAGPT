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
using OfficeOpenXml.Style;
using ViDi2;
using ViDi2.Local;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;


// JK - Modification log
// Note : JK modify code 2023.05.04 : processing Red Tool(HDM, FSu, FUn)  and Saving result and Chart in Excel
// Note : JK modify code 2023.05.08 : Adding Green Tools(HDM, Focused, HDM Quick), This data save in Excel file as above.
// Note : JK Modify code 2023.05.09 : To analsy result, Add data which are max, min, average in excel file. And include test system's configuration
// Note : JK have checked the setting build to add result data about all tools after gettting the processing tiems - 20230704
// Note : // JK-AddResult-Start - 2023.07.06 : JK Modify code 2023.07.06 : Adding results which got from tool after processing tool.
// Note : //JK-AddResult-End - 2023.07.11 : Regarding Blue Read, Adding code
// Note : // JK - 2023.07.12 - End : Adding code which add results of Green and Red, Fix font, color, etc in Excel file.
// Note : // JK-AddResultOFGreen-2023.07.12- Start & // JK-AddResultOFRed-2023.07.12- Start
// Note : // JK-ModifyCodeRedHDM-2023.07.13- Start
// Note : // JK-Modified-2023.07.17 - Start  : Modified contents in TestConfiguration and the saved file name which have to include VPDL version number.


namespace QAGPTPJT
{
    static class Constants
    {
        public const int RepeatProcess = 5; // When initialised variabl, if It include 'const' in type string, This value never change.        
    }

    public static class TestConfigurationItems
    {
        public static string VPDLVers = "0.0.0.00000";
    }

    public class JKCtrlDirectory
    {
        public static void CreateDir(string path)
        {
            string currentPath = Environment.CurrentDirectory;
            try
            {
                // Determine whether the directory exists.
                if (Directory.Exists(path))
                {
                    Console.WriteLine("That path exists already.");
                    Console.WriteLine(" - The existed directory : \n\t" + currentPath + "\\" + path);
                    return;
                }
                // Try to create the directory.
                DirectoryInfo di = Directory.CreateDirectory(path);
                Console.WriteLine("The directory was created successfully at {0}.", Directory.GetCreationTime(path));
                Console.WriteLine(" - the created directory : \n\t" + currentPath + "\\" + path);

                // Delete the directory.
                //di.Delete();
                //Console.WriteLine("The directory was deleted successfully.");
            }
            catch (System.Exception e)
            {
                Console.WriteLine("The process failed: {0}", e.ToString());
            }
            finally { }
        }
    }

    class Program
    {
        // JK Note - start : 2023.07.11 - Adding a structure of result of matching blue read model
        // refer to  : tnmsoft.tistory.com/304
        // refer to : nonstop-antoine.tistory.com/47#google_vignette
        public struct BlueReadMatchFeatureResult
        {
            public string Name;
            public double Score;
            public double PosX;
            public double PosY;
            public double Angle;
            public double SizeHeight;
            public double SizeWidth;
            public BlueReadMatchFeatureResult(string name, double score, double posx, double posy, double angle, double sizeheight, double sizewidth)
            {
                Name = name;
                Score = score;
                PosX = posx;
                PosY = posy;
                Angle = angle;
                SizeHeight = sizeheight;
                SizeWidth = sizewidth;
            }
        }
        // JK Note - end : 2023.07.11 - Adding a structure of result of matching blue read model

        // JK-AddResultOFGreen-2023.07.12- Start
        public struct GreenHDMMatchAndViewResult
        {
            // view
            public string BestTagName;
            public double BestTagScore;
            public double Threshold;
            public double SizeHeight;
            public double SizeWidth;

            public GreenHDMMatchAndViewResult(string besttagname, double besttagscore, double threshold, double sizeheight, double sizewidth)
            {
                BestTagName = besttagname;
                BestTagScore = besttagscore;
                Threshold = threshold;
                SizeHeight = sizeheight;
                SizeWidth = sizewidth;
            }


            //// JK Memo : If you want to gather all tag names, refer to 'view.Tags', This array type variable has four case of Name and score. For example, A, B, C, D. then modify the code below.
            //// Match
            //public string Name;
            //public double Score;
            //// view
            //public string BestTagName;
            //public double BestTagScore;            
            //public double Threshold;
            //public double SizeHeight;
            //public double SizeWidth;

            //public GreenHDMMatchAndViewResult(string name, double score, string besttagname, double besttagscore, double threshold, double sizeheight, double sizewidth)
            //{
            //    Name = name;
            //    Score = score;
            //    BestTagName = besttagname;
            //    BestTagScore = besttagscore;
            //    Threshold = threshold;
            //    SizeHeight = sizeheight;
            //    SizeWidth = sizewidth;
            //}
        }

        public struct GreenFocusedMatchAndViewResult
        {
            public string BestTagName;
            public double BestTagScore;
            public double Threshold;
            public double SizeHeight;
            public double SizeWidth;

            public GreenFocusedMatchAndViewResult(string besttagname, double besttagscore, double threshold, double sizeheight, double sizewidth)
            {
                BestTagName = besttagname;
                BestTagScore = besttagscore;
                Threshold = threshold;
                SizeHeight = sizeheight;
                SizeWidth = sizewidth;
            }
        }

        public struct GreenHDMQuickMatchAndViewResult
        {
            public string BestTagName;
            public double BestTagScore;
            public double Threshold;
            public double SizeHeight;
            public double SizeWidth;

            public GreenHDMQuickMatchAndViewResult(string besttagname, double besttagscore, double threshold, double sizeheight, double sizewidth)
            {
                BestTagName = besttagname;
                BestTagScore = besttagscore;
                Threshold = threshold;
                SizeHeight = sizeheight;
                SizeWidth = sizewidth;
            }
        }

        // JK-AddResultOFGreen-2023.07.12- End

        // JK-AddResultOFRed-2023.07.12- Start // Red HDM, Focused Supervised, Focused Unsupervised.
        public struct RedHDMRegionResult
        {
            public string Name;
            public double Score;
            public double Area;
            public double CenterX;
            public double CenterY;
            public int OuterCount;
            public int InnerCount;

            public RedHDMRegionResult(string name, double score, double area, double centerx, double centery, int outercount, int innercount)
            {
                Name = name;
                Score = score;
                Area = area;
                CenterX = centerx;
                CenterY = centery;
                OuterCount = outercount;
                InnerCount = innercount;
            }
        }

        public struct RedFocusedSupervisedRegionResult
        {
            public string Name;
            public double Score;
            public double Area;
            public double CenterX;
            public double CenterY;
            public int OuterCount;
            public int InnerCount;

            public RedFocusedSupervisedRegionResult(string name, double score, double area, double centerx, double centery, int outercount, int innercount)
            {
                Name = name;
                Score = score;
                Area = area;
                CenterX = centerx;
                CenterY = centery;
                OuterCount = outercount;
                InnerCount = innercount;
            }
        }

        public struct RedFocusedUnsupervisedRegionResult
        {
            public string Name;
            public double Score;
            public double Area;
            public double CenterX;
            public double CenterY;
            public int OuterCount;
            public int InnerCount;

            public RedFocusedUnsupervisedRegionResult(string name, double score, double area, double centerx, double centery, int outercount, int innercount)
            {
                Name = name;
                Score = score;
                Area = area;
                CenterX = centerx;
                CenterY = centery;
                OuterCount = outercount;
                InnerCount = innercount;
            }
        }

        // JK-AddResultOFRed-2023.07.12- End // Red HDM, Focused Supervised, Focused Unsupervised.



        static void Main(string[] args)
        {
            // JK Test - Start -  using Structure of List type : 2023.07.11
            List<BlueReadMatchFeatureResult> ResultOfBlueReadMatchFeature = new List<BlueReadMatchFeatureResult>();
            // JK-AddResultOFGreen-2023.07.12- Start
            List<GreenHDMMatchAndViewResult> ResultOfGreenHDMMatchAndViewResult = new List<GreenHDMMatchAndViewResult>();
            List<GreenFocusedMatchAndViewResult> ResultOfGreenFocusedMatchAndViewResult = new List<GreenFocusedMatchAndViewResult>();
            List<GreenHDMQuickMatchAndViewResult> ResultOfGreenHDMQuickMatchAndViewResult = new List<GreenHDMQuickMatchAndViewResult>();
            // JK-AddResultOFGreen-2023.07.12- End

            // JK-AddResultOFRed-2023.07.12- Start // Red HDM, Focused Supervised, Focused Unsupervised.

            List<string> RedHDMDetectedRegions = new List<string>(); // int
            List<string> RedFocusedSupervisedDetectedRegions = new List<string>(); // int
            List<string> RedFocusedUnsupervisedDetectedRegions = new List<string>(); // int

            List<RedHDMRegionResult> ResultOfRedHDMRegionResult = new List<RedHDMRegionResult>();
            List<RedFocusedSupervisedRegionResult> ResultOfRedFocusedSupervisedRegionResult = new List<RedFocusedSupervisedRegionResult>();
            List<RedFocusedUnsupervisedRegionResult> ResultOFRedFocusedUnsupervisedRegionResult = new List<RedFocusedUnsupervisedRegionResult>();
            // JK-AddResultOFRed-2023.07.12- End // Red HDM, Focused Supervised, Focused Unsupervised.


            Console.WriteLine($"\nStep 0. Preparation : Create Directory");
            string fBin = "Bin";
            string fCDLS = "Cognex Deep Learning Studio";
            JKCtrlDirectory.CreateDir(fBin);
            JKCtrlDirectory.CreateDir(fCDLS);
            Console.WriteLine($"\n - Complete Step 0 : Created directories");

            Console.WriteLine($"\n*** QA-Get Processing Time:" + DateTime.Now.ToString("yyyy-MM-dd") + " ***\n");
            Console.WriteLine($"\nStep 1. Start getting processing time");
            using (ViDi2.Runtime.Local.Control control = new ViDi2.Runtime.Local.Control(GpuMode.Deferred))
            {
                Console.WriteLine($"\n - initialize GPU Device");
                control.InitializeComputeDevices(GpuMode.SingleDevicePerTool, new List<int>() { });

                /* Getting configuration in system e.g., GPU model, Driver Version, OS etc - It's next task*/
                List<string> TestConfigurationList = new List<string>();
                string tempLine = " ";

                Console.WriteLine($"\n***[Configuration of the current agent in teamcity]***");
                tempLine = $"***[Configuration of the current agent in teamcity]***";
                TestConfigurationList.Add(tempLine);

                Console.WriteLine($"PC OS Info."); // refer to : //www.techiedelight.com/determine-operating-system-csharp/
                tempLine = $"PC OS Info.";
                TestConfigurationList.Add(tempLine);

                if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                {
                    Console.WriteLine(" - OS: Windows");
                    tempLine = " - OS: Windows";
                    TestConfigurationList.Add(tempLine);
                }
                else if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux))
                {
                    Console.WriteLine(" - OS: Linux");
                    tempLine = " - OS: Linux";
                    TestConfigurationList.Add(tempLine);
                }
                else if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
                {
                    Console.WriteLine(" - OS: MacOS");
                    tempLine = $" - OS: MacOS";
                    TestConfigurationList.Add(tempLine);
                }
                Console.WriteLine(" - OSDescription: {0}", RuntimeInformation.OSDescription);
                tempLine = $" - OSDescription: {0}" + RuntimeInformation.OSDescription.ToString();
                TestConfigurationList.Add(tempLine);

                // ********** Notify : If These is not GPU in using Agent e.g., #7, You need to skip this code line.
                Console.WriteLine($"GPU Info.");
                TestConfigurationList.Add($"GPU Info."); //                TestConfigurationList.Add();
                Console.WriteLine($" - Model: " + control.ComputeDevices[0].Name);// Index: control.ComputeDevices[0].Index.ToString()
                TestConfigurationList.Add($" - Model: " + control.ComputeDevices[0].Name.ToString());
                Console.WriteLine($" - Memory: " + control.ComputeDevices[0].Memory);
                TestConfigurationList.Add($" - Memory: " + control.ComputeDevices[0].Memory.ToString());
                Console.WriteLine($" - Opt Memory: " + control.ComputeDevices[0].OptimizedGpuMemory);
                TestConfigurationList.Add($" - Opt Memory: " + control.ComputeDevices[0].OptimizedGpuMemory);
                Console.WriteLine($" - Opt.Mem Status: " + control.ComputeDevices[0].OptimizedGpuMemoryEnabled.ToString());
                TestConfigurationList.Add($" - Opt.Mem Status: " + control.ComputeDevices[0].OptimizedGpuMemoryEnabled.ToString());
                Console.WriteLine($" - Type: " + control.ComputeDevices[0].Type);
                TestConfigurationList.Add($" - Type: " + control.ComputeDevices[0].Type.ToString());
                Console.WriteLine($" - Vers: " + control.ComputeDevices[0].Version);
                TestConfigurationList.Add($" - Vers: " + control.ComputeDevices[0].Version.ToString());

                Console.WriteLine($"VPDL Info.");
                TestConfigurationList.Add($"VPDL Info.");
                Console.WriteLine($" - Version: " + control.CLibraryVersion);
                TestConfigurationList.Add($" - Version: " + control.CLibraryVersion.ToString());

                // JK-Modified-2023.07.17 - Start
                TestConfigurationItems.VPDLVers = control.CLibraryVersion.ToString(); // To add excel file name,                 
                // JK-Modified-2023.07.17 - End


                Console.WriteLine($"License Info.: ");
                TestConfigurationList.Add($"License Info.: ");
                Console.WriteLine($" - SerialNumber: " + control.License.SerialNumber);
                TestConfigurationList.Add($" - SerialNumber: " + control.License.SerialNumber.ToString());
                Console.WriteLine($" - Performance Level: " + control.License.PerformanceLevel.ToString());
                TestConfigurationList.Add($" - Performance Level: " + control.License.PerformanceLevel.ToString());
                Console.WriteLine($" - PreviewChannel: " + control.License.PreviewChannel.ToString());
                TestConfigurationList.Add($" - PreviewChannel: " + control.License.PreviewChannel.ToString());
                Console.WriteLine($" - Vaild Tools Count: " + control.License.Tools.Count.ToString());
                TestConfigurationList.Add($" - Vaild Tools Count: " + control.License.Tools.Count.ToString());
                for (int index = 0; index < control.License.Tools.Count; index++)
                {
                    Console.WriteLine($"\t Tool {index}. " + control.License.Tools.ElementAt(index).Key.ToString());
                    TestConfigurationList.Add($"\t Tool {index}. " + control.License.Tools.ElementAt(index).Key.ToString());
                }
                //TestConfigurationList.Add(control.CLibraryVersion.ToString());
                //TestConfigurationList.Add(control.ComputeDevices[0].ToString());                               
                //Console.WriteLine($"\n - checking test");

                // JK-Modified-2023.07.17 - Start
                Console.WriteLine($"Runtime Workspaces");
                TestConfigurationList.Add($"Runtime Workspaces");

                DirectoryInfo dirRuntimeworkspaceInfo = new DirectoryInfo(@"..\..\..\..\..\TestResource\Runtime\"); // refer to: //timeboxstory.tistory.com/107
                foreach (FileInfo rwsFiles in dirRuntimeworkspaceInfo.GetFiles())
                {
                    Console.WriteLine(rwsFiles.Name);   // Console.WriteLine(rwsFiles.FullName);
                    TestConfigurationList.Add($" - " + rwsFiles.Name);
                }
                Console.WriteLine($"Test Image files Info.");
                TestConfigurationList.Add($"Test Image files Info.");

                TestConfigurationList.Add($"Blue Locate");
                DirectoryInfo dirBLImageInfo = new DirectoryInfo(@"..\..\..\..\..\TestResource\Images_BlueLocate\");
                foreach (FileInfo imgBL in dirBLImageInfo.GetFiles())
                {
                    Console.WriteLine(imgBL.Name);
                    TestConfigurationList.Add($" - " + imgBL.Name);
                }
                TestConfigurationList.Add($"Blue Read");
                DirectoryInfo dirBRImageInfo = new DirectoryInfo(@"..\..\..\..\..\TestResource\Images_BlueRead\");
                foreach (FileInfo imgBR in dirBRImageInfo.GetFiles())
                {
                    Console.WriteLine(imgBR.Name);
                    TestConfigurationList.Add($" - " + imgBR.Name);
                }
                TestConfigurationList.Add($"Grenn HDM/Focused/HDM Quick");
                DirectoryInfo dirGImageInfo = new DirectoryInfo(@"..\..\..\..\..\TestResource\Images_Green\");
                foreach (FileInfo imgG in dirGImageInfo.GetFiles())
                {
                    Console.WriteLine(imgG.Name);
                    TestConfigurationList.Add($" - " + imgG.Name);
                }
                TestConfigurationList.Add($"Red HDM/Focused Supervised/Focused Unsupervised");
                DirectoryInfo dirRImageInfo = new DirectoryInfo(@"..\..\..\..\..\TestResource\Images_Red\");
                foreach (FileInfo imgR in dirRImageInfo.GetFiles())
                {
                    Console.WriteLine(imgR.Name);
                    TestConfigurationList.Add($" - " + imgR.Name);
                }
                // JK-Modified-2023.07.17 - Start


                Stopwatch stopWatch = new Stopwatch();

                // Blue Locate - Start // BlueLocate
                Console.WriteLine($"\n - Blue Locate - Start");

                // JK-AddResult-Start - 2023.07.06
                // The number of features
                // Feature Name and Score

                // View Inspector - Feature
                List<string> BlueLocateNumFeatures = new List<string>(); // int
                List<string> BlueLocateFeaturesName = new List<string>(); // string
                List<string> BlueLocateFeaturesScore = new List<string>(); // doublue
                List<string> BlueLocateFeaturesPosX = new List<string>(); // doublue
                List<string> BlueLocateFeaturesPosY = new List<string>(); // doublue
                List<string> BlueLocateFeaturesAngle = new List<string>(); // doublue
                List<string> BlueLocateFeaturesSizeHeight = new List<string>(); // dounle
                List<string> BlueLocateFeaturesSizeWidth = new List<string>(); // double

                // View Inspector - Node Model Match(es)
                List<string> BlueLocareMatchModelName = new List<string>(); // string
                List<string> BlueLocareMatchScore = new List<string>(); // double

                // View Inspector - View Properties
                List<string> BlueLocareViewHeight = new List<string>(); // double
                List<string> BlueLocareViewWidth = new List<string>(); // double


                List<string> FirstFeaturesName = new List<string>(); // string
                List<string> FirstFeaturesScore = new List<string>(); // doublue
                List<string> FirstFeaturesPosX = new List<string>(); // doublue
                List<string> FirstFeaturesPosY = new List<string>(); // doublue

                List<string> SecondFeaturesName = new List<string>(); // string
                List<string> SecondFeaturesScore = new List<string>(); // doublue
                List<string> SecondFeaturesPosX = new List<string>(); // doublue
                List<string> SecondFeaturesPosY = new List<string>(); // doublue

                // JK-AddResult-End - 2023.07.06

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

                            // JK-AddResult-Start - 2023.07.06

                            IBlueMarking blueMarking = sample.Markings[BlueLocateTool.Name] as IBlueMarking;
                            foreach (IBlueView view in blueMarking.Views)
                            {
                                //Console.WriteLine("BlueLocate - Getting the result data of feautres : NumF/Name/Score/PosXY/Angle/Size");

                                BlueLocateNumFeatures.Add(view.Features.Count.ToString()); // Add the number of features in List after getting a count of features.

                                foreach (IFeature feature in view.Features) // The number of features is two.
                                {
                                    BlueLocateFeaturesName.Add(feature.Name);
                                    BlueLocateFeaturesScore.Add(feature.Score.ToString());
                                    BlueLocateFeaturesPosX.Add(feature.Position.X.ToString());
                                    BlueLocateFeaturesPosY.Add(feature.Position.Y.ToString());
                                    BlueLocateFeaturesAngle.Add(feature.Angle.ToString());
                                    BlueLocateFeaturesSizeHeight.Add(feature.Size.Height.ToString());
                                    BlueLocateFeaturesSizeWidth.Add(feature.Size.Width.ToString());

                                    var test = feature.Size.Height.ToString();
                                }
                                foreach (IMatch match in view.Matches)
                                {
                                    BlueLocareMatchModelName.Add(match.ModelName);
                                    BlueLocareMatchScore.Add(match.Score.ToString());
                                }

                                // View Inspector - View Properties
                                BlueLocareViewHeight.Add(view.Size.Height.ToString());
                                BlueLocareViewWidth.Add(view.Size.Width.ToString());
                            }
                            // JK-AddResult-End - 2023.07.06


                        }
                    }
                }
                double avgBlueLocate = sumBlueLocate / (double)countBlueLocate;
                Console.WriteLine(" - Processing Time Average({0} images): {1} [msec]", (int)countBlueLocate, avgBlueLocate);
                // Blue Locate - End


                // Blue Read - Start // BlueRead
                Console.WriteLine($"\n - Blue Read - Start");

                // JK Test - Start -  using Structure of List type : 2023.07.11
                //List<BlueReadMatchFeatureResult> ResultOfBlueReadMatchFeature = new List<BlueReadMatchFeatureResult>();

                // JK-AddResult- Start - 2023.07.11
                // View Inspector - Feature
                List<string> BlueReadNumFeatures = new List<string>(); // int
                List<string> BlueReadFeaturesName = new List<string>(); // string
                List<string> BlueReadFeaturesScore = new List<string>(); // doublue
                List<string> BlueReadFeaturesPosX = new List<string>(); // doublue
                List<string> BlueReadFeaturesPosY = new List<string>(); // doublue
                List<string> BlueReadFeaturesAngle = new List<string>(); // doublue
                List<string> BlueReadFeaturesSizeHeight = new List<string>(); // dounle
                List<string> BlueReadFeaturesSizeWidth = new List<string>(); // double


                // View Inspector - Model Match(es) + Matching result of each festures
                List<string> BlueReadMatchModelName = new List<string>(); // string
                List<string> BlueReadMatchScore = new List<string>(); // double

                List<string> BlueReadMatchCountFeatures = new List<string>(); // int                
                List<string> BlueReadMatchFetureString = new List<string>(); // string

                List<string> BlueReadMatchFetureName = new List<string>(); // string
                List<string> BlueReadMatchFeaturesScore = new List<string>(); // doublue
                List<string> BlueReadMatchFeaturesPosX = new List<string>(); // doublue
                List<string> BlueReadMatchFeaturesPosY = new List<string>(); // doublue
                List<string> BlueReadMatchFeaturesAngle = new List<string>(); // doublue
                List<string> BlueReadMatchFeaturesSizeHeight = new List<string>(); // dounle
                List<string> BlueReadMatchFeaturesSizeWidth = new List<string>(); // double
                // JK-AddResult-End - 2023.07.11



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

                            // JK-AddResult- Start - 2023.07.11
                            IBlueMarking blueMarking = sample.Markings[BlueReadTool.Name] as IBlueMarking;
                            // JK Test - Start -  using Structure of List type : 2023.07.11
                            //List<BlueReadMatchFeatureResult> ResultOfBlueReadMatchFeature = new List<BlueReadMatchFeatureResult>();
                            //

                            foreach (IBlueView view in blueMarking.Views)
                            {
                                BlueReadNumFeatures.Add(view.Features.Count.ToString());
                                foreach (IFeature feature in view.Features)
                                {
                                    BlueReadFeaturesName.Add(feature.Name);
                                    BlueReadFeaturesScore.Add(feature.Score.ToString());
                                    BlueReadFeaturesPosX.Add(feature.Position.X.ToString());
                                    BlueReadFeaturesPosY.Add(feature.Position.Y.ToString());
                                    BlueReadFeaturesAngle.Add(feature.Angle.ToString());
                                    BlueReadFeaturesSizeHeight.Add(feature.Size.Height.ToString());
                                    BlueReadFeaturesSizeWidth.Add(feature.Size.Width.ToString());
                                }
                                foreach (IMatch match in view.Matches)
                                {
                                    //BlueReadMatchModelName.Add(match.ModelName);
                                    //BlueReadMatchScore.Add(match.Score.ToString());
                                    BlueReadMatchCountFeatures.Add(match.Features.Count.ToString());
                                    BlueReadMatchFetureString.Add(match.FeatureString);
                                    foreach (IFeature featureResultFromMatch in match.Features)
                                    {
                                        BlueReadMatchFetureName.Add(featureResultFromMatch.Name);
                                        BlueReadMatchFeaturesScore.Add(featureResultFromMatch.Score.ToString());
                                        BlueReadMatchFeaturesPosX.Add(featureResultFromMatch.Position.X.ToString());
                                        BlueReadMatchFeaturesPosY.Add(featureResultFromMatch.Position.Y.ToString());
                                        BlueReadMatchFeaturesAngle.Add(featureResultFromMatch.Angle.ToString());
                                        BlueReadMatchFeaturesSizeHeight.Add(featureResultFromMatch.Size.Height.ToString());
                                        BlueReadMatchFeaturesSizeWidth.Add(featureResultFromMatch.Size.Width.ToString());
                                        // JK Test - Start -  using Structure of List type : 2023.07.11
                                        ResultOfBlueReadMatchFeature.Add(new BlueReadMatchFeatureResult(
                                            featureResultFromMatch.Name,
                                            featureResultFromMatch.Score,
                                            featureResultFromMatch.Position.X,
                                            featureResultFromMatch.Position.Y,
                                            featureResultFromMatch.Angle,
                                            featureResultFromMatch.Size.Height,
                                            featureResultFromMatch.Size.Width
                                            ));
                                    }
                                }
                            }
                            // JK-AddResult- End - 2023.07.11
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

                            // JK-AddResultOFGreen-2023.07.12- Start
                            IGreenMarking greenHDMMarking = sample.Markings[GreenHDMTool.Name] as IGreenMarking;

                            foreach (IGreenView view in greenHDMMarking.Views)
                            {
                                ////Console.WriteLine("\n\r");
                                //Console.WriteLine($"View Inspector - Marking Information");
                                //foreach (ITag match in view.Tags) // view.tags[4] = A, B, C, D
                                //{
                                //    Console.WriteLine($"\t{match.Name}: {match.Score} [%]");
                                //}
                                //Console.WriteLine($"\t>> BestTag/Score: {view.BestTag.Name}/{view.BestTag.Score}");
                                //Console.WriteLine($"\nView Inspector - View Properties");
                                //Console.WriteLine($"\tHeight: {view.Size.Height}");
                                //Console.WriteLine($"\tWidth: {view.Size.Width}");
                                //Console.WriteLine($"\tPose: {view.Pose}");
                                //Console.WriteLine($"\nView Inspector - Other imformation of View Properties");
                                //Console.WriteLine($"\tThreshold: {view.Threshold}");
                                //Console.WriteLine($"\tUncertainty: {view.Uncertainty}");
                                //Console.WriteLine($"\tIsLabeled: {view.IsLabeled}");
                                //Console.WriteLine($"\tBookmark: {view.Bookmark}");
                                //Console.WriteLine($"\tHasMask: {view.HasMask}");

                                ResultOfGreenHDMMatchAndViewResult.Add(new GreenHDMMatchAndViewResult(
                                    view.BestTag.Name,
                                    view.BestTag.Score,
                                    view.Threshold,
                                    view.Size.Height,
                                    view.Size.Width
                                    ));
                            }
                            // JK-AddResultOFGreen-2023.07.12- End


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

                            // JK-AddResultOFGreen-2023.07.12- Start
                            IGreenMarking greenFocusedMarking = sample.Markings[GreenFocusedTool.Name] as IGreenMarking;

                            foreach (IGreenView view in greenFocusedMarking.Views)
                            {
                                ResultOfGreenFocusedMatchAndViewResult.Add(new GreenFocusedMatchAndViewResult(
                                    view.BestTag.Name,
                                    view.BestTag.Score,
                                    view.Threshold,
                                    view.Size.Height,
                                    view.Size.Width
                                    ));
                            }
                            // JK-AddResultOFGreen-2023.07.12- End        
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

                            // JK-AddResultOFGreen-2023.07.12- Start
                            IGreenMarking greenHDMQMarking = sample.Markings[GreenHDMQTool.Name] as IGreenMarking;

                            foreach (IGreenView view in greenHDMQMarking.Views)
                            {
                                ResultOfGreenHDMQuickMatchAndViewResult.Add(new GreenHDMQuickMatchAndViewResult(
                                    view.BestTag.Name,
                                    view.BestTag.Score,
                                    view.Threshold,
                                    view.Size.Height,
                                    view.Size.Width
                                    ));
                            }
                            // JK-AddResultOFGreen-2023.07.12- End
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

                            // JK-AddResultOFRed-2023.07.12- Start // Red HDM, Focused Supervised, Focused Unsupervised.

                            IRedMarking redHDMMarking = sample.Markings[RedHDMTool.Name] as IRedMarking;
                            foreach (IRedView view in redHDMMarking.Views)
                            {
                                //Console.WriteLine($"\tDetected Regions: {view.Regions.Count}\n");
                                RedHDMDetectedRegions.Add(view.Regions.Count.ToString());
                                foreach (IRegion rhdmregion in view.Regions)
                                {
                                    //Console.WriteLine($"\tName: {region.Name}");
                                    //Console.WriteLine($"\tScore: {region.Score}");
                                    //Console.WriteLine($"\tArea: {region.Area}");
                                    //Console.WriteLine($"\tX(Center): {region.Center.X}");
                                    //Console.WriteLine($"\tY(Center): {region.Center.Y}");
                                    //Console.WriteLine($"\tOuter Polygon: {region.Outer.Count}");
                                    //Console.WriteLine($"\tInner Polygon: {region.Inners.Count}\n");

                                    ResultOfRedHDMRegionResult.Add(new RedHDMRegionResult(
                                        rhdmregion.Name,
                                        rhdmregion.Score,
                                        rhdmregion.Area,
                                        rhdmregion.Center.X,
                                        rhdmregion.Center.Y,
                                        rhdmregion.Outer.Count,
                                        rhdmregion.Inners.Count
                                        ));
                                }
                                //Console.WriteLine($"View Inspector - View Properties");
                                //Console.WriteLine($"\tHeight: {view.Size.Height}");
                                //Console.WriteLine($"\tWidth: {view.Size.Width}");
                                //Console.WriteLine($"\tPose: {view.Pose}");
                                //Console.WriteLine($"\tHasmask: {view.HasMask}");
                                //Console.WriteLine($"\nView Inspector - Other imformation of View Properties");
                                //Console.WriteLine($"\tThreshold: {view.Threshold.Lower}");
                                //Console.WriteLine($"\tThreshold: {view.Threshold.Upper}");
                                //Console.WriteLine($"\tUncertainty: {view.Uncertainty}");
                                //Console.WriteLine($"\tIsLabeled: {view.IsLabeled}");
                                //Console.WriteLine($"\tBookmark: {view.Bookmark}");
                            }

                            // JK-AddResultOFRed-2023.07.12- End // Red HDM, Focused Supervised, Focused Unsupervised.

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

                            // JK-AddResultOFRed-2023.07.12- Start // Red HDM, Focused Supervised, Focused Unsupervised.
                            IRedMarking redFSMarking = sample.Markings[RedFSuTool.Name] as IRedMarking;
                            foreach (IRedView view in redFSMarking.Views)
                            {
                                RedFocusedSupervisedDetectedRegions.Add(view.Regions.Count.ToString());
                                foreach (IRegion rfsregion in view.Regions)
                                {
                                    ResultOfRedFocusedSupervisedRegionResult.Add(new RedFocusedSupervisedRegionResult(
                                        rfsregion.Name,
                                        rfsregion.Score,
                                        rfsregion.Area,
                                        rfsregion.Center.X,
                                        rfsregion.Center.Y,
                                        rfsregion.Outer.Count,
                                        rfsregion.Inners.Count
                                        ));
                                }
                            }
                            // JK-AddResultOFRed-2023.07.12- End // Red HDM, Focused Supervised, Focused Unsupervised.                           

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

                            // JK-AddResultOFRed-2023.07.12- Start // Red HDM, Focused Supervised, Focused Unsupervised.
                            IRedMarking redFUMarking = sample.Markings[RedFUnTool.Name] as IRedMarking;
                            foreach (IRedView view in redFUMarking.Views)
                            {
                                RedFocusedUnsupervisedDetectedRegions.Add(view.Regions.Count.ToString());
                                foreach (IRegion rfuregion in view.Regions)
                                {
                                    ResultOFRedFocusedUnsupervisedRegionResult.Add(new RedFocusedUnsupervisedRegionResult(
                                        rfuregion.Name,
                                        rfuregion.Score,
                                        rfuregion.Area,
                                        rfuregion.Center.X,
                                        rfuregion.Center.Y,
                                        rfuregion.Outer.Count,
                                        rfuregion.Inners.Count
                                        ));
                                }
                            }
                            // JK-AddResultOFRed-2023.07.12- End // Red HDM, Focused Supervised, Focused Unsupervised.
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

                //JK-AddResult-Start - 2023.07.06
                // Blue Locate - Feature's result
                // View Inspector - Feature
                var getBlueLocateNumFeatures = new List<string>();
                // Odd number : Tail
                var getBlueLocateFeaturesNameOddNum = new List<string>();
                var getBlueLocateFeaturesScoreOddNum = new List<string>();
                var getBlueLocateFeaturesPosXOddNum = new List<string>();
                var getBlueLocateFeaturesPosYOddNum = new List<string>();
                var getBlueLocateFeaturesAngleOddNum = new List<string>();
                var getBlueLocateFeaturesSizeHeightOddNum = new List<string>();
                var getBlueLocateFeaturesSizeWidthOddNum = new List<string>();
                // Even number : Head                
                var getBlueLocateFeaturesNameEvenNum = new List<string>();
                var getBlueLocateFeaturesScoreEvenNum = new List<string>();
                var getBlueLocateFeaturesPosXEvenNum = new List<string>();
                var getBlueLocateFeaturesPosYEvenNum = new List<string>();
                var getBlueLocateFeaturesAngleEvenNum = new List<string>();
                var getBlueLocateFeaturesSizeHeightEvenNum = new List<string>();
                var getBlueLocateFeaturesSizeWidthEvenNum = new List<string>();

                // View Inspector - Node Model Match(es)
                var getBlueLocareMatchModelName = new List<string>();
                var getBlueLocareMatchScore = new List<string>();
                // View Inspector - View Properties
                var getBlueLocareViewHeight = new List<string>();
                var getBlueLocareViewWidth = new List<string>();

                for (int indexcount = 0; indexcount < Constants.RepeatProcess; indexcount++)
                {
                    getBlueLocateNumFeatures.Add(BlueLocateNumFeatures[indexcount].ToString());
                    // View Inspector - Feature

                    //getBlueLocateFeaturesName.Add(BlueLocateFeaturesName[indexcount].ToString());
                    //getBlueLocateFeaturesScore.Add(BlueLocateFeaturesScore[indexcount].ToString());
                    //getBlueLocateFeaturesPosX.Add(BlueLocateFeaturesPosX[indexcount].ToString());
                    //getBlueLocateFeaturesPosY.Add(BlueLocateFeaturesPosY[indexcount].ToString());
                    //getBlueLocateFeaturesAngle.Add(BlueLocateFeaturesAngle[indexcount].ToString());
                    //getBlueLocateFeaturesSizeHeight.Add(BlueLocateFeaturesSizeHeight[indexcount].ToString());
                    //getBlueLocateFeaturesSizeWidth.Add(BlueLocateFeaturesSizeWidth[indexcount].ToString());

                    // View Inspector - Node Model Match(es)
                    getBlueLocareMatchModelName.Add(BlueLocareMatchModelName[indexcount].ToString());
                    getBlueLocareMatchScore.Add(BlueLocareMatchScore[indexcount].ToString());
                    // View Inspector - View Properties
                    getBlueLocareViewHeight.Add(BlueLocareViewHeight[indexcount].ToString());
                    getBlueLocareViewWidth.Add(BlueLocareViewWidth[indexcount].ToString());
                }
                for (int indexcount = 0; indexcount < (Constants.RepeatProcess * 2); indexcount++)
                {
                    if ((indexcount % 2) == 0) // checking even number
                    {
                        getBlueLocateFeaturesNameEvenNum.Add(BlueLocateFeaturesName[indexcount].ToString());
                        getBlueLocateFeaturesScoreEvenNum.Add(BlueLocateFeaturesScore[indexcount].ToString());
                        getBlueLocateFeaturesPosXEvenNum.Add(BlueLocateFeaturesPosX[indexcount].ToString());
                        getBlueLocateFeaturesPosYEvenNum.Add(BlueLocateFeaturesPosY[indexcount].ToString());
                        getBlueLocateFeaturesAngleEvenNum.Add(BlueLocateFeaturesAngle[indexcount].ToString());
                        getBlueLocateFeaturesSizeHeightEvenNum.Add(BlueLocateFeaturesSizeHeight[indexcount].ToString());
                        getBlueLocateFeaturesSizeWidthEvenNum.Add(BlueLocateFeaturesSizeWidth[indexcount].ToString());
                    }
                    else // checking odd number
                    {
                        getBlueLocateFeaturesNameOddNum.Add(BlueLocateFeaturesName[indexcount].ToString());
                        getBlueLocateFeaturesScoreOddNum.Add(BlueLocateFeaturesScore[indexcount].ToString());
                        getBlueLocateFeaturesPosXOddNum.Add(BlueLocateFeaturesPosX[indexcount].ToString());
                        getBlueLocateFeaturesPosYOddNum.Add(BlueLocateFeaturesPosY[indexcount].ToString());
                        getBlueLocateFeaturesAngleOddNum.Add(BlueLocateFeaturesAngle[indexcount].ToString());
                        getBlueLocateFeaturesSizeHeightOddNum.Add(BlueLocateFeaturesSizeHeight[indexcount].ToString());
                        getBlueLocateFeaturesSizeWidthOddNum.Add(BlueLocateFeaturesSizeWidth[indexcount].ToString());
                    }
                }
                Console.WriteLine("\nchecking result - list buffer");
                //JK-AddResult- End - 2023.07.06 // Blue Locate


                //JK-AddResult- Start - 2023.07.11 // Blue Read
                // 아래 내용으로 저장괸 구조체 리스트에 Blue read 매칭 결과가 정상적으로 입력되었는지 확인함.

                List<string> getBlueReadMatchCountFeatures = new List<string>();
                for (int indexcount = 0; indexcount < Constants.RepeatProcess; indexcount++)
                {
                    getBlueReadMatchCountFeatures.Add(BlueReadMatchCountFeatures[indexcount].ToString());
                }

                List<BlueReadMatchFeatureResult> getBlueReadMatchReault = new List<BlueReadMatchFeatureResult>();
                foreach (var mresult in ResultOfBlueReadMatchFeature) // m.blog.naver.com/vesmir/222442589978
                {
                    //Console.WriteLine($"\t\t\t Blue Read Matching results: {mresult.Name}, {mresult.Score}, {mresult.PosX}, {mresult.PosY}, {mresult.Angle}, {mresult.SizeHeight}, {mresult.SizeWidth} \n");                    
                    getBlueReadMatchReault.Add(new BlueReadMatchFeatureResult(
                        mresult.Name,
                        mresult.Score,
                        mresult.PosX,
                        mresult.PosY,
                        mresult.Angle,
                        mresult.SizeHeight,
                        mresult.SizeWidth
                        ));
                }
                //JK-AddResult- End - 2023.07.11 // Blue Read

                // JK-AddResultOFGreen-2023.07.12- Starat // Green HDM, Focused, HDM Quick
                List<GreenHDMMatchAndViewResult> getGreenHDMMatchAndViewResult = new List<GreenHDMMatchAndViewResult>();
                foreach (var ghdmviewresult in ResultOfGreenHDMMatchAndViewResult)
                {
                    getGreenHDMMatchAndViewResult.Add(new GreenHDMMatchAndViewResult(
                        ghdmviewresult.BestTagName,
                        ghdmviewresult.BestTagScore,
                        ghdmviewresult.Threshold,
                        ghdmviewresult.SizeHeight,
                        ghdmviewresult.SizeWidth
                        ));
                }

                List<GreenFocusedMatchAndViewResult> getGreenFocusedMatchAndViewResult = new List<GreenFocusedMatchAndViewResult>();
                foreach (var gfviewresult in ResultOfGreenFocusedMatchAndViewResult)
                {
                    getGreenFocusedMatchAndViewResult.Add(new GreenFocusedMatchAndViewResult(
                        gfviewresult.BestTagName,
                        gfviewresult.BestTagScore,
                        gfviewresult.Threshold,
                        gfviewresult.SizeHeight,
                        gfviewresult.SizeWidth
                        ));
                }

                List<GreenHDMQuickMatchAndViewResult> getGreenHDMQuickMatchAndViewResult = new List<GreenHDMQuickMatchAndViewResult>();
                foreach (var ghdmqresult in ResultOfGreenHDMQuickMatchAndViewResult)
                {
                    getGreenHDMQuickMatchAndViewResult.Add(new GreenHDMQuickMatchAndViewResult(
                        ghdmqresult.BestTagName,
                        ghdmqresult.BestTagScore,
                        ghdmqresult.Threshold,
                        ghdmqresult.SizeHeight,
                        ghdmqresult.SizeWidth
                        ));
                }
                // JK-AddResultOFGreen-2023.07.12- End // Green HDM, Focused, HDM Quick

                // JK-AddResultOFRed-2023.07.12- Start // Red HDM, Focused Supervised, Focused Unsupervised.

                List<string> getRedHDMDetectedRegions = new List<string>(); // int
                List<string> getRedFocusedSupervisedDetectedRegions = new List<string>(); // int
                List<string> getRedFocusedUnsupervisedDetectedRegions = new List<string>(); // int

                for (int indexcount = 0; indexcount < Constants.RepeatProcess; indexcount++)
                {
                    getRedHDMDetectedRegions.Add(RedHDMDetectedRegions[indexcount].ToString());
                    getRedFocusedSupervisedDetectedRegions.Add(RedFocusedSupervisedDetectedRegions[indexcount].ToString());
                    getRedFocusedUnsupervisedDetectedRegions.Add(RedFocusedUnsupervisedDetectedRegions[indexcount].ToString());
                }

                List<RedHDMRegionResult> getRedHDMRegionResult = new List<RedHDMRegionResult>();
                foreach (var rhdmresult in ResultOfRedHDMRegionResult)
                {
                    getRedHDMRegionResult.Add(new RedHDMRegionResult(
                        rhdmresult.Name,
                        rhdmresult.Score,
                        rhdmresult.Area,
                        rhdmresult.CenterX,
                        rhdmresult.CenterY,
                        rhdmresult.OuterCount,
                        rhdmresult.InnerCount
                        ));
                }

                List<RedFocusedSupervisedRegionResult> getRedFocusedSupervisedRegionResult = new List<RedFocusedSupervisedRegionResult>();
                foreach (var rfsresult in ResultOfRedFocusedSupervisedRegionResult)
                {
                    getRedFocusedSupervisedRegionResult.Add(new RedFocusedSupervisedRegionResult(
                        rfsresult.Name,
                        rfsresult.Score,
                        rfsresult.Area,
                        rfsresult.CenterX,
                        rfsresult.CenterY,
                        rfsresult.OuterCount,
                        rfsresult.InnerCount
                        ));
                }

                List<RedFocusedUnsupervisedRegionResult> getRedFocusedUnsupervisedRegionResult = new List<RedFocusedUnsupervisedRegionResult>();
                foreach (var rfuresult in ResultOFRedFocusedUnsupervisedRegionResult)
                {
                    getRedFocusedUnsupervisedRegionResult.Add(new RedFocusedUnsupervisedRegionResult(
                        rfuresult.Name,
                        rfuresult.Score,
                        rfuresult.Area,
                        rfuresult.CenterX,
                        rfuresult.CenterY,
                        rfuresult.OuterCount,
                        rfuresult.InnerCount
                        ));
                }
                // JK-AddResultOFRed-2023.07.12- End // Red HDM, Focused Supervised, Focused Unsupervised.

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
                string strExcelFileName = "QAGetProcessingTime_VPDL_" + TestConfigurationItems.VPDLVers + "_" + getDateInfo + ".xlsx";
                string strExcelFileDirectory = Path.GetFullPath(@"..\..\..\..\..\TestResultCSV\") + strExcelFileName;   // Refer to - Processing file path name in using C# : //myoung-min.tistory.com/45
                Console.WriteLine(strExcelFileDirectory);

                // This under line is the existance before 2023.07.06
                //ExcelDataEPPlusRedTools(getResultListRedHDM, getResultListRedFSu, getResultListRedFUn, getResultListGreenHDM, getResultListGreenFocused, getResultListGreenHDMQucik, getResultListBlueLocate, getResultListBlueRead, strExcelFileDirectory); // Adding Green HDM Tool in the create epplus excel  - 20230508

                // JK-AddResult-Start - 2023.07.06
                ExcelDataEPPlusRedTools(
                    getResultListRedHDM,
                    getResultListRedFSu,
                    getResultListRedFUn,
                    getResultListGreenHDM,
                    getResultListGreenFocused,
                    getResultListGreenHDMQucik,
                    getResultListBlueLocate,

                    //JK-AddResult-Start - 2023.07.06                                       
                    getBlueLocateNumFeatures,
                    getBlueLocareMatchModelName,
                    getBlueLocareMatchScore,
                    getBlueLocareViewHeight,
                    getBlueLocareViewWidth,

                    getBlueLocateFeaturesNameEvenNum,
                    getBlueLocateFeaturesScoreEvenNum,
                    getBlueLocateFeaturesPosXEvenNum,
                    getBlueLocateFeaturesPosYEvenNum,
                    getBlueLocateFeaturesAngleEvenNum,
                    getBlueLocateFeaturesSizeHeightEvenNum,
                    getBlueLocateFeaturesSizeWidthEvenNum,

                    getBlueLocateFeaturesNameOddNum,
                    getBlueLocateFeaturesScoreOddNum,
                    getBlueLocateFeaturesPosXOddNum,
                    getBlueLocateFeaturesPosYOddNum,
                    getBlueLocateFeaturesAngleOddNum,
                    getBlueLocateFeaturesSizeHeightOddNum,
                    getBlueLocateFeaturesSizeWidthOddNum,
                    //JK-AddResult-End - 2023.07.06

                    //JK-AddResult-Start - 2023.07.11
                    getBlueReadMatchCountFeatures,
                    getBlueReadMatchReault,
                    //JK-AddResult-End - 2023.07.11

                    // JK-AddResultOFGreen-2023.07.12- Start // Green HDM, Focused, HDM Quick
                    getGreenHDMMatchAndViewResult,
                    getGreenFocusedMatchAndViewResult,
                    getGreenHDMQuickMatchAndViewResult,
                    // JK-AddResultOFGreen-2023.07.12- End // Green HDM, Focused, HDM Quick

                    // JK-AddResultOFRed-2023.07.12- Start // Red HDM, Focused Supervised, Focused Unsupervised.
                    getRedHDMDetectedRegions,
                    getRedFocusedSupervisedDetectedRegions,
                    getRedFocusedUnsupervisedDetectedRegions,
                    getRedHDMRegionResult,
                    getRedFocusedSupervisedRegionResult,
                    getRedFocusedUnsupervisedRegionResult,
                    // JK-AddResultOFRed-2023.07.12- Start // Red HDM, Focused Supervised, Focused Unsupervised.

                    getResultListBlueRead,
                    strExcelFileDirectory
                    ); // Adding Green HDM Tool in the create epplus excel  - 20230508

                TestConfiguration(TestConfigurationList, strExcelFileDirectory); // saving test configuration

                Console.WriteLine("\nStep 4. Complete QA Test - Get Processing time of Red Tool");
            }
        }

        private static void TestConfiguration(List<string> getTestConfigurationList, string savePath)
        {
            FileInfo existFile = new FileInfo(savePath);
            using (ExcelPackage excelPackage = new ExcelPackage(existFile)) // refer to //riptutorial.com/epplus
            {
                // Create TestPC's Configuration sheet
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("TestConfiguration");

                // JK-Modified-2023.07.17 - Start

                using (ExcelRange Rng = worksheet.Cells[1, 1, 40, 1]) // PC OS Info.
                {
                    Rng.Style.Font.Italic = true;
                }
                using (ExcelRange Rng = worksheet.Cells[1, 1, 1, 1]) // Title
                {
                    Rng.Style.Font.Size = 11;
                    Rng.Style.Font.Bold = true;
                    Rng.Style.Font.Italic = false;
                }
                using (ExcelRange Rng = worksheet.Cells[2, 1, 2, 1]) // PC OS Info.
                {
                    Rng.Style.Font.Size = 11;
                    Rng.Style.Font.Bold = true;
                    Rng.Style.Font.Italic = false;
                }

                using (ExcelRange Rng = worksheet.Cells[5, 1, 5, 1]) // GUI Info.
                {
                    Rng.Style.Font.Size = 11;
                    Rng.Style.Font.Bold = true;
                    Rng.Style.Font.Italic = false;
                }

                using (ExcelRange Rng = worksheet.Cells[12, 1, 12, 1]) // VPDL Info.
                {
                    Rng.Style.Font.Size = 11;
                    Rng.Style.Font.Bold = true;
                    Rng.Style.Font.Italic = false;
                }
                using (ExcelRange Rng = worksheet.Cells[14, 1, 14, 1]) // License Info.
                {
                    Rng.Style.Font.Size = 11;
                    Rng.Style.Font.Bold = true;
                    Rng.Style.Font.Italic = false;
                }
                using (ExcelRange Rng = worksheet.Cells[23, 1, 23, 1]) // Runtime workspaces
                {
                    Rng.Style.Font.Size = 11;
                    Rng.Style.Font.Bold = true;
                    Rng.Style.Font.Italic = false;
                }
                using (ExcelRange Rng = worksheet.Cells[32, 1, 32, 1]) // Test Image files Info.
                {
                    Rng.Style.Font.Size = 11;
                    Rng.Style.Font.Bold = true;
                    Rng.Style.Font.Italic = false;
                }
                // JK-Modified-2023.07.17 - End

                // Fill in the system's information
                int col = 0;
                for (int row = 0; row < getTestConfigurationList.Count; row++) // 2023.07.17 getTestConfigurationList.Count=40 >> 0~39
                    worksheet.Cells[row + 1, col + 1].Value = getTestConfigurationList.ElementAt(row);
                excelPackage.Save();
            }
        }

        private static void ExcelDataEPPlusRedTools(
            List<string> GetPTimesRedHDM,
            List<string> GetPTimesRedFSu,
            List<string> GetPTimesRedFUn,
            List<string> GetPTimesGreenHDM,
            List<string> GetPTimesGreenFocused,
            List<string> GetPTimesGreenHDMQuick,
            List<string> GetPTimesBlueLocate,

            //JK-AddResult-Start - 2023.07.06
            List<string> GetBlueLocateNumFeatures,
            List<string> GetBlueLocareMatchModelName,
            List<string> GetBlueLocareMatchScore,
            List<string> GetBlueLocareViewHeight,
            List<string> GetBlueLocareViewWidth,

            List<string> GetBlueLocateFeaturesNameEvenNum,
            List<string> GetBlueLocateFeaturesScoreEvenNum,
            List<string> GetBlueLocateFeaturesPosXEvenNum,
            List<string> GetBlueLocateFeaturesPosYEvenNum,
            List<string> GetBlueLocateFeaturesAngleEvenNum,
            List<string> GetBlueLocateFeaturesSizeHeightEvenNum,
            List<string> GetBlueLocateFeaturesSizeWidthEvenNum,

            List<string> GetBlueLocateFeaturesNameOddNum,
            List<string> GetBlueLocateFeaturesScoreOddNum,
            List<string> GetBlueLocateFeaturesPosXOddNum,
            List<string> GetBlueLocateFeaturesPosYOddNum,
            List<string> GetBlueLocateFeaturesAngleOddNum,
            List<string> GetBlueLocateFeaturesSizeHeightOddNum,
            List<string> GetBlueLocateFeaturesSizeWidthOddNum,
            //JK-AddResult-End - 2023.07.06

            //JK-AddResult-Start - 2023.07.11
            List<string> GetBlueReadMatchCountFeatures,
            List<BlueReadMatchFeatureResult> GetBlueReadMatchReault,
            //JK-AddResult-End - 2023.07.11

            // JK-AddResultOFGreen-2023.07.12- Start // Green HDM, Focused, HDM Quick
            List<GreenHDMMatchAndViewResult> GetGreenHDMMatchAndViewResult,
            List<GreenFocusedMatchAndViewResult> GetGreenFocusedMatchAndViewResult,
            List<GreenHDMQuickMatchAndViewResult> GetGreenHDMQuickMatchAndViewResult,
            // JK-AddResultOFGreen-2023.07.12- End // Green HDM, Focused, HDM Quick

            // JK-AddResultOFRed-2023.07.12- Start // Red HDM, Focused Supervised, Focused Unsupervised.
            List<string> GetRedHDMDetectedRegions,
            List<string> GetRedFocusedSupervisedDetectedRegions,
            List<string> GetRedFocusedUnsupervisedDetectedRegions,

            List<RedHDMRegionResult> GetRedHDMRegionResult,
            List<RedFocusedSupervisedRegionResult> GetRedFocusedSupervisedRegionResult,
            List<RedFocusedUnsupervisedRegionResult> GetRedFocusedUnsupervisedRegionResult,
            // JK-AddResultOFRed-2023.07.12- End // Red HDM, Focused Supervised, Focused Unsupervised.

            List<string> GetPTimesBlueRead,
            string savePath)
        {
            Console.WriteLine("JK Test 1. Create Excel File");
            ExcelPackage ExcelPkg = new ExcelPackage();

            // Red - Start
            ExcelWorksheet wsSheetRed = ExcelPkg.Workbook.Worksheets.Add("RedTools");
            using (ExcelRange Rng = wsSheetRed.Cells[1, 1, 1, 1])
            {
                Rng.Value = "Repeat";
                Rng.Style.Font.Size = 11;
                Rng.Style.Font.Bold = true;
                Rng.Style.Font.Italic = true;
                Rng.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(217, 217, 217)); // Color is gray
                //Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(238, 46, 34));
                Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }
            using (ExcelRange Rng = wsSheetRed.Cells[2, 1, 2, 1])
            {
                Rng.Value = "Max.";
                Rng.Style.Font.Size = 11;
                Rng.Style.Font.Bold = true;
                Rng.Style.Font.Italic = true;
                Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }
            using (ExcelRange Rng = wsSheetRed.Cells[3, 1, 3, 1])
            {
                Rng.Value = "Min.";
                Rng.Style.Font.Size = 11;
                Rng.Style.Font.Bold = true;
                Rng.Style.Font.Italic = true;
                Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }
            using (ExcelRange Rng = wsSheetRed.Cells[4, 1, 4, 1])
            {
                Rng.Value = "Avrg.";
                Rng.Style.Font.Size = 11;
                Rng.Style.Font.Bold = true;
                Rng.Style.Font.Italic = true;
                Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }
            using (ExcelRange Rng = wsSheetRed.Cells[1, 2, 1, 2])
            {
                Rng.Value = "Red HDM";
                Rng.Style.Font.Size = 11;
                Rng.Style.Font.Bold = true;
                Rng.Style.Font.Italic = true;
                Rng.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(217, 217, 217));
                //Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(238, 46, 34));
                Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }
            using (ExcelRange Rng = wsSheetRed.Cells[1, 3, 1, 3])
            {
                Rng.Value = "Red FSu";
                Rng.Style.Font.Size = 11;
                Rng.Style.Font.Bold = true;
                Rng.Style.Font.Italic = true;
                Rng.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(217, 217, 217));
                //Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(238, 46, 34));
                Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }
            using (ExcelRange Rng = wsSheetRed.Cells[1, 4, 1, 4])
            {
                Rng.Value = "Red FUn";
                Rng.Style.Font.Size = 11;
                Rng.Style.Font.Bold = true;
                Rng.Style.Font.Italic = true;
                Rng.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(217, 217, 217));
                //Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(238, 46, 34));
                Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }

            // JK - 2023.07.12 - Start

            // Red HMD
            using (ExcelRange Rng = wsSheetRed.Cells[1, 5, 1, 5])
            {
                Rng.Value = "RHDMDetect";
                Rng.Style.Font.Size = 11;
                Rng.Style.Font.Bold = true;
                Rng.Style.Font.Italic = true;
                Rng.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(217, 217, 217));
                Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 60, 60));
                Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }
            using (ExcelRange Rng = wsSheetRed.Cells[1, 6, 1, 6])
            {
                Rng.Value = "RHDMName1";
                Rng.Style.Font.Size = 11;
                Rng.Style.Font.Bold = true;
                Rng.Style.Font.Italic = true;
                Rng.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(217, 217, 217));
                Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 60, 60));
                Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }
            using (ExcelRange Rng = wsSheetRed.Cells[1, 7, 1, 7])
            {
                Rng.Value = "RHDMScore1";
                Rng.Style.Font.Size = 11;
                Rng.Style.Font.Bold = true;
                Rng.Style.Font.Italic = true;
                Rng.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(217, 217, 217));
                Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 60, 60));
                Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }
            using (ExcelRange Rng = wsSheetRed.Cells[1, 8, 1, 8])
            {
                Rng.Value = "RHDMName2";
                Rng.Style.Font.Size = 11;
                Rng.Style.Font.Bold = true;
                Rng.Style.Font.Italic = true;
                Rng.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(217, 217, 217));
                Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 60, 60));
                Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }
            using (ExcelRange Rng = wsSheetRed.Cells[1, 9, 1, 9])
            {
                Rng.Value = "RHDMScore2";
                Rng.Style.Font.Size = 11;
                Rng.Style.Font.Bold = true;
                Rng.Style.Font.Italic = true;
                Rng.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(217, 217, 217));
                Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 60, 60));
                Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }

            // JK-ModifyCodeRedHDM-2023.07.13- Start >>> In case of using Red HDM runtime, The number of defect regions is therr. So, Need to add result cell of third region.
            using (ExcelRange Rng = wsSheetRed.Cells[1, 10, 1, 10])
            {
                Rng.Value = "RHDMName2";
                Rng.Style.Font.Size = 11;
                Rng.Style.Font.Bold = true;
                Rng.Style.Font.Italic = true;
                Rng.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(217, 217, 217));
                Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 60, 60));
                Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }
            using (ExcelRange Rng = wsSheetRed.Cells[1, 11, 1, 11])
            {
                Rng.Value = "RHDMScore2";
                Rng.Style.Font.Size = 11;
                Rng.Style.Font.Bold = true;
                Rng.Style.Font.Italic = true;
                Rng.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(217, 217, 217));
                Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 60, 60));
                Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }

            // JK-ModifyCodeRedHDM-2023.07.13- Start

            //using (ExcelRange Rng = wsSheetRed.Cells[1, 10, 1, 10])
            //{
            //    Rng.Value = "RHDMThreshold";
            //    Rng.Style.Font.Size = 11;
            //    Rng.Style.Font.Bold = true;
            //    Rng.Style.Font.Italic = true;
            //    Rng.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            //    //Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(217, 217, 217));
            //    Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 60, 60));
            //    Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            //}

            // Red Focused Supervised
            using (ExcelRange Rng = wsSheetRed.Cells[1, 12, 1, 12])
            {
                Rng.Value = "RFSDetect";
                Rng.Style.Font.Size = 11;
                Rng.Style.Font.Bold = true;
                Rng.Style.Font.Italic = true;
                Rng.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(217, 217, 217));
                Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 100, 100));
                Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }
            using (ExcelRange Rng = wsSheetRed.Cells[1, 13, 1, 13])
            {
                Rng.Value = "RFSName1";
                Rng.Style.Font.Size = 11;
                Rng.Style.Font.Bold = true;
                Rng.Style.Font.Italic = true;
                Rng.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(217, 217, 217));
                Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 100, 100));
                Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }
            using (ExcelRange Rng = wsSheetRed.Cells[1, 14, 1, 14])
            {
                Rng.Value = "RFSScore1";
                Rng.Style.Font.Size = 11;
                Rng.Style.Font.Bold = true;
                Rng.Style.Font.Italic = true;
                Rng.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(217, 217, 217));
                Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 100, 100));
                Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }
            using (ExcelRange Rng = wsSheetRed.Cells[1, 15, 1, 15])
            {
                Rng.Value = "RFSName2";
                Rng.Style.Font.Size = 11;
                Rng.Style.Font.Bold = true;
                Rng.Style.Font.Italic = true;
                Rng.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(217, 217, 217));
                Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 100, 100));
                Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }
            using (ExcelRange Rng = wsSheetRed.Cells[1, 16, 1, 16])
            {
                Rng.Value = "RFSScore2";
                Rng.Style.Font.Size = 11;
                Rng.Style.Font.Bold = true;
                Rng.Style.Font.Italic = true;
                Rng.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(217, 217, 217));
                Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 100, 100));
                Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }
            //using (ExcelRange Rng = wsSheetRed.Cells[1, 16, 1, 16])
            //{
            //    Rng.Value = "RFSThreshold";
            //    Rng.Style.Font.Size = 11;
            //    Rng.Style.Font.Bold = true;
            //    Rng.Style.Font.Italic = true;
            //    Rng.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            //    //Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(217, 217, 217));
            //    Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 100, 100));
            //    Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            //}

            // Red Focused Unsupervised
            using (ExcelRange Rng = wsSheetRed.Cells[1, 17, 1, 17])
            {
                Rng.Value = "RFUDetect";
                Rng.Style.Font.Size = 11;
                Rng.Style.Font.Bold = true;
                Rng.Style.Font.Italic = true;
                Rng.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(217, 217, 217));
                Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 160, 160));
                Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }
            using (ExcelRange Rng = wsSheetRed.Cells[1, 18, 1, 18])
            {
                Rng.Value = "RFUName1";
                Rng.Style.Font.Size = 11;
                Rng.Style.Font.Bold = true;
                Rng.Style.Font.Italic = true;
                Rng.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(217, 217, 217));
                Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 160, 160));
                Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }
            using (ExcelRange Rng = wsSheetRed.Cells[1, 19, 1, 19])
            {
                Rng.Value = "RFUScore1";
                Rng.Style.Font.Size = 11;
                Rng.Style.Font.Bold = true;
                Rng.Style.Font.Italic = true;
                Rng.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(217, 217, 217));
                Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 160, 160));
                Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }
            using (ExcelRange Rng = wsSheetRed.Cells[1, 20, 1, 20])
            {
                Rng.Value = "RFUName2";
                Rng.Style.Font.Size = 11;
                Rng.Style.Font.Bold = true;
                Rng.Style.Font.Italic = true;
                Rng.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(217, 217, 217));
                Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 160, 160));
                Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }
            using (ExcelRange Rng = wsSheetRed.Cells[1, 21, 1, 21])
            {
                Rng.Value = "RFUScore2";
                Rng.Style.Font.Size = 11;
                Rng.Style.Font.Bold = true;
                Rng.Style.Font.Italic = true;
                Rng.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(217, 217, 217));
                Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 160, 160));
                Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }
            //using (ExcelRange Rng = wsSheetRed.Cells[1, 20, 1, 20])
            //{
            //    Rng.Value = "RFUThreshold";
            //    Rng.Style.Font.Size = 11;
            //    Rng.Style.Font.Bold = true;
            //    Rng.Style.Font.Italic = true;
            //    Rng.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            //    //Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(217, 217, 217));
            //    Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 160, 160));
            //    Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            //}

            int REndRow = (Constants.RepeatProcess + 4); // '4' means 'repeat+max+min+avrg'
            int REndColumn = 4 + (6 + 6 + 4); // Repeat, Red HDM, Red FSu, Red FUn, RHDMDetect, RHDMName1, RHDMScore1, RHDMName2, RHDMScore2, RHDMThreshold, RFSDetect, RFSName1, RFSScore, 1RFSName2, RFSScore2, RFSThreshold, RFUDetect, RFUName1, RFUScore1, RFUThreshold

            using (ExcelRange Rng = wsSheetRed.Cells[1, 1, REndRow, REndColumn])
            {
                Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }
            // JK - 2023.07.12 - End

            using (ExcelRange Rng = wsSheetRed.Cells[2, 4, 2, 4])
            {
                Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }

            wsSheetRed.Protection.IsProtected = false;
            wsSheetRed.Protection.AllowSelectLockedCells = false;
            // Red - End

            // Green - Start
            ExcelWorksheet wsSheetGreen = ExcelPkg.Workbook.Worksheets.Add("GreenTools"); // Green Tools

            using (ExcelRange Rng = wsSheetGreen.Cells[1, 1, 1, 1])
            {
                Rng.Value = "Repeat";
                Rng.Style.Font.Size = 11;
                Rng.Style.Font.Bold = true;
                Rng.Style.Font.Italic = true;
                Rng.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(217, 217, 217));
                Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }
            using (ExcelRange Rng = wsSheetGreen.Cells[2, 1, 2, 1])
            {
                Rng.Value = "Max.";
                Rng.Style.Font.Size = 11;
                Rng.Style.Font.Bold = true;
                Rng.Style.Font.Italic = true;
                Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }
            using (ExcelRange Rng = wsSheetGreen.Cells[3, 1, 3, 1])
            {
                Rng.Value = "Min.";
                Rng.Style.Font.Size = 11;
                Rng.Style.Font.Bold = true;
                Rng.Style.Font.Italic = true;
                Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }
            using (ExcelRange Rng = wsSheetGreen.Cells[4, 1, 4, 1])
            {
                Rng.Value = "Avrg.";
                Rng.Style.Font.Size = 11;
                Rng.Style.Font.Bold = true;
                Rng.Style.Font.Italic = true;
                Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }
            using (ExcelRange Rng = wsSheetGreen.Cells[1, 2, 1, 2])
            {
                Rng.Value = "Green HDM";
                Rng.Style.Font.Size = 11;
                Rng.Style.Font.Bold = true;
                Rng.Style.Font.Italic = true;
                Rng.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(217, 217, 217));
                Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }
            using (ExcelRange Rng = wsSheetGreen.Cells[1, 3, 1, 3])
            {
                Rng.Value = "Green Fcs";
                Rng.Style.Font.Size = 11;
                Rng.Style.Font.Bold = true;
                Rng.Style.Font.Italic = true;
                Rng.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(217, 217, 217));
                Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }
            using (ExcelRange Rng = wsSheetGreen.Cells[1, 4, 1, 4])
            {
                Rng.Value = "Green HDMQ";
                Rng.Style.Font.Size = 11;
                Rng.Style.Font.Bold = true;
                Rng.Style.Font.Italic = true;
                Rng.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(217, 217, 217));
                Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }

            // JK - 2023.07.12 - Start
            using (ExcelRange Rng = wsSheetGreen.Cells[1, 5, 1, 5])
            {
                Rng.Value = "GHDMBestTag";
                Rng.Style.Font.Size = 11;
                Rng.Style.Font.Bold = true;
                Rng.Style.Font.Italic = true;
                Rng.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(16, 203, 34));
                Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }
            using (ExcelRange Rng = wsSheetGreen.Cells[1, 6, 1, 6])
            {
                Rng.Value = "GHDMScore";
                Rng.Style.Font.Size = 11;
                Rng.Style.Font.Bold = true;
                Rng.Style.Font.Italic = true;
                Rng.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(16, 203, 34));
                Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }
            using (ExcelRange Rng = wsSheetGreen.Cells[1, 7, 1, 7])
            {
                Rng.Value = "GFBestTag";
                Rng.Style.Font.Size = 11;
                Rng.Style.Font.Bold = true;
                Rng.Style.Font.Italic = true;
                Rng.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(38, 238, 58));
                Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }
            using (ExcelRange Rng = wsSheetGreen.Cells[1, 8, 1, 8])
            {
                Rng.Value = "GFScore";
                Rng.Style.Font.Size = 11;
                Rng.Style.Font.Bold = true;
                Rng.Style.Font.Italic = true;
                Rng.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(38, 238, 58));
                Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }
            using (ExcelRange Rng = wsSheetGreen.Cells[1, 9, 1, 9])
            {
                Rng.Value = "GHDMQBestTag";
                Rng.Style.Font.Size = 11;
                Rng.Style.Font.Bold = true;
                Rng.Style.Font.Italic = true;
                Rng.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(128, 255, 128));
                Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }
            using (ExcelRange Rng = wsSheetGreen.Cells[1, 10, 1, 10])
            {
                Rng.Value = "GHDMQScore";
                Rng.Style.Font.Size = 11;
                Rng.Style.Font.Bold = true;
                Rng.Style.Font.Italic = true;
                Rng.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(128, 255, 128));
                Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }

            int GEndRow = (Constants.RepeatProcess + 4); // '4' means 'repeat+max+min+avrg'
            int GEndColumn = (2 * 3) + 4;
            using (ExcelRange Rng = wsSheetGreen.Cells[1, 1, GEndRow, GEndColumn])
            {
                Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }

            // JK - 2023.07.12 - End


            //using (ExcelRange Rng = wsSheetGreen.Cells[2, 4, 2, 4])
            //{
            //    Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            //}

            //wsSheetGreen.Protection.IsProtected = false;
            //wsSheetGreen.Protection.AllowSelectLockedCells = false;
            // Green - End

            // Blue Locate - Start //GetPTimesBlueLocate
            ExcelWorksheet wsSheetBlueL = ExcelPkg.Workbook.Worksheets.Add("BlueLocateTool");

            using (ExcelRange Rng = wsSheetBlueL.Cells[1, 1, 1, 1])
            {
                Rng.Value = "Repeat";
                Rng.Style.Font.Size = 11;
                Rng.Style.Font.Bold = true;
                Rng.Style.Font.Italic = true;
                Rng.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(217, 217, 217));
                Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }
            using (ExcelRange Rng = wsSheetBlueL.Cells[2, 1, 2, 1])
            {
                Rng.Value = "Max.";
                Rng.Style.Font.Size = 11;
                Rng.Style.Font.Bold = true;
                Rng.Style.Font.Italic = true;
                Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }
            using (ExcelRange Rng = wsSheetBlueL.Cells[3, 1, 3, 1])
            {
                Rng.Value = "Min.";
                Rng.Style.Font.Size = 11;
                Rng.Style.Font.Bold = true;
                Rng.Style.Font.Italic = true;
                Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }
            using (ExcelRange Rng = wsSheetBlueL.Cells[4, 1, 4, 1])
            {
                Rng.Value = "Avrg.";
                Rng.Style.Font.Size = 11;
                Rng.Style.Font.Bold = true;
                Rng.Style.Font.Italic = true;
                Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }
            using (ExcelRange Rng = wsSheetBlueL.Cells[1, 2, 1, 2])
            {
                Rng.Value = "BL_PTime";
                Rng.Style.Font.Size = 11;
                Rng.Style.Font.Bold = true;
                Rng.Style.Font.Italic = true;
                Rng.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(217, 217, 217));
                Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }

            //JK-AddResult-Start- 2023.07.06

            //using (ExcelRange Rng = wsSheetBlueL.Cells[1, 3, 1, 3])
            //{
            //    Rng.Value = "Features";
            //    Rng.Style.Font.Size = 11;
            //    Rng.Style.Font.Bold = true;
            //    Rng.Style.Font.Italic = true;
            //    Rng.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            //    Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(217, 217, 217));
            //    Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            //}
            // Total items - Blue Locate result
            // Odd Number : Tail
            using (ExcelRange Rng = wsSheetBlueL.Cells[1, 3, 1, 3])
            {
                Rng.Value = "Features";
                Rng.Style.Font.Size = 11;
                Rng.Style.Font.Bold = true;
                Rng.Style.Font.Italic = true;
                Rng.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(217, 217, 217));
                //Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(60, 120, 200));
                Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0, 125, 220));
                Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                Rng.Style.Font.Color.SetColor(Color.FromArgb(255, 255, 255));
            }
            using (ExcelRange Rng = wsSheetBlueL.Cells[1, 4, 1, 4])
            {
                Rng.Value = "FName1";
                Rng.Style.Font.Size = 11;
                Rng.Style.Font.Bold = true;
                Rng.Style.Font.Italic = true;
                Rng.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(217, 217, 217));
                Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0, 145, 255));
                Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                Rng.Style.Font.Color.SetColor(Color.FromArgb(255, 255, 255));
            }
            using (ExcelRange Rng = wsSheetBlueL.Cells[1, 5, 1, 5])
            {
                Rng.Value = "FScore1";
                Rng.Style.Font.Size = 11;
                Rng.Style.Font.Bold = true;
                Rng.Style.Font.Italic = true;
                Rng.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(217, 217, 217));
                Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0, 145, 255));
                Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                Rng.Style.Font.Color.SetColor(Color.FromArgb(255, 255, 255));
            }
            using (ExcelRange Rng = wsSheetBlueL.Cells[1, 6, 1, 6])
            {
                Rng.Value = "FPosX1";
                Rng.Style.Font.Size = 11;
                Rng.Style.Font.Bold = true;
                Rng.Style.Font.Italic = true;
                Rng.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(217, 217, 217));
                Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0, 145, 255));
                Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                Rng.Style.Font.Color.SetColor(Color.FromArgb(255, 255, 255));
            }
            using (ExcelRange Rng = wsSheetBlueL.Cells[1, 7, 1, 7])
            {
                Rng.Value = "FPosY1";
                Rng.Style.Font.Size = 11;
                Rng.Style.Font.Bold = true;
                Rng.Style.Font.Italic = true;
                Rng.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(217, 217, 217));
                Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0, 145, 255));
                Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                Rng.Style.Font.Color.SetColor(Color.FromArgb(255, 255, 255));
            }
            using (ExcelRange Rng = wsSheetBlueL.Cells[1, 8, 1, 8])
            {
                Rng.Value = "FAngle1";
                Rng.Style.Font.Size = 11;
                Rng.Style.Font.Bold = true;
                Rng.Style.Font.Italic = true;
                Rng.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(217, 217, 217));
                Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0, 145, 255));
                Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                Rng.Style.Font.Color.SetColor(Color.FromArgb(255, 255, 255));
            }
            using (ExcelRange Rng = wsSheetBlueL.Cells[1, 9, 1, 9])
            {
                Rng.Value = "FSizeH1";
                Rng.Style.Font.Size = 11;
                Rng.Style.Font.Bold = true;
                Rng.Style.Font.Italic = true;
                Rng.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(217, 217, 217));
                Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0, 145, 255));
                Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                Rng.Style.Font.Color.SetColor(Color.FromArgb(255, 255, 255));
            }
            using (ExcelRange Rng = wsSheetBlueL.Cells[1, 10, 1, 10])
            {
                Rng.Value = "FSizeW1";
                Rng.Style.Font.Size = 11;
                Rng.Style.Font.Bold = true;
                Rng.Style.Font.Italic = true;
                Rng.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(217, 217, 217));
                Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0, 145, 255));
                Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                Rng.Style.Font.Color.SetColor(Color.FromArgb(255, 255, 255));
            }

            // Even Number : Head
            using (ExcelRange Rng = wsSheetBlueL.Cells[1, 11, 1, 11])
            {
                Rng.Value = "FName2";
                Rng.Style.Font.Size = 11;
                Rng.Style.Font.Bold = true;
                Rng.Style.Font.Italic = true;
                Rng.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(217, 217, 217));
                Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(60, 170, 220));
                Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                Rng.Style.Font.Color.SetColor(Color.FromArgb(255, 255, 255));
            }
            using (ExcelRange Rng = wsSheetBlueL.Cells[1, 12, 1, 12])
            {
                Rng.Value = "FScore2";
                Rng.Style.Font.Size = 11;
                Rng.Style.Font.Bold = true;
                Rng.Style.Font.Italic = true;
                Rng.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(217, 217, 217));
                Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(60, 170, 220));
                Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                Rng.Style.Font.Color.SetColor(Color.FromArgb(255, 255, 255));
            }
            using (ExcelRange Rng = wsSheetBlueL.Cells[1, 13, 1, 13])
            {
                Rng.Value = "FPosX2";
                Rng.Style.Font.Size = 11;
                Rng.Style.Font.Bold = true;
                Rng.Style.Font.Italic = true;
                Rng.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(217, 217, 217));
                Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(60, 170, 220));
                Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                Rng.Style.Font.Color.SetColor(Color.FromArgb(255, 255, 255));
            }
            using (ExcelRange Rng = wsSheetBlueL.Cells[1, 14, 1, 14])
            {
                Rng.Value = "FPosY2";
                Rng.Style.Font.Size = 11;
                Rng.Style.Font.Bold = true;
                Rng.Style.Font.Italic = true;
                Rng.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(217, 217, 217));
                Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(60, 170, 220));
                Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                Rng.Style.Font.Color.SetColor(Color.FromArgb(255, 255, 255));
            }
            using (ExcelRange Rng = wsSheetBlueL.Cells[1, 15, 1, 15])
            {
                Rng.Value = "FAngle2";
                Rng.Style.Font.Size = 11;
                Rng.Style.Font.Bold = true;
                Rng.Style.Font.Italic = true;
                Rng.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(217, 217, 217));
                Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(60, 170, 220));
                Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                Rng.Style.Font.Color.SetColor(Color.FromArgb(255, 255, 255));
            }
            using (ExcelRange Rng = wsSheetBlueL.Cells[1, 16, 1, 16])
            {
                Rng.Value = "FSizeH2";
                Rng.Style.Font.Size = 11;
                Rng.Style.Font.Bold = true;
                Rng.Style.Font.Italic = true;
                Rng.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(217, 217, 217));
                Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(60, 170, 220));
                Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                Rng.Style.Font.Color.SetColor(Color.FromArgb(255, 255, 255));
            }
            using (ExcelRange Rng = wsSheetBlueL.Cells[1, 17, 1, 17])
            {
                Rng.Value = "FSizeW2";
                Rng.Style.Font.Size = 11;
                Rng.Style.Font.Bold = true;
                Rng.Style.Font.Italic = true;
                Rng.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(217, 217, 217));
                Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(60, 170, 220));
                Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                Rng.Style.Font.Color.SetColor(Color.FromArgb(255, 255, 255));
            }



            int BLEndRow = (Constants.RepeatProcess + 4); // '4' means 'repeat+max+min+avrg'
            int BLEndColumn = (7 * 2) + 3;
            using (ExcelRange Rng = wsSheetBlueL.Cells[1, 1, BLEndRow, BLEndColumn])
            {
                Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }

            //JK-AddResult-End - 2023.07.06

            wsSheetBlueL.Protection.IsProtected = false;
            wsSheetBlueL.Protection.AllowSelectLockedCells = false;
            // Blue Locate - End


            // Blue Read - Start
            ExcelWorksheet wsSheetBlueR = ExcelPkg.Workbook.Worksheets.Add("BlueReadTool");

            using (ExcelRange Rng = wsSheetBlueR.Cells[1, 1, 1, 1])
            {
                Rng.Value = "Repeat";
                Rng.Style.Font.Size = 11;
                Rng.Style.Font.Bold = true;
                Rng.Style.Font.Italic = true;
                Rng.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(217, 217, 217));
                Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }
            using (ExcelRange Rng = wsSheetBlueR.Cells[2, 1, 2, 1])
            {
                Rng.Value = "Max.";
                Rng.Style.Font.Size = 11;
                Rng.Style.Font.Bold = true;
                Rng.Style.Font.Italic = true;
                Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }
            using (ExcelRange Rng = wsSheetBlueR.Cells[3, 1, 3, 1])
            {
                Rng.Value = "Min.";
                Rng.Style.Font.Size = 11;
                Rng.Style.Font.Bold = true;
                Rng.Style.Font.Italic = true;
                Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }
            using (ExcelRange Rng = wsSheetBlueR.Cells[4, 1, 4, 1])
            {
                Rng.Value = "Avrg.";
                Rng.Style.Font.Size = 11;
                Rng.Style.Font.Bold = true;
                Rng.Style.Font.Italic = true;
                Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }
            using (ExcelRange Rng = wsSheetBlueR.Cells[1, 2, 1, 2])
            {
                Rng.Value = "Blue Read";
                Rng.Style.Font.Size = 11;
                Rng.Style.Font.Bold = true;
                Rng.Style.Font.Italic = true;
                Rng.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(217, 217, 217));
                Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }
            wsSheetBlueR.Protection.IsProtected = false;
            wsSheetBlueR.Protection.AllowSelectLockedCells = false;

            //JK-AddResult-Start - 2023.07.11
            using (ExcelRange Rng = wsSheetBlueR.Cells[1, 3, 1, 3])
            {
                Rng.Value = "Features";
                Rng.Style.Font.Size = 11;
                Rng.Style.Font.Bold = true;
                Rng.Style.Font.Italic = true;
                Rng.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(217, 217, 217)); // Color is gray
                Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0, 75, 163));
                Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                Rng.Style.Font.Color.SetColor(Color.FromArgb(255, 255, 255));

            }

            using (ExcelRange Rng = wsSheetBlueR.Cells[1, 4, 1, 4])
            {
                Rng.Value = "FName1";
                Rng.Style.Font.Size = 11;
                Rng.Style.Font.Bold = true;
                Rng.Style.Font.Italic = true;
                Rng.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(217, 217, 217));
                Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0, 95, 210));
                Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                Rng.Style.Font.Color.SetColor(Color.FromArgb(255, 255, 255));
            }
            using (ExcelRange Rng = wsSheetBlueR.Cells[1, 5, 1, 5])
            {
                Rng.Value = "FScore1";
                Rng.Style.Font.Size = 11;
                Rng.Style.Font.Bold = true;
                Rng.Style.Font.Italic = true;
                Rng.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0, 95, 210));
                Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                Rng.Style.Font.Color.SetColor(Color.FromArgb(255, 255, 255));
            }

            using (ExcelRange Rng = wsSheetBlueR.Cells[1, 6, 1, 6])
            {
                Rng.Value = "FName2";
                Rng.Style.Font.Size = 11;
                Rng.Style.Font.Bold = true;
                Rng.Style.Font.Italic = true;
                Rng.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0, 95, 210));
                Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                Rng.Style.Font.Color.SetColor(Color.FromArgb(255, 255, 255));
            }
            using (ExcelRange Rng = wsSheetBlueR.Cells[1, 7, 1, 7])
            {
                Rng.Value = "FScore2";
                Rng.Style.Font.Size = 11;
                Rng.Style.Font.Bold = true;
                Rng.Style.Font.Italic = true;
                Rng.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0, 95, 210));
                Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                Rng.Style.Font.Color.SetColor(Color.FromArgb(255, 255, 255));
            }

            using (ExcelRange Rng = wsSheetBlueR.Cells[1, 8, 1, 8])
            {
                Rng.Value = "FName3";
                Rng.Style.Font.Size = 11;
                Rng.Style.Font.Bold = true;
                Rng.Style.Font.Italic = true;
                Rng.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0, 95, 210));
                Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                Rng.Style.Font.Color.SetColor(Color.FromArgb(255, 255, 255));
            }
            using (ExcelRange Rng = wsSheetBlueR.Cells[1, 9, 1, 9])
            {
                Rng.Value = "FScore3";
                Rng.Style.Font.Size = 11;
                Rng.Style.Font.Bold = true;
                Rng.Style.Font.Italic = true;
                Rng.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0, 95, 210));
                Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                Rng.Style.Font.Color.SetColor(Color.FromArgb(255, 255, 255));
            }

            using (ExcelRange Rng = wsSheetBlueR.Cells[1, 10, 1, 10])
            {
                Rng.Value = "FName4";
                Rng.Style.Font.Size = 11;
                Rng.Style.Font.Bold = true;
                Rng.Style.Font.Italic = true;
                Rng.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0, 95, 210));
                Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                Rng.Style.Font.Color.SetColor(Color.FromArgb(255, 255, 255));
            }
            using (ExcelRange Rng = wsSheetBlueR.Cells[1, 11, 1, 11])
            {
                Rng.Value = "FScore4";
                Rng.Style.Font.Size = 11;
                Rng.Style.Font.Bold = true;
                Rng.Style.Font.Italic = true;
                Rng.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0, 95, 210));
                Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                Rng.Style.Font.Color.SetColor(Color.FromArgb(255, 255, 255));
            }

            using (ExcelRange Rng = wsSheetBlueR.Cells[1, 12, 1, 12])
            {
                Rng.Value = "FName5";
                Rng.Style.Font.Size = 11;
                Rng.Style.Font.Bold = true;
                Rng.Style.Font.Italic = true;
                Rng.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0, 95, 210));
                Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                Rng.Style.Font.Color.SetColor(Color.FromArgb(255, 255, 255));
            }
            using (ExcelRange Rng = wsSheetBlueR.Cells[1, 13, 1, 13])
            {
                Rng.Value = "FScore5";
                Rng.Style.Font.Size = 11;
                Rng.Style.Font.Bold = true;
                Rng.Style.Font.Italic = true;
                Rng.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0, 95, 210));
                Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                Rng.Style.Font.Color.SetColor(Color.FromArgb(255, 255, 255));
            }

            using (ExcelRange Rng = wsSheetBlueR.Cells[1, 14, 1, 14])
            {
                Rng.Value = "FName6";
                Rng.Style.Font.Size = 11;
                Rng.Style.Font.Bold = true;
                Rng.Style.Font.Italic = true;
                Rng.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0, 95, 210));
                Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                Rng.Style.Font.Color.SetColor(Color.FromArgb(255, 255, 255));
            }
            using (ExcelRange Rng = wsSheetBlueR.Cells[1, 15, 1, 15])
            {
                Rng.Value = "FScore6";
                Rng.Style.Font.Size = 11;
                Rng.Style.Font.Bold = true;
                Rng.Style.Font.Italic = true;
                Rng.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0, 95, 210));
                Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                Rng.Style.Font.Color.SetColor(Color.FromArgb(255, 255, 255));
            }

            using (ExcelRange Rng = wsSheetBlueR.Cells[1, 16, 1, 16])
            {
                Rng.Value = "FName7";
                Rng.Style.Font.Size = 11;
                Rng.Style.Font.Bold = true;
                Rng.Style.Font.Italic = true;
                Rng.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0, 95, 210));
                Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                Rng.Style.Font.Color.SetColor(Color.FromArgb(255, 255, 255));
            }
            using (ExcelRange Rng = wsSheetBlueR.Cells[1, 17, 1, 17])
            {
                Rng.Value = "FScore7";
                Rng.Style.Font.Size = 11;
                Rng.Style.Font.Bold = true;
                Rng.Style.Font.Italic = true;
                Rng.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0, 95, 210));
                Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                Rng.Style.Font.Color.SetColor(Color.FromArgb(255, 255, 255));
            }

            using (ExcelRange Rng = wsSheetBlueR.Cells[1, 18, 1, 18])
            {
                Rng.Value = "FName8";
                Rng.Style.Font.Size = 11;
                Rng.Style.Font.Bold = true;
                Rng.Style.Font.Italic = true;
                Rng.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0, 95, 210));
                Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                Rng.Style.Font.Color.SetColor(Color.FromArgb(255, 255, 255));
            }
            using (ExcelRange Rng = wsSheetBlueR.Cells[1, 19, 1, 19])
            {
                Rng.Value = "FScore8";
                Rng.Style.Font.Size = 11;
                Rng.Style.Font.Bold = true;
                Rng.Style.Font.Italic = true;
                Rng.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0, 95, 210));
                Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                Rng.Style.Font.Color.SetColor(Color.FromArgb(255, 255, 255));
            }

            using (ExcelRange Rng = wsSheetBlueR.Cells[1, 20, 1, 20])
            {
                Rng.Value = "FName9";
                Rng.Style.Font.Size = 11;
                Rng.Style.Font.Bold = true;
                Rng.Style.Font.Italic = true;
                Rng.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0, 95, 210));
                Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                Rng.Style.Font.Color.SetColor(Color.FromArgb(255, 255, 255));
            }
            using (ExcelRange Rng = wsSheetBlueR.Cells[1, 21, 1, 21])
            {
                Rng.Value = "FScore9";
                Rng.Style.Font.Size = 11;
                Rng.Style.Font.Bold = true;
                Rng.Style.Font.Italic = true;
                Rng.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                Rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0, 95, 210));
                Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                Rng.Style.Font.Color.SetColor(Color.FromArgb(255, 255, 255));
            }

            int BRcountFeatures = int.Parse(GetBlueReadMatchCountFeatures.ElementAt(0)); //GetBlueReadMatchCountFeatures 값은 Read 처리시 검출된 feature 수.
            int BREndRow = (Constants.RepeatProcess + 4); // '4' means 'repeat+max+min+avrg'
            int BREndColumn = (BRcountFeatures * 2) + 3;
            using (ExcelRange Rng = wsSheetBlueR.Cells[1, 1, BREndRow, BREndColumn])
            {
                Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }

            //JK-AddResult-End - 2023.07.11
            // Blue Read - End

            ExcelPkg.SaveAs(new FileInfo(@savePath));
            Console.WriteLine(" - Complete the creating excel file!");

            Console.WriteLine("JK Test 2. Adding Chart after loading the created excel.");
            string pathExcelFile = savePath;
            Console.WriteLine(" - Load ExcelInfo: {0}", pathExcelFile);

            FileInfo existingFile = new FileInfo(pathExcelFile);
            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                int startCellindex = 5; // in case of adding max, min, average value If you change value to '2' from '5', The process time insert to next cell from cell(A2) as like first test.

                // *** RedTools Chart - Start
                // Create RedTools sheet
                ExcelWorksheet worksheetRedTools = package.Workbook.Worksheets["RedTools"];
                // Fill in index number and process time regarding each red tools.
                int columnRedTools = 1;
                for (int row = startCellindex; row < (Constants.RepeatProcess + startCellindex); row++)
                    worksheetRedTools.Cells[row, columnRedTools].Value = row - (startCellindex - 1);

                int colRedTools = 2;    // Red HDM
                for (int row = startCellindex; row < (Constants.RepeatProcess + startCellindex); row++)
                    worksheetRedTools.Cells[row, colRedTools].Value = int.Parse(GetPTimesRedHDM[row - startCellindex]);

                colRedTools = 3;        // Red Focused Supervised
                for (int row = startCellindex; row < (Constants.RepeatProcess + startCellindex); row++)
                    worksheetRedTools.Cells[row, colRedTools].Value = int.Parse(GetPTimesRedFSu[row - startCellindex]);

                colRedTools = 4;        // Red Focused Unsupervised
                for (int row = startCellindex; row < (Constants.RepeatProcess + startCellindex); row++)
                    worksheetRedTools.Cells[row, colRedTools].Value = int.Parse(GetPTimesRedFUn[row - startCellindex]);

                // JK-AddResultOFRed-2023.07.12- Start // Red HDM, Focused Supervised, Focused Unsupervised.
                // Red HDM
                colRedTools = 5;        // Red HDM - Defect regions
                for (int row = startCellindex; row < (Constants.RepeatProcess + startCellindex); row++)
                    worksheetRedTools.Cells[row, colRedTools].Value = int.Parse(GetRedHDMDetectedRegions[row - startCellindex]);
                colRedTools = 6;        // Red HDM - Defect Name1
                for (int row = startCellindex; row < (Constants.RepeatProcess + startCellindex); row++)
                    worksheetRedTools.Cells[row, colRedTools].Value = GetRedHDMRegionResult[(((row - startCellindex) * 3) + 0)].Name;
                colRedTools = 7;        // Red HDM - Defect Score1
                for (int row = startCellindex; row < (Constants.RepeatProcess + startCellindex); row++)
                    worksheetRedTools.Cells[row, colRedTools].Value = GetRedHDMRegionResult[(((row - startCellindex) * 3) + 0)].Score;
                colRedTools = 8;        // Red HDM - Defect Name2
                for (int row = startCellindex; row < (Constants.RepeatProcess + startCellindex); row++)
                    worksheetRedTools.Cells[row, colRedTools].Value = GetRedHDMRegionResult[(((row - startCellindex) * 3) + 1)].Name;
                colRedTools = 9;        // Red HDM - Defect Score2
                for (int row = startCellindex; row < (Constants.RepeatProcess + startCellindex); row++)
                    worksheetRedTools.Cells[row, colRedTools].Value = GetRedHDMRegionResult[(((row - startCellindex) * 3) + 1)].Score;

                // JK-ModifyCodeRedHDM-2023.07.13- Start
                colRedTools = 10;        // Red HDM - Defect Name3
                for (int row = startCellindex; row < (Constants.RepeatProcess + startCellindex); row++)
                    worksheetRedTools.Cells[row, colRedTools].Value = GetRedHDMRegionResult[(((row - startCellindex) * 3) + 2)].Name;
                colRedTools = 11;        // Red HDM - Defect Score3
                for (int row = startCellindex; row < (Constants.RepeatProcess + startCellindex); row++)
                    worksheetRedTools.Cells[row, colRedTools].Value = GetRedHDMRegionResult[(((row - startCellindex) * 3) + 2)].Score;

                // JK-ModifyCodeRedHDM-2023.07.13- End

                //colRedTools = 10;        // Red HDM - Threshold
                //for (int row = startCellindex; row < (Constants.RepeatProcess + startCellindex); row++)
                //    //worksheetRedTools.Cells[row, colRedTools].Value = GetRedHDMRegionResult[(((row - startCellindex) * 2) + 1)].Score;

                // Red Focused Supervised
                colRedTools = 12;        // Red Focused Supervised - Defect regions
                for (int row = startCellindex; row < (Constants.RepeatProcess + startCellindex); row++)
                    worksheetRedTools.Cells[row, colRedTools].Value = int.Parse(GetRedFocusedSupervisedDetectedRegions[row - startCellindex]);
                colRedTools = 13;        // Red Focused Supervised - Defect Name1
                for (int row = startCellindex; row < (Constants.RepeatProcess + startCellindex); row++)
                    worksheetRedTools.Cells[row, colRedTools].Value = GetRedFocusedSupervisedRegionResult[(((row - startCellindex) * 2) + 0)].Name;
                colRedTools = 14;        // Red Focused Supervised - Defect Score1
                for (int row = startCellindex; row < (Constants.RepeatProcess + startCellindex); row++)
                    worksheetRedTools.Cells[row, colRedTools].Value = GetRedFocusedSupervisedRegionResult[(((row - startCellindex) * 2) + 0)].Score;
                colRedTools = 15;        // Red Focused Supervised - Defect Name2
                for (int row = startCellindex; row < (Constants.RepeatProcess + startCellindex); row++)
                    worksheetRedTools.Cells[row, colRedTools].Value = GetRedFocusedSupervisedRegionResult[(((row - startCellindex) * 2) + 1)].Name;
                colRedTools = 16;        // Red Focused Supervised - Defect Score2
                for (int row = startCellindex; row < (Constants.RepeatProcess + startCellindex); row++)
                    worksheetRedTools.Cells[row, colRedTools].Value = GetRedFocusedSupervisedRegionResult[(((row - startCellindex) * 2) + 1)].Score;
                //colRedTools = 15;        // Red Focused Supervised - Threashold
                //for (int row = startCellindex; row < (Constants.RepeatProcess + startCellindex); row++)
                //    //worksheetRedTools.Cells[row, colRedTools].Value = GetRedFocusedSupervisedRegionResult[(((row - startCellindex) * 2) + 1)].Score;

                // Red Focused Unsupervised
                colRedTools = 17;        // Red Focused Unsupervised - Defect regions
                for (int row = startCellindex; row < (Constants.RepeatProcess + startCellindex); row++)
                    worksheetRedTools.Cells[row, colRedTools].Value = int.Parse(GetRedFocusedUnsupervisedDetectedRegions[row - startCellindex]);
                colRedTools = 18;        // Red Focused Unsupervised - Defect Name1
                for (int row = startCellindex; row < (Constants.RepeatProcess + startCellindex); row++)
                    worksheetRedTools.Cells[row, colRedTools].Value = GetRedFocusedUnsupervisedRegionResult[(((row - startCellindex) * 2) + 0)].Name;
                colRedTools = 19;        // Red Focused Unsupervised - Defect Score1
                for (int row = startCellindex; row < (Constants.RepeatProcess + startCellindex); row++)
                    worksheetRedTools.Cells[row, colRedTools].Value = GetRedFocusedUnsupervisedRegionResult[(((row - startCellindex) * 2) + 0)].Score;
                colRedTools = 20;        // Red Focused Unsupervised - Defect Name2
                for (int row = startCellindex; row < (Constants.RepeatProcess + startCellindex); row++)
                    worksheetRedTools.Cells[row, colRedTools].Value = GetRedFocusedUnsupervisedRegionResult[(((row - startCellindex) * 2) + 1)].Name;
                colRedTools = 21;        // Red Focused Unsupervised - Defect Score2
                for (int row = startCellindex; row < (Constants.RepeatProcess + startCellindex); row++)
                    worksheetRedTools.Cells[row, colRedTools].Value = GetRedFocusedUnsupervisedRegionResult[(((row - startCellindex) * 2) + 1)].Score;
                //colRedTools = 20;        // Red Focused Unsupervised - Threashold
                //for (int row = startCellindex; row < (Constants.RepeatProcess + startCellindex); row++)
                //    //worksheetRedTools.Cells[row, colRedTools].Value = GetRedFocusedSupervisedRegionResult[(((row - startCellindex) * 2) + 1)].Score;

                // JK-AddResultOFRed-2023.07.12- End // Red HDM, Focused Supervised, Focused Unsupervised.


                // Fill in max, min , average for analysing process time each red tools.                
                worksheetRedTools.Cells["B2"].Formula = $"MAX(B{startCellindex}:B{(Constants.RepeatProcess + startCellindex - 1)})";        // Red HDM : maximum
                worksheetRedTools.Cells["B3"].Formula = $"MIN(B{startCellindex}:B{(Constants.RepeatProcess + startCellindex - 1)})";        // Red HDM : minimum                
                worksheetRedTools.Cells["B4"].Formula = $"AVERAGE(B{startCellindex}:B{(Constants.RepeatProcess + startCellindex - 1)})";    // Red HDM : Average
                worksheetRedTools.Cells["C2"].Formula = $"MAX(C{startCellindex}:C{(Constants.RepeatProcess + startCellindex - 1)})";        // Red Focused Supervised : maximum                
                worksheetRedTools.Cells["C3"].Formula = $"MIN(C{startCellindex}:C{(Constants.RepeatProcess + startCellindex - 1)})";        // Red Focused Supervised : minimum                
                worksheetRedTools.Cells["C4"].Formula = $"AVERAGE(C{startCellindex}:C{(Constants.RepeatProcess + startCellindex - 1)})";    // Red Focused Supervised : Average
                worksheetRedTools.Cells["D2"].Formula = $"MAX(D{startCellindex}:D{(Constants.RepeatProcess + startCellindex - 1)})";          // Red Focused Unsupervised : maximum                
                worksheetRedTools.Cells["D3"].Formula = $"MIN(D{startCellindex}:D{(Constants.RepeatProcess + startCellindex - 1)})";          // Red Focused Unsupervised : minimum                
                worksheetRedTools.Cells["D4"].Formula = $"AVERAGE(D{startCellindex}:D{(Constants.RepeatProcess + startCellindex - 1)})";      // Red Focused Unsupervised : Average
                // Adding chart for the visibility of analysing data.
                var chartRedTools = worksheetRedTools.Drawings.AddChart("Chart_Red", eChartType.Line);
                chartRedTools.Title.Text = "Processing Time Red Tool(HDM/FSu/FUn)[ms]";
                chartRedTools.Title.Font.Size = 14; //chartRedTools.Title.Font.Color = Color.FromArgb(238, 46, 34);
                chartRedTools.Title.Font.Bold = true;
                chartRedTools.Title.Font.Italic = true;
                chartRedTools.SetPosition(7, 7, 6, 6); // Start point to dispale of Chart  ex) 0,0,5,5 : Draw a chart from F1 Cell vs 1,1,6,6 : Draw a chart from G2 Cell
                chartRedTools.SetSize(800, 600);

                ExcelAddress valueAddress_Data1_RedTools = new ExcelAddress(startCellindex, 2, (Constants.RepeatProcess + (startCellindex - 1)), 2);
                ExcelAddress RepeatAddress_Data1_RedTools = new ExcelAddress(startCellindex, 1, (Constants.RepeatProcess + (startCellindex - 1)), 1);
                var ser1_RedTools = (chartRedTools.Series.Add(valueAddress_Data1_RedTools.Address, RepeatAddress_Data1_RedTools.Address) as ExcelLineChartSerie);
                ser1_RedTools.Header = "Red HDM";

                ExcelAddress valueAddress_Data2_RedTools = new ExcelAddress(startCellindex, 3, (Constants.RepeatProcess + (startCellindex - 1)), 3);
                ExcelAddress RepeatAddress_Data2_RedTools = new ExcelAddress(startCellindex, 1, (Constants.RepeatProcess + (startCellindex - 1)), 1);
                var ser2_RedTools = (chartRedTools.Series.Add(valueAddress_Data2_RedTools.Address, RepeatAddress_Data2_RedTools.Address) as ExcelLineChartSerie);
                ser2_RedTools.Header = "Red FSu";

                ExcelAddress valueAddress_Data3_RedTools = new ExcelAddress(startCellindex, 4, (Constants.RepeatProcess + (startCellindex - 1)), 4);
                ExcelAddress RepeatAddress_Data3_RedTools = new ExcelAddress(startCellindex, 1, (Constants.RepeatProcess + (startCellindex - 1)), 1);
                var ser3_RedTools = (chartRedTools.Series.Add(valueAddress_Data3_RedTools.Address, RepeatAddress_Data3_RedTools.Address) as ExcelLineChartSerie);
                ser3_RedTools.Header = "Red FUn";

                chartRedTools.Legend.Border.LineStyle = eLineStyle.Solid;
                chartRedTools.Legend.Border.Fill.Style = eFillStyle.SolidFill;
                chartRedTools.Legend.Border.Fill.Color = Color.DarkRed;
                //chartRedTools.Border.Width = 1;
                chartRedTools.Border.Fill.Color = Color.DarkRed; // Color.FromArgb(238, 46, 34);
                                                                 // *** RedTools Chart - End               

                // *** GreenTools Chart - Start
                // Create GreenTools sheet
                ExcelWorksheet worksheetGreenTools = package.Workbook.Worksheets["GreenTools"];
                // Fill in index number and process time regarding each green tools.
                int columnGreenTools = 1;
                for (int row = startCellindex; row < (Constants.RepeatProcess + startCellindex); row++)
                    worksheetGreenTools.Cells[row, columnGreenTools].Value = row - (startCellindex - 1);

                int colGreenTools = 2;    // Green HDM
                for (int row = startCellindex; row < (Constants.RepeatProcess + startCellindex); row++)
                    worksheetGreenTools.Cells[row, colGreenTools].Value = int.Parse(GetPTimesGreenHDM[row - startCellindex]);

                colGreenTools = 3;        // Green Focused
                for (int row = startCellindex; row < (Constants.RepeatProcess + startCellindex); row++)
                    worksheetGreenTools.Cells[row, colGreenTools].Value = int.Parse(GetPTimesGreenFocused[row - startCellindex]);

                colGreenTools = 4;        // Green HDM Quick
                for (int row = startCellindex; row < (Constants.RepeatProcess + startCellindex); row++)
                    worksheetGreenTools.Cells[row, colGreenTools].Value = int.Parse(GetPTimesGreenHDMQuick[row - startCellindex]);

                // JK-AddResultOFGreen-2023.07.12- Start // Green HDM, Focused, HDM Quick
                // Green HDM
                colGreenTools = 5;        // Green HDM - BestTag.Name
                for (int row = startCellindex; row < (Constants.RepeatProcess + startCellindex); row++)
                    worksheetGreenTools.Cells[row, colGreenTools].Value = GetGreenHDMMatchAndViewResult[(row - startCellindex)].BestTagName;
                colGreenTools = 6;        // Green HDM - BestTag.Score
                for (int row = startCellindex; row < (Constants.RepeatProcess + startCellindex); row++)
                    worksheetGreenTools.Cells[row, colGreenTools].Value = GetGreenHDMMatchAndViewResult[(row - startCellindex)].BestTagScore;
                //colGreenTools = 7;        // Green HDM - Threshold
                //for (int row = startCellindex; row < (Constants.RepeatProcess + startCellindex); row++)                    
                //    worksheetGreenTools.Cells[row, colGreenTools].Value = GetGreenHDMMatchAndViewResult[(row - startCellindex)].Threshold;
                //colGreenTools = 8;        // Green HDM - Size.Height
                //for (int row = startCellindex; row < (Constants.RepeatProcess + startCellindex); row++)                    
                //    worksheetGreenTools.Cells[row, colGreenTools].Value = GetGreenHDMMatchAndViewResult[(row - startCellindex)].SizeHeight;
                //colGreenTools = 9;        // Green HDM - Size.Width
                //for (int row = startCellindex; row < (Constants.RepeatProcess + startCellindex); row++)
                //    worksheetGreenTools.Cells[row, colGreenTools].Value = GetGreenHDMMatchAndViewResult[(row - startCellindex)].SizeWidth;

                colGreenTools = 7;        // Green Focused - BestTag.Name
                for (int row = startCellindex; row < (Constants.RepeatProcess + startCellindex); row++)
                    worksheetGreenTools.Cells[row, colGreenTools].Value = GetGreenFocusedMatchAndViewResult[(row - startCellindex)].BestTagName;
                colGreenTools = 8;        // Green Focused - BestTag.Score
                for (int row = startCellindex; row < (Constants.RepeatProcess + startCellindex); row++)
                    worksheetGreenTools.Cells[row, colGreenTools].Value = GetGreenFocusedMatchAndViewResult[(row - startCellindex)].BestTagScore;

                colGreenTools = 9;        // Green HDM Quick - BestTag.Name
                for (int row = startCellindex; row < (Constants.RepeatProcess + startCellindex); row++)
                    worksheetGreenTools.Cells[row, colGreenTools].Value = GetGreenHDMQuickMatchAndViewResult[(row - startCellindex)].BestTagName;
                colGreenTools = 10;        // Green HDM Quick - BestTag.Score
                for (int row = startCellindex; row < (Constants.RepeatProcess + startCellindex); row++)
                    worksheetGreenTools.Cells[row, colGreenTools].Value = GetGreenHDMQuickMatchAndViewResult[(row - startCellindex)].BestTagScore;
                // JK-AddResultOFGreen-2023.07.12- End // Green HDM, Focused, HDM Quick

                // Fill in max, min, average for analysing process time each green tools.
                worksheetGreenTools.Cells["B2"].Formula = $"MAX(B{startCellindex}:B{(Constants.RepeatProcess + startCellindex - 1)})";
                worksheetGreenTools.Cells["B3"].Formula = $"MIN(B{startCellindex}:B{(Constants.RepeatProcess + startCellindex - 1)})";
                worksheetGreenTools.Cells["B4"].Formula = $"AVERAGE(B{startCellindex}:B{(Constants.RepeatProcess + startCellindex - 1)})";
                worksheetGreenTools.Cells["C2"].Formula = $"MAX(C{startCellindex}:C{(Constants.RepeatProcess + startCellindex - 1)})";
                worksheetGreenTools.Cells["C3"].Formula = $"MIN(C{startCellindex}:C{(Constants.RepeatProcess + startCellindex - 1)})";
                worksheetGreenTools.Cells["C4"].Formula = $"AVERAGE(C{startCellindex}:C{(Constants.RepeatProcess + startCellindex - 1)})";
                worksheetGreenTools.Cells["D2"].Formula = $"MAX(D{startCellindex}:D{(Constants.RepeatProcess + startCellindex - 1)})";
                worksheetGreenTools.Cells["D3"].Formula = $"MIN(D{startCellindex}:D{(Constants.RepeatProcess + startCellindex - 1)})";
                worksheetGreenTools.Cells["D4"].Formula = $"AVERAGE(D{startCellindex}:D{(Constants.RepeatProcess + startCellindex - 1)})";
                // Adding chart for the visibility of analysing data.
                var chartGreenTools = worksheetGreenTools.Drawings.AddChart("Chart_Green", eChartType.Line);
                chartGreenTools.Title.Text = "Processing Time Green Tool(HDM/Focused/HDMQuick)[ms]";     //chartGreenTools.Title.Font.Color = Color.FromArgb(16, 203, 34);
                chartGreenTools.Title.Font.Size = 14;
                chartGreenTools.Title.Font.Bold = true;
                chartGreenTools.Title.Font.Italic = true;
                chartGreenTools.SetPosition(7, 7, 6, 6); // Start point to dispale of Chart  ex) 0,0,5,5 : Draw a chart from F1 Cell vs 1,1,6,6 : Draw a chart from G2 Cell
                chartGreenTools.SetSize(800, 600);

                ExcelAddress valueAddress_Data1_GreenTools = new ExcelAddress(startCellindex, 2, (Constants.RepeatProcess + (startCellindex - 1)), 2);
                ExcelAddress RepeatAddress_Data1_GreenTools = new ExcelAddress(startCellindex, 1, (Constants.RepeatProcess + (startCellindex - 1)), 1);
                var ser1_GreenTools = (chartGreenTools.Series.Add(valueAddress_Data1_GreenTools.Address, RepeatAddress_Data1_GreenTools.Address) as ExcelLineChartSerie);
                ser1_GreenTools.Header = "Green HDM";

                ExcelAddress valueAddress_Data2_GreenTools = new ExcelAddress(startCellindex, 3, (Constants.RepeatProcess + (startCellindex - 1)), 3);
                ExcelAddress RepeatAddress_Data2_GreenTools = new ExcelAddress(startCellindex, 1, (Constants.RepeatProcess + (startCellindex - 1)), 1);
                var ser2_GreenTools = (chartGreenTools.Series.Add(valueAddress_Data2_GreenTools.Address, RepeatAddress_Data2_GreenTools.Address) as ExcelLineChartSerie);
                ser2_GreenTools.Header = "Green Focused";

                ExcelAddress valueAddress_Data3_GreenTools = new ExcelAddress(startCellindex, 4, (Constants.RepeatProcess + (startCellindex - 1)), 4);
                ExcelAddress RepeatAddress_Data3_GreenTools = new ExcelAddress(startCellindex, 1, (Constants.RepeatProcess + (startCellindex - 1)), 1);
                var ser3_GreenTools = (chartGreenTools.Series.Add(valueAddress_Data3_GreenTools.Address, RepeatAddress_Data3_GreenTools.Address) as ExcelLineChartSerie);
                ser3_GreenTools.Header = "Green HDMQuick";

                chartGreenTools.Legend.Border.LineStyle = eLineStyle.Solid;
                chartGreenTools.Legend.Border.Fill.Style = eFillStyle.SolidFill;
                chartGreenTools.Legend.Border.Fill.Color = Color.DarkGreen;                //chartGreenTools.Border.Width = 1;
                chartGreenTools.Border.Fill.Color = Color.DarkGreen; //Color.FromArgb(16, 203, 34);
                // GreenTools Chart - End

                // BlueLocate Chart - Start
                // Creat BlueLocate Tool sheet
                ExcelWorksheet worksheetBlueLocateTool = package.Workbook.Worksheets["BlueLocateTool"];
                // Fill in index number and process time
                int columnBlueLocateTool = 1;
                for (int row = startCellindex; row < (Constants.RepeatProcess + startCellindex); row++)
                    worksheetBlueLocateTool.Cells[row, columnBlueLocateTool].Value = row - (startCellindex - 1);

                int colBlueLocateTool = 2;
                for (int row = startCellindex; row < (Constants.RepeatProcess + startCellindex); row++)
                    //worksheetBlueLocateTool.Cells[row, colBlueLocateTool].Value = int.Parse(GetPTimesBlueLocate[row - startCellindex]);
                    worksheetBlueLocateTool.Cells[row, colBlueLocateTool].Value = double.Parse(GetPTimesBlueLocate[row - startCellindex]);

                //JK-AddResult-Start - 2023.07.06
                int colBLFeatures = 3;
                for (int row = startCellindex; row < (Constants.RepeatProcess + startCellindex); row++)
                    worksheetBlueLocateTool.Cells[row, colBLFeatures].Value = int.Parse(GetBlueLocateNumFeatures[row - startCellindex]);

                // EvenNum : Tail
                int colBLFName1 = 4;
                for (int row = startCellindex; row < (Constants.RepeatProcess + startCellindex); row++)
                    worksheetBlueLocateTool.Cells[row, colBLFName1].Value = GetBlueLocateFeaturesNameEvenNum[row - startCellindex]; // First feature name is string type 
                int colBLFScore1 = 5;
                for (int row = startCellindex; row < (Constants.RepeatProcess + startCellindex); row++)
                    worksheetBlueLocateTool.Cells[row, colBLFScore1].Value = double.Parse(GetBlueLocateFeaturesScoreEvenNum[row - startCellindex]);
                int colBLFPosX1 = 6;
                for (int row = startCellindex; row < (Constants.RepeatProcess + startCellindex); row++)
                    worksheetBlueLocateTool.Cells[row, colBLFPosX1].Value = double.Parse(GetBlueLocateFeaturesPosXEvenNum[row - startCellindex]); // double

                int colBLFPosY1 = 7;
                for (int row = startCellindex; row < (Constants.RepeatProcess + startCellindex); row++)
                    worksheetBlueLocateTool.Cells[row, colBLFPosY1].Value = double.Parse(GetBlueLocateFeaturesPosYEvenNum[row - startCellindex]);
                int colBLFAngle1 = 8;
                for (int row = startCellindex; row < (Constants.RepeatProcess + startCellindex); row++)
                    worksheetBlueLocateTool.Cells[row, colBLFAngle1].Value = double.Parse(GetBlueLocateFeaturesAngleEvenNum[row - startCellindex]);
                int colBLFSizeH1 = 9;
                for (int row = startCellindex; row < (Constants.RepeatProcess + startCellindex); row++)
                    worksheetBlueLocateTool.Cells[row, colBLFSizeH1].Value = int.Parse(GetBlueLocateFeaturesSizeHeightEvenNum[row - startCellindex]);
                int colBLFSizeW1 = 10;
                for (int row = startCellindex; row < (Constants.RepeatProcess + startCellindex); row++)
                    worksheetBlueLocateTool.Cells[row, colBLFSizeW1].Value = int.Parse(GetBlueLocateFeaturesSizeWidthEvenNum[row - startCellindex]);

                // OddNum : Head
                int colBLFName2 = 11;
                for (int row = startCellindex; row < (Constants.RepeatProcess + startCellindex); row++)
                    worksheetBlueLocateTool.Cells[row, colBLFName2].Value = GetBlueLocateFeaturesNameOddNum[row - startCellindex]; // Second feature name is string type 
                int colBLFScore2 = 12;
                for (int row = startCellindex; row < (Constants.RepeatProcess + startCellindex); row++)
                    worksheetBlueLocateTool.Cells[row, colBLFScore2].Value = double.Parse(GetBlueLocateFeaturesScoreOddNum[row - startCellindex]);
                int colBLFPosX2 = 13;
                for (int row = startCellindex; row < (Constants.RepeatProcess + startCellindex); row++)
                    worksheetBlueLocateTool.Cells[row, colBLFPosX2].Value = double.Parse(GetBlueLocateFeaturesPosXOddNum[row - startCellindex]);
                int colBLFPosY2 = 14;
                for (int row = startCellindex; row < (Constants.RepeatProcess + startCellindex); row++)
                    worksheetBlueLocateTool.Cells[row, colBLFPosY2].Value = double.Parse(GetBlueLocateFeaturesPosYOddNum[row - startCellindex]);
                int colBLFAngle2 = 15;
                for (int row = startCellindex; row < (Constants.RepeatProcess + startCellindex); row++)
                    worksheetBlueLocateTool.Cells[row, colBLFAngle2].Value = double.Parse(GetBlueLocateFeaturesAngleOddNum[row - startCellindex]);
                int colBLFSizeH2 = 16;
                for (int row = startCellindex; row < (Constants.RepeatProcess + startCellindex); row++)
                    worksheetBlueLocateTool.Cells[row, colBLFSizeH2].Value = int.Parse(GetBlueLocateFeaturesSizeHeightOddNum[row - startCellindex]);
                int colBLFSizeW2 = 17;
                for (int row = startCellindex; row < (Constants.RepeatProcess + startCellindex); row++)
                    worksheetBlueLocateTool.Cells[row, colBLFSizeW2].Value = int.Parse(GetBlueLocateFeaturesSizeWidthOddNum[row - startCellindex]);
                //JK-AddResult-End - 2023.07.06

                // Fill in max, min, average for analysing process time of blue locate. 
                worksheetBlueLocateTool.Cells["B2"].Formula = $"MAX(B{startCellindex}:B{(Constants.RepeatProcess + startCellindex - 1)})";
                worksheetBlueLocateTool.Cells["B3"].Formula = $"MIN(B{startCellindex}:B{(Constants.RepeatProcess + startCellindex - 1)})";
                worksheetBlueLocateTool.Cells["B4"].Formula = $"AVERAGE(B{startCellindex}:B{(Constants.RepeatProcess + startCellindex - 1)})";
                // Adding chart for visibility fo analysing data.
                var chartBlueLocateTool = worksheetBlueLocateTool.Drawings.AddChart("Chart_BlueLocate", eChartType.Line);
                chartBlueLocateTool.Title.Text = "Processing Time Blue Locate Tool[ms]";                 //chartBlueLocateTool.Title.Font.Color = Color.FromArgb(0, 145, 255);
                chartBlueLocateTool.Title.Font.Size = 14;
                chartBlueLocateTool.Title.Font.Bold = true;
                chartBlueLocateTool.Title.Font.Italic = true;
                //chartBlueLocateTool.SetPosition(1, 1, 6, 6); // Graph position before 2023.07.06
                chartBlueLocateTool.SetPosition(7, 7, 6, 6);
                chartBlueLocateTool.SetSize(800, 600);

                ExcelAddress valueAddress_Data1_BlueLocateTool = new ExcelAddress(startCellindex, 2, (Constants.RepeatProcess + (startCellindex - 1)), 2);
                ExcelAddress RepeatAddress_Data1_BlueLocateTool = new ExcelAddress(startCellindex, 1, (Constants.RepeatProcess + (startCellindex - 1)), 1);
                var ser1_BlueLocateTool = (chartBlueLocateTool.Series.Add(valueAddress_Data1_BlueLocateTool.Address, RepeatAddress_Data1_BlueLocateTool.Address) as ExcelLineChartSerie);
                ser1_BlueLocateTool.Header = "Blue Locate";

                chartBlueLocateTool.Legend.Border.LineStyle = eLineStyle.Solid;
                chartBlueLocateTool.Legend.Border.Fill.Style = eFillStyle.SolidFill;
                chartBlueLocateTool.Legend.Border.Fill.Color = Color.DarkBlue;                //chartBlueLocateTool.Border.Width = 1;
                chartBlueLocateTool.Border.Fill.Color = Color.DarkBlue;
                // BlueLocare Chart - End

                // BlueRead Chart - Start
                // Create Blue Read sheet
                ExcelWorksheet worksheetBlueReadTool = package.Workbook.Worksheets["BlueReadTool"];
                // Fill in index number and process time
                int columnBlueRead = 1;
                for (int row = startCellindex; row < (Constants.RepeatProcess + startCellindex); row++)
                    worksheetBlueReadTool.Cells[row, columnBlueRead].Value = row - (startCellindex - 1);
                int colBlueReadTool = 2;
                for (int row = startCellindex; row < (Constants.RepeatProcess + startCellindex); row++)
                    worksheetBlueReadTool.Cells[row, colBlueReadTool].Value = int.Parse(GetPTimesBlueRead[row - startCellindex]);

                //JK-AddResult-Start - 2023.07.11
                int colBRFeatures = 3;
                for (int row = startCellindex; row < (Constants.RepeatProcess + startCellindex); row++)
                    worksheetBlueReadTool.Cells[row, colBRFeatures].Value = int.Parse(GetBlueReadMatchCountFeatures[row - startCellindex]);

                int colBRFName1 = 4;
                for (int row = startCellindex; row < (Constants.RepeatProcess + startCellindex); row++)
                    worksheetBlueReadTool.Cells[row, colBRFName1].Value = GetBlueReadMatchReault[(((row - startCellindex) * 9) + 0)].Name;//.ToString();                                
                int colBRFScore1 = 5;
                for (int row = startCellindex; row < (Constants.RepeatProcess + startCellindex); row++)
                    worksheetBlueReadTool.Cells[row, colBRFScore1].Value = GetBlueReadMatchReault[(((row - startCellindex) * 9) + 0)].Score;
                int colBRFName2 = 6;
                for (int row = startCellindex; row < (Constants.RepeatProcess + startCellindex); row++)
                    worksheetBlueReadTool.Cells[row, colBRFName2].Value = GetBlueReadMatchReault[(((row - startCellindex) * 9) + 1)].Name;//.ToString();
                int colBRFScore2 = 7;
                for (int row = startCellindex; row < (Constants.RepeatProcess + startCellindex); row++)
                    worksheetBlueReadTool.Cells[row, colBRFScore2].Value = GetBlueReadMatchReault[(((row - startCellindex) * 9) + 1)].Score;
                int colBRFName3 = 8;
                for (int row = startCellindex; row < (Constants.RepeatProcess + startCellindex); row++)
                    worksheetBlueReadTool.Cells[row, colBRFName3].Value = GetBlueReadMatchReault[(((row - startCellindex) * 9) + 2)].Name;//.ToString();
                int colBRFScore3 = 9;
                for (int row = startCellindex; row < (Constants.RepeatProcess + startCellindex); row++)
                    worksheetBlueReadTool.Cells[row, colBRFScore3].Value = GetBlueReadMatchReault[(((row - startCellindex) * 9) + 2)].Score;
                int colBRFName4 = 10;
                for (int row = startCellindex; row < (Constants.RepeatProcess + startCellindex); row++)
                    worksheetBlueReadTool.Cells[row, colBRFName4].Value = GetBlueReadMatchReault[(((row - startCellindex) * 9) + 3)].Name;//.ToString();
                int colBRFScore4 = 11;
                for (int row = startCellindex; row < (Constants.RepeatProcess + startCellindex); row++)
                    worksheetBlueReadTool.Cells[row, colBRFScore4].Value = GetBlueReadMatchReault[(((row - startCellindex) * 9) + 3)].Score;

                int colBRFName5 = 12;
                for (int row = startCellindex; row < (Constants.RepeatProcess + startCellindex); row++)
                    worksheetBlueReadTool.Cells[row, colBRFName5].Value = GetBlueReadMatchReault[(((row - startCellindex) * 9) + 4)].Name;//.ToString();
                int colBRFScore5 = 13;
                for (int row = startCellindex; row < (Constants.RepeatProcess + startCellindex); row++)
                    worksheetBlueReadTool.Cells[row, colBRFScore5].Value = GetBlueReadMatchReault[(((row - startCellindex) * 9) + 4)].Score;
                int colBRFName6 = 14;
                for (int row = startCellindex; row < (Constants.RepeatProcess + startCellindex); row++)
                    worksheetBlueReadTool.Cells[row, colBRFName6].Value = GetBlueReadMatchReault[(((row - startCellindex) * 9) + 5)].Name;//.ToString();
                int colBRFScore6 = 15;
                for (int row = startCellindex; row < (Constants.RepeatProcess + startCellindex); row++)
                    worksheetBlueReadTool.Cells[row, colBRFScore6].Value = GetBlueReadMatchReault[(((row - startCellindex) * 9) + 5)].Score;
                int colBRFName7 = 16;
                for (int row = startCellindex; row < (Constants.RepeatProcess + startCellindex); row++)
                    worksheetBlueReadTool.Cells[row, colBRFName7].Value = GetBlueReadMatchReault[(((row - startCellindex) * 9) + 6)].Name;//.ToString();
                int colBRFScore7 = 17;
                for (int row = startCellindex; row < (Constants.RepeatProcess + startCellindex); row++)
                    worksheetBlueReadTool.Cells[row, colBRFScore7].Value = GetBlueReadMatchReault[(((row - startCellindex) * 9) + 6)].Score;
                int colBRFName8 = 18;
                for (int row = startCellindex; row < (Constants.RepeatProcess + startCellindex); row++)
                    worksheetBlueReadTool.Cells[row, colBRFName8].Value = GetBlueReadMatchReault[(((row - startCellindex) * 9) + 7)].Name;//.ToString();
                int colBRFScore8 = 19;
                for (int row = startCellindex; row < (Constants.RepeatProcess + startCellindex); row++)
                    worksheetBlueReadTool.Cells[row, colBRFScore8].Value = GetBlueReadMatchReault[(((row - startCellindex) * 9) + 7)].Score;
                int colBRFName9 = 20;
                for (int row = startCellindex; row < (Constants.RepeatProcess + startCellindex); row++)
                    worksheetBlueReadTool.Cells[row, colBRFName9].Value = GetBlueReadMatchReault[(((row - startCellindex) * 9) + 8)].Name;//.ToString();
                int colBRFScore9 = 21;
                for (int row = startCellindex; row < (Constants.RepeatProcess + startCellindex); row++)
                    worksheetBlueReadTool.Cells[row, colBRFScore9].Value = GetBlueReadMatchReault[(((row - startCellindex) * 9) + 8)].Score;

                //  GetBlueReadMatchReault.ElementAt(1); // 결과들을 입력 하면서 배열화 된 것인가.?
                //JK-AddResult-End - 2023.07.11
                // Fill in max, min, average for analysing process time of blue locate. 
                worksheetBlueReadTool.Cells["B2"].Formula = $"MAX(B{startCellindex}:B{(Constants.RepeatProcess + startCellindex - 1)})";
                worksheetBlueReadTool.Cells["B3"].Formula = $"MIN(B{startCellindex}:B{(Constants.RepeatProcess + startCellindex - 1)})";
                worksheetBlueReadTool.Cells["B4"].Formula = $"AVERAGE(B{startCellindex}:B{(Constants.RepeatProcess + startCellindex - 1)})";
                // Adding chart for visibility fo analysing data.
                var chartBlueReadTool = worksheetBlueReadTool.Drawings.AddChart("Chart_BlueRead", eChartType.Line);
                chartBlueReadTool.Title.Text = "Processing Time Blue Read Tool[ms]";
                chartBlueReadTool.Title.Font.Size = 14; //chartBlueReadTool.Title.Font.Color = Color.FromArgb(0, 75, 163);
                chartBlueReadTool.Title.Font.Bold = true;
                chartBlueReadTool.Title.Font.Italic = true;
                //chartBlueReadTool.SetPosition(1, 1, 6, 6); // Start point to dispale of Chart  ex) 0,0,5,5 : Draw a chart from F1 Cell vs 1,1,6,6 : Draw a chart from G2 Cell
                chartBlueReadTool.SetPosition(7, 7, 6, 6); // Start point to dispale of Chart  ex) 0,0,5,5 : Draw a chart from F1 Cell vs 1,1,6,6 : Draw a chart from G2 Cell
                chartBlueReadTool.SetSize(800, 600);

                ExcelAddress valueAddress_Data1_BlueReadTool = new ExcelAddress(startCellindex, 2, (Constants.RepeatProcess + (startCellindex - 1)), 2);
                ExcelAddress RepeatAddress_Data1_BlueReadTool = new ExcelAddress(startCellindex, 1, (Constants.RepeatProcess + (startCellindex - 1)), 1);
                var ser1_BlueReadTool = (chartBlueReadTool.Series.Add(valueAddress_Data1_BlueReadTool.Address, RepeatAddress_Data1_BlueReadTool.Address) as ExcelLineChartSerie);
                ser1_BlueReadTool.Header = "Blue Read";

                chartBlueReadTool.Legend.Border.LineStyle = eLineStyle.Solid;
                chartBlueReadTool.Legend.Border.Fill.Style = eFillStyle.SolidFill;
                chartBlueReadTool.Legend.Border.Fill.Color = Color.DarkBlue;                //chartBlueReadTool.Border.Width = 1;
                chartBlueReadTool.Border.Fill.Color = Color.DarkBlue;
                // BlueRead Chart - End

                package.Save();
            }
            Console.WriteLine("Complete - adding chart with using EPPlus.4.5.3.3");
            Console.WriteLine();
        }
    }
}
// Save info
// H:\_2_CCCCC_modification_addingresult\QAGPT_No164_28557\QAGPT_Build_164_artifacts\target_directory\QAGPTPJT\QAGPTPJT
// Program_TaskDay_20230717-2043.cs
// 오후 8:44 2023-07-17
