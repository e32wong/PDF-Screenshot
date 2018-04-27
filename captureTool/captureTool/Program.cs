using System;
using System.Runtime.InteropServices;
using System.Drawing;
using System.Drawing.Imaging;
using System.Diagnostics;
using System.Threading;
using System.IO;

using System.Windows.Forms;
using IronOcr;

namespace ConsoleApp1
{
    static class SendKey
    {
        const UInt32 WM_KEYDOWN = 0x0100;
        const int VK_F5 = 0x74;



    }

    class Hello
    {
        [DllImport("user32.dll", ExactSpelling = true, CharSet = CharSet.Auto)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        public static extern bool PostMessage(IntPtr hWnd, UInt32 Msg, int wParam, int lParam);

        [DllImport("user32.dll", EntryPoint = "ShowWindow", SetLastError = true)]
        static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        public class ScreenCapture
        {


            public Image CaptureScreen()
            {
                return CaptureWindow(User32.GetDesktopWindow());
            }

            public static IntPtr WinGetHandle(string wName)
            {
                IntPtr hWnd = IntPtr.Zero;
                foreach (Process pList in Process.GetProcesses())
                {
                    if (pList.MainWindowTitle.Contains(wName))
                    {
                        Debug.WriteLine("found the window");
                        hWnd = pList.MainWindowHandle;
                    }
                }
                return hWnd;
            }

            /// <summary>
            /// Creates an Image object containing a screen shot of a specific window
            /// </summary>
            /// <param name="handle">The handle to the window. (In windows forms, this is obtained by the Handle property)</param>
            /// <returns></returns>
            public Image CaptureWindow(IntPtr handle)
            {
                // get te hDC of the target window
                IntPtr hdcSrc = User32.GetWindowDC(handle);
                // get the size
                User32.RECT windowRect = new User32.RECT();
                User32.GetWindowRect(handle, ref windowRect);
                int width = windowRect.right - windowRect.left;
                int height = windowRect.bottom - windowRect.top;
                // create a device context we can copy to
                IntPtr hdcDest = GDI32.CreateCompatibleDC(hdcSrc);
                // create a bitmap we can copy it to,
                // using GetDeviceCaps to get the width/height
                IntPtr hBitmap = GDI32.CreateCompatibleBitmap(hdcSrc, width, height);
                // select the bitmap object
                IntPtr hOld = GDI32.SelectObject(hdcDest, hBitmap);
                // bitblt over
                GDI32.BitBlt(hdcDest, 0, 0, width, height, hdcSrc, 0, 0, GDI32.SRCCOPY);
                // restore selection
                GDI32.SelectObject(hdcDest, hOld);
                // clean up 
                GDI32.DeleteDC(hdcDest);
                User32.ReleaseDC(handle, hdcSrc);
                // get a .NET image object for it
                Image img = Image.FromHbitmap(hBitmap);
                // free up the Bitmap object
                GDI32.DeleteObject(hBitmap);
                return img;
            }
            /// <summary>
            /// Captures a screen shot of a specific window, and saves it to a file
            /// </summary>
            /// <param name="handle"></param>
            /// <param name="filename"></param>
            /// <param name="format"></param>
            public void CaptureWindowToFile(IntPtr handle, string filename, ImageFormat format)
            {
                Image img = CaptureWindow(handle);
                img.Save(filename, format);
                img.Dispose();
            }
            /// <summary>
            /// Captures a screen shot of the entire desktop, and saves it to a file
            /// </summary>
            /// <param name="filename"></param>
            /// <param name="format"></param>
            public void CaptureScreenToFile(string filename, ImageFormat format)
            {
                Image img = CaptureScreen();
                img.Save(filename, format);
            }

            /// <summary>
            /// Helper class containing Gdi32 API functions
            /// </summary>
            private class GDI32
            {

                public const int SRCCOPY = 0x00CC0020; // BitBlt dwRop parameter
                [DllImport("gdi32.dll")]
                public static extern bool BitBlt(IntPtr hObject, int nXDest, int nYDest,
                    int nWidth, int nHeight, IntPtr hObjectSource,
                    int nXSrc, int nYSrc, int dwRop);
                [DllImport("gdi32.dll")]
                public static extern IntPtr CreateCompatibleBitmap(IntPtr hDC, int nWidth,
                    int nHeight);
                [DllImport("gdi32.dll")]
                public static extern IntPtr CreateCompatibleDC(IntPtr hDC);
                [DllImport("gdi32.dll")]
                public static extern bool DeleteDC(IntPtr hDC);
                [DllImport("gdi32.dll")]
                public static extern bool DeleteObject(IntPtr hObject);
                [DllImport("gdi32.dll")]
                public static extern IntPtr SelectObject(IntPtr hDC, IntPtr hObject);
            }

            /// <summary>
            /// Helper class containing User32 API functions
            /// </summary>
            private class User32
            {
                [StructLayout(LayoutKind.Sequential)]
                public struct RECT
                {
                    public int left;
                    public int top;
                    public int right;
                    public int bottom;
                }
                [DllImport("user32.dll")]
                public static extern IntPtr GetDesktopWindow();
                [DllImport("user32.dll")]
                public static extern IntPtr GetWindowDC(IntPtr hWnd);
                [DllImport("user32.dll")]
                public static extern IntPtr ReleaseDC(IntPtr hWnd, IntPtr hDC);
                [DllImport("user32.dll")]
                public static extern IntPtr GetWindowRect(IntPtr hWnd, ref RECT rect);
            }
        }

        public static IntPtr WinGetHandle2(string wName)
        {
            IntPtr hWnd = IntPtr.Zero;
            foreach (Process pList in Process.GetProcesses())
            {
                if (pList.MainWindowTitle.Contains(wName))
                {
                    hWnd = pList.MainWindowHandle;
                    Console.WriteLine(hWnd.ToString());
                }
            }
            return hWnd;
        }

        private static void screenshotSpecificWindow(string programName, string savePath)
        {
            // take the screenshot
            ScreenCapture screencap = new ScreenCapture();
            IntPtr handle = ScreenCapture.WinGetHandle(programName);
            screencap.CaptureWindowToFile(handle, savePath, ImageFormat.Png);
        }

        private static Boolean launchFile(string fileName, string programName)
        {
            try
            {
                // open a pdf file
                string file = fileName;
                ProcessStartInfo pi = new ProcessStartInfo(file);
                pi.Arguments = Path.GetFileName(file);
                pi.UseShellExecute = true;
                pi.WorkingDirectory = Path.GetDirectoryName(file);
                pi.FileName = programName;
                pi.Verb = "OPEN";
                Process.Start(pi);
                return true;
            } catch (Exception e)
            {
                Console.WriteLine("Error at opening file");
                return false;
            }

        }

        private static void sendKeyToApplication(Process[] processes, string command)
        {
            foreach (Process proc in processes)
            {
                IntPtr h = proc.MainWindowHandle;
                SetForegroundWindow(h);
                SendKeys.SendWait(command);
                //Console.WriteLine("sent to " + proc.ToString());
            }
        }

        private static void killProcess(Process[] processes)
        {
            foreach (Process proc in processes)
            {
                try
                {
                    //proc.Close();
                    proc.Kill();
                    //proc.CloseMainWindow();
                    //proc.WaitForExit();
                }
                catch (Exception e)
                {
                    Console.WriteLine("No instance of the app running");
                    Console.WriteLine(e.ToString());

                }
            }
        }

        private static Boolean checkErrorMessage(string imagePath)
        {
            Boolean hasError = false;

            /*
            Console.WriteLine("Checking for error messages...");
            AutoOcr Ocr = new AutoOcr();
            OcrResult result = Ocr.Read(imagePath);
            Console.WriteLine("Result:");
            Console.WriteLine(result.Text);
            */
            
            AdvancedOcr OcrAdvanced = new AdvancedOcr()
            {
                CleanBackgroundNoise = false,
                EnhanceContrast = false,
                EnhanceResolution = false,
                Language = IronOcr.Languages.English.OcrLanguagePack,
                Strategy = IronOcr.AdvancedOcr.OcrStrategy.Advanced,
                ColorSpace = AdvancedOcr.OcrColorSpace.Color,
                DetectWhiteTextOnDarkBackgrounds = true,
                InputImageType = AdvancedOcr.InputTypes.AutoDetect,
                RotateAndStraighten = false,
                ReadBarCodes = false,
                ColorDepth = 4
            };
            OcrResult result = OcrAdvanced.Read(imagePath);
            Console.WriteLine(result.Text);

            String[] errorTerms = new string[] { "File", "Edit", "Home", "Tools" };
            String resultStr = result.Text;
            foreach (String badTerm in errorTerms)
            {
                if (resultStr.Contains(badTerm))
                {
                    Console.WriteLine("detected error message with term: " + badTerm);
                    hasError = true;
                }
            }
            
            return hasError;
        }

        private static Boolean CropImage(int x, int y, int width, int height, string imagePath, string cropPath)
        {
            Boolean hasFailed = false;

            //using (var originalImage = new Bitmap(imagePath))
            var originalImage = new Bitmap(imagePath);
            // check dimension of the file
            if (originalImage.Height == 1 && originalImage.Width == 1)
            {
                hasFailed = true;
            }

            if (hasFailed == false)
            {
                Bitmap croppedImage;
                Rectangle crop = new Rectangle(x, y, width, height);
                croppedImage = originalImage.Clone(crop, originalImage.PixelFormat);
                //croppedImage = new Bitmap(originalImage);
                croppedImage.Save(cropPath, ImageFormat.Png);
                croppedImage.Dispose();
                croppedImage = null;
            }

            originalImage.Dispose();
            originalImage = null;

            return hasFailed;
        }

        private static void screenshotEntireScreen(String savePath)
        {
            var bmpScreenshot = new Bitmap(Screen.PrimaryScreen.Bounds.Width,
                Screen.PrimaryScreen.Bounds.Height,
                PixelFormat.Format32bppArgb);
            var gfxScreenshot = Graphics.FromImage(bmpScreenshot);
            gfxScreenshot.CopyFromScreen(0, 0,
                0,
                0,
                Screen.PrimaryScreen.Bounds.Size);
            bmpScreenshot.Save(savePath, ImageFormat.Png);
        }

        private static float getCpuUsage()
        {
            float usage = 0;

            PerformanceCounter totalCPU = new PerformanceCounter("Processor", "% Processor Time", "_Total");
            totalCPU.NextValue(); // will always be 0
            System.Threading.Thread.Sleep(1000);
            usage = totalCPU.NextValue();

            return usage;
        }

        private static Boolean waitForCPU(int numTriesMax, String screenshotFolder)
        {
            Boolean status = false;

            int numTries = 0;
            while (numTries < numTriesMax)
            {
                float usageValue = getCpuUsage();
                logMessage("CPU @ " + usageValue + "%\n", screenshotFolder);
                if (usageValue >= 10)
                {
                    Thread.Sleep(1000);
                }
                else
                {
                    status = true;
                    break;
                }
                numTries = numTries + 1;
            }

            return status;
        }

        private static void processFile(string targetFile, string screenshotFolder)
        {

            var watch = System.Diagnostics.Stopwatch.StartNew();
            
            Process[] processesFault = Process.GetProcessesByName("WerFault");
            killProcess(processesFault);
            Thread.Sleep(100);

            Boolean status = true;
            String pureFileName = Path.GetFileName(targetFile);
            String screenshotSavePath = screenshotFolder + pureFileName.Substring(0, pureFileName.Length - 4) + ".png";
            //String screenshotSavePath = screenshotFolder + "screenshot.png";
            Console.WriteLine(screenshotSavePath);

            // https://stackoverflow.com/questions/42498440/sendkeys-to-a-specific-program-without-it-being-in-focus
            Boolean launchSuccessful = launchFile(targetFile, "C:\\Program Files (x86)\\Adobe\\Acrobat Reader DC\\Reader\\AcroRd32.exe");
            if (launchSuccessful == false)
            {
                status = false;
            } else
            {
                Thread.Sleep(500);
                Process[] processes = Process.GetProcessesByName("AcroRd32");
                foreach (Process proc in processes)
                {
                    ShowWindow(proc.MainWindowHandle, 3);
                }

                // send esc in the case that it crashed
                sendKeyToApplication(processes, "{ESC}");
                Thread.Sleep(200);
                
                Boolean cpuGood = waitForCPU(20, screenshotFolder);
                if (cpuGood)
                {
                    //Thread.Sleep(1000);

                    /*
                    Process[] processesFault3 = Process.GetProcessesByName("WerFault");
                    killProcess(processesFault3);
                    Thread.Sleep(100);
                    */

                    processes = Process.GetProcessesByName("AcroRd32");
                    if (processes.Length > 0)
                    {
                        File.AppendAllText(@"C:\Users\edmund\Desktop\log.txt", "Found application!\n");

                        // try full screen it
                        sendKeyToApplication(processes, "{ENTER}");
                        Thread.Sleep(200);
                        sendKeyToApplication(processes, "{ENTER}");
                        Thread.Sleep(200);
                        sendKeyToApplication(processes, "{HOME}");
                        Thread.Sleep(200);
                        sendKeyToApplication(processes, "^(l)");
                        Thread.Sleep(2000);

                        // take screenshot
                        screenshotSpecificWindow("Adobe Acrobat Reader DC", screenshotSavePath);
                        //screenshotEntireScreen(screenshotSavePath);
                        //string cropPath = @"C:\Users\edmund\Desktop\cropped.png";

                        //CropImage(0, 0, 300, 100, screenshotSavePath, cropPath);
                        // check if the screenshot located an error message
                        //Boolean hasError = checkErrorMessage(cropPath);

                        // close the application depending if there is an error

                        // send enter command and full screen command
                        logMessage("Closing the app\n", screenshotFolder);
                        sendKeyToApplication(processes, "{ESC}");
                        Thread.Sleep(200);
                        sendKeyToApplication(processes, "{ENTER}");
                        Thread.Sleep(200);
                        processes = Process.GetProcessesByName("AcroRd32");
                        if (processes.Length == 0)
                        {
                            File.AppendAllText(@"C:\Users\edmund\Desktop\log.txt", "Crashed after full screen\n");
                            status = false;
                        }
                        /*else
                        {
                            foreach (Process proc in processes)
                            {
                                // ShowWindow(proc.MainWindowHandle, 3);

                                sendKeyToApplication(processes, "%{F4}");
                                Thread.Sleep(1000);

                                // sometimes it is blocked asking if you want to save
                                processes = Process.GetProcessesByName("AcroRd32");
                                if (processes.Length > 0)
                                {
                                    sendKeyToApplication(processes, "n");
                                    Thread.Sleep(1000);
                                }

                                break;
                            }
                        }*/
                    }
                    else
                    {
                        logMessage("Crashed right after launch\n", screenshotFolder);
                        status = false;
                    }
                }
                else
                {
                    status = false;
                    logMessage("Entire program froze\n", screenshotFolder);
                }
            }


            // kill and clean environment
            Process[] processesEnd = Process.GetProcessesByName("AcroRd32");
            killProcess(processesEnd);

            Console.WriteLine("Exception in windows, killing it and sleeping 5 seconds");
            Thread.Sleep(500);

            Process[] processesFault2 = Process.GetProcessesByName("WerFault");
            killProcess(processesFault2);
            Thread.Sleep(500);
            
            watch.Stop();
            long elapsedMs = watch.ElapsedMilliseconds;

            logMessage(status.ToString() + "\n", screenshotFolder);
            logMessage((elapsedMs / 1000).ToString() + "\n", screenshotFolder);
            writeToCSV(Path.GetFileName(targetFile), status, elapsedMs, screenshotFolder);
        }

        private static void logMessage(String message, String screenshotFolder)
        {
            Console.Write(message);
            File.AppendAllText(screenshotFolder + "log.txt", message);
        }

        private static void writeToCSV(String fileName, Boolean result, long elapsedMs, String screenshotFolder)
        {
            var newLine = string.Format("{0},{1},{2}\n", fileName, (float)elapsedMs / 1000, result);
            File.AppendAllText(screenshotFolder + "log.csv", newLine);
        }

        private static void batchProcess(string sourcePath, string outputPath)
        {

            foreach (Process PPath in Process.GetProcesses())
            {
                Console.WriteLine(PPath.ProcessName.ToString());
            }

            var watch = System.Diagnostics.Stopwatch.StartNew();

            // get list of files from folder
            string[] filePaths = Directory.GetFiles(sourcePath, "*.pdf");
            string screenshotFolder = outputPath;

            // check if screenshot folder exists
            try
            {
                if (!Directory.Exists(screenshotFolder))
                {
                    Console.WriteLine("deleting the old folder");
                    Directory.CreateDirectory(screenshotFolder);
                }
                else
                {
                    Console.WriteLine("deleting the old folder and recreating it");
                    Directory.Delete(screenshotFolder, true);
                    Thread.Sleep(5000);
                    Directory.CreateDirectory(screenshotFolder);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("Exception with the screenshot folder:" + e);
            }

            // kill and clean environment before we start the screenshot
            Process[] processes = Process.GetProcessesByName("AcroRd32");
            killProcess(processes);

            // remove the old log file
            if (File.Exists(@"C:\Users\edmund\Desktop\log.txt"))
            {
                File.Delete(@"C:\Users\edmund\Desktop\log.txt");
            }
            if (File.Exists(@"C:\Users\edmund\Desktop\log.csv"))
            {
                File.Delete(@"C:\Users\edmund\Desktop\log.csv");
            }

            int index = 1;
            foreach (string pdfFilePath in filePaths)
            {
                logMessage("\n\n" + index + "/" + filePaths.Length + "\n", screenshotFolder);
                logMessage(pdfFilePath + "\n", screenshotFolder);

                processFile(pdfFilePath, screenshotFolder);

                
                
                index = index + 1;
            }

            watch.Stop();
            long elapsedMs = watch.ElapsedMilliseconds;

            logMessage("Total execution time:\n", screenshotFolder);
            logMessage((elapsedMs / 1000).ToString() + "\n", screenshotFolder);

        }



        static void Main()
        {
            batchProcess(@"C:\Users\edmund\Desktop\testSuite\test2", @"C:\Users\edmund\Desktop\screenshots\test\");

            //batchProcess(@"C:\Users\edmund\Desktop\phase2\chrome\", @"C:\Users\edmund\Desktop\screenshots\chromeP2\");

            //batchProcess(@"C:\Users\edmund\Desktop\filteredFiles\chrome\", @"C:\Users\edmund\Desktop\screenshots\chrome\");
            /*
            batchProcess(@"C:\Users\edmund\Desktop\filteredFiles\apache\", @"C:\Users\edmund\Desktop\screenshots\apache\");
            batchProcess(@"C:\Users\edmund\Desktop\filteredFiles\freedesktop\", @"C:\Users\edmund\Desktop\screenshots\freedesktop\");
            batchProcess(@"C:\Users\edmund\Desktop\filteredFiles\ghostscript\", @"C:\Users\edmund\Desktop\screenshots\ghostscript\");
            batchProcess(@"C:\Users\edmund\Desktop\filteredFiles\gnome\", @"C:\Users\edmund\Desktop\screenshots\gnome\");
            batchProcess(@"C:\Users\edmund\Desktop\filteredFiles\govdocs\", @"C:\Users\edmund\Desktop\screenshots\govdocs\");
            batchProcess(@"C:\Users\edmund\Desktop\filteredFiles\kde\", @"C:\Users\edmund\Desktop\screenshots\kde\");
            batchProcess(@"C:\Users\edmund\Desktop\filteredFiles\launchpad\", @"C:\Users\edmund\Desktop\screenshots\launchpad\");
            batchProcess(@"C:\Users\edmund\Desktop\filteredFiles\mozilla\", @"C:\Users\edmund\Desktop\screenshots\mozilla\");

            batchProcess(@"C:\Users\edmund\Desktop\filteredFiles\forceopen\corrupted1_1\", @"C:\Users\edmund\Desktop\screenshots\corrupted1_1");
            batchProcess(@"C:\Users\edmund\Desktop\filteredFiles\forceopen\corrupted1_2\", @"C:\Users\edmund\Desktop\screenshots\corrupted1_2");
            batchProcess(@"C:\Users\edmund\Desktop\filteredFiles\forceopen\corrupted1_4\", @"C:\Users\edmund\Desktop\screenshots\corrupted1_4");
            batchProcess(@"C:\Users\edmund\Desktop\filteredFiles\forceopen\corrupted1_8\", @"C:\Users\edmund\Desktop\screenshots\corrupted1_8");
            batchProcess(@"C:\Users\edmund\Desktop\filteredFiles\forceopen\corrupted1_16\", @"C:\Users\edmund\Desktop\screenshots\corrupted1_16");

            batchProcess(@"C:\Users\edmund\Desktop\filteredFiles\forceopen\corrupted2_1\", @"C:\Users\edmund\Desktop\screenshots\corrupted2_1");
            batchProcess(@"C:\Users\edmund\Desktop\filteredFiles\forceopen\corrupted2_2\", @"C:\Users\edmund\Desktop\screenshots\corrupted2_2");
            batchProcess(@"C:\Users\edmund\Desktop\filteredFiles\forceopen\corrupted2_4\", @"C:\Users\edmund\Desktop\screenshots\corrupted2_4");
            batchProcess(@"C:\Users\edmund\Desktop\filteredFiles\forceopen\corrupted2_8\", @"C:\Users\edmund\Desktop\screenshots\corrupted2_8");
            batchProcess(@"C:\Users\edmund\Desktop\filteredFiles\forceopen\corrupted2_16\", @"C:\Users\edmund\Desktop\screenshots\corrupted2_16");
            */

            //string sourcePath = @"C:\Users\edmund\Desktop\testSuite\test2";
            /*
            foreach (Process PPath in Process.GetProcesses())
            {
                Console.WriteLine(PPath.ProcessName.ToString());
            }

            // https://stackoverflow.com/questions/42498440/sendkeys-to-a-specific-program-without-it-being-in-focus
            launchFile(@"C:\Users\edmund\Desktop\minimal.pdf", "C:\\Program Files (x86)\\Adobe\\Acrobat Reader DC\\Reader\\AcroRd32.exe");

            Thread.Sleep(3000);

            Process[] processes = Process.GetProcessesByName("AcroRd32");
            foreach (Process proc in processes)
            {
                ShowWindow(proc.MainWindowHandle, 3);
            }


            string imagePath = @"C:\Users\edmund\Desktop\screenshot.png";

            takescreenshot("Adobe", imagePath);

            // check if the screenshot located an error message
            Boolean hasError = checkErrorMessage(imagePath);

            if (!hasError)
            {
                // send enter command and full screen command
                Console.WriteLine("Total number of processes: " + processes.Length);
                sendKeyToApplication(processes, "{ENTER}");
                sendKeyToApplication(processes, "^(l)");

                Thread.Sleep(3000);
                
                sendKeyToApplication(processes, "{ESC}");
                sendKeyToApplication(processes, "%{F4}");
            } else
            {
                // send enter command only and close it
                sendKeyToApplication(processes, "{ENTER}");
                sendKeyToApplication(processes, "%{F4}");

                //killProcess(processes);
            }
            */
            Console.WriteLine("Press any key to exit the tool..");
            Console.Read();
            
        }

    }
}
