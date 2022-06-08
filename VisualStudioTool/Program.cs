using System;
using System.IO;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Linq;
using EnvDTE;

namespace VisualStudioTool
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            if (args.Length < 3)
            {
                Console.WriteLine("Too few arguments.");
                return;
            }

            bool attachDebugger = args[0].Equals("debug", StringComparison.OrdinalIgnoreCase);
            bool stopDebugger = args[0].Equals("stopdebug", StringComparison.OrdinalIgnoreCase);
            bool openFile = args[0].Equals("openf", StringComparison.OrdinalIgnoreCase);
            bool buildSolution = args[0].Equals("build", StringComparison.OrdinalIgnoreCase);
            bool reBuildSolution = args[0].Equals("rebuild", StringComparison.OrdinalIgnoreCase);

            if(buildSolution && args.Length < 4)
            {
                Console.WriteLine("Too few arguments to build the project.");
                return;
            }
            else if (attachDebugger && args.Length < 4)
            {
                Console.WriteLine("Too few arguments to attach to debugger.");
                return;
            }
            else if (openFile && args.Length < 5)
            {
                Console.WriteLine("Too few arguments to open file.");
                return;
            }

            string vsPath = args[1];
            string solutionPath = args[2];
            try
            {
                if (buildSolution)
                {
                    string buildType = args[3];
                    BuildSolution(vsPath, solutionPath, buildType);
                    return;
                }
                if(reBuildSolution)
                {
                    string buildType = args[3];
                    RebuildSolution(vsPath, solutionPath, buildType);
                    return;
                }
                DTE dte = FindRunningVSWithOurSolution(solutionPath);
                if (dte == null)
                {
                    dte = CreateNewRunningVSWithOurSolution(vsPath, solutionPath, openFile);
                }

                //Handle the case where the project solution is already opened
                if (!MessageFilter.FilterRegistered)
                {
                    MessageFilter.Register();
                }

                if (openFile && int.TryParse(args[4], out int fileLine))
                {
                    string filePath = args[3];
                    HaveRunningVSOpenFile(dte, filePath, fileLine);
                }
                else if (attachDebugger && int.TryParse(args[3], out int processID))
                {
                    AttachDebugger(dte, processID);
                }
                else if (stopDebugger)
                {
                    StopDebugger(dte);
                }
                else
                {
                    OpenSolutionInVS(dte);
                }
                Marshal.ReleaseComObject(dte);
                MessageFilter.Revoke();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                Console.WriteLine(e.InnerException);
                Console.WriteLine(e.StackTrace);
            }
        }

        [DllImport("ole32.dll")]
        private static extern int CreateBindCtx(uint reserved, out IBindCtx ppbc);

        static DTE FindRunningVSWithOurSolution(string solutionPath)
        {
            DTE dte = null;
            IBindCtx bindCtx = null;
            IRunningObjectTable rot = null;
            IEnumMoniker enumMonikers = null;

            try
            {
                Marshal.ThrowExceptionForHR(CreateBindCtx(reserved: 0, ppbc: out bindCtx));
                bindCtx.GetRunningObjectTable(out rot);
                rot.EnumRunning(out enumMonikers);

                IMoniker[] moniker = new IMoniker[1];
                IntPtr numberFetched = IntPtr.Zero;
                while (enumMonikers.Next(1, moniker, numberFetched) == 0)
                {
                    IMoniker runningObjectMoniker = moniker[0];
                    Marshal.ThrowExceptionForHR(rot.GetObject(runningObjectMoniker, out object runningObject));
                    DTE dte2 = runningObject as DTE;
                    if (dte2 != null)
                    {
                        if (dte2.Solution.FullName == solutionPath)
                        {
                            dte = dte2;
                            break;
                        }
                    }
                }
            }
            finally
            {
                if (enumMonikers != null)
                {
                    Marshal.ReleaseComObject(enumMonikers);
                }

                if (rot != null)
                {
                    Marshal.ReleaseComObject(rot);
                }

                if (bindCtx != null)
                {
                    Marshal.ReleaseComObject(bindCtx);
                }
            }

            return dte;
        }

        static DTE FindRunningVSWithOurProcess(int processId)
        {
            string progId = ":" + processId.ToString();
            DTE dte = null;
            IBindCtx bindCtx = null;
            IRunningObjectTable rot = null;
            IEnumMoniker enumMonikers = null;

            try
            {
                Marshal.ThrowExceptionForHR(CreateBindCtx(reserved: 0, ppbc: out bindCtx));
                bindCtx.GetRunningObjectTable(out rot);
                rot.EnumRunning(out enumMonikers);

                IMoniker[] moniker = new IMoniker[1];
                IntPtr numberFetched = IntPtr.Zero;
                while (enumMonikers.Next(1, moniker, numberFetched) == 0)
                {
                    IMoniker runningObjectMoniker = moniker[0];
                    string name = null;

                    try
                    {
                        if (runningObjectMoniker != null)
                        {
                            runningObjectMoniker.GetDisplayName(bindCtx, null, out name);
                        }
                    }
                    catch (UnauthorizedAccessException)
                    {
                        // Do nothing, there is something in the ROT that we do not have access to.
                    }

                    if (!string.IsNullOrEmpty(name) && name.Contains(progId))
                    {
                        Marshal.ThrowExceptionForHR(rot.GetObject(runningObjectMoniker, out object runningObject));
                        dte = runningObject as DTE;
                        if (dte != null)
                        {
                            break;
                        }
                    }
                }
            }
            finally
            {
                if (enumMonikers != null)
                {
                    Marshal.ReleaseComObject(enumMonikers);
                }

                if (rot != null)
                {
                    Marshal.ReleaseComObject(rot);
                }

                if (bindCtx != null)
                {
                    Marshal.ReleaseComObject(bindCtx);
                }
            }

            return dte;
        }

        [STAThread]
        static DTE CreateNewRunningVSWithOurSolution(string vsPath, string solutionPath, bool fileOp)
        {
            if (!File.Exists(vsPath))
            {
                return null;
            }

            System.Diagnostics.Process devenv = System.Diagnostics.Process.Start(vsPath, QuotePathIfNeeded(solutionPath));


            DTE dte = null;
            do
            {
                System.Threading.Thread.Sleep(2000);
                dte = FindRunningVSWithOurProcess(devenv.Id);
            } while (dte == null);
            MessageFilter.Register();
            do
            {
                System.Threading.Thread.Sleep(1000);
            } while (dte.ItemOperations == null);
            return dte;
        }

        private static string QuotePathIfNeeded(string path)
        {
            if (!path.Contains(" "))
            {
                return path;
            }
            return "\"" + path + "\"";
        }

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        private static void OpenSolutionInVS(DTE dte)
        {
            if (dte == null)
            {
                return;
            }
            dte.MainWindow.Activate();
            SetForegroundWindow(dte.MainWindow.HWnd);
        }

        private static void HaveRunningVSOpenFile(DTE dte, string filePath, int fileLine)
        {
            OpenSolutionInVS(dte);

            dte.ItemOperations.OpenFile(filePath);
            var textSelection = dte.ActiveDocument.Selection as TextSelection;
            textSelection.GotoLine(fileLine, true);
        }

        private static void AttachDebugger(DTE dte, int processID)
        {
            IEnumerable<Process> processes = dte.Debugger.LocalProcesses.OfType<Process>();
            Process process = processes.SingleOrDefault(x => x.ProcessID == processID);
            if (process != null)
            {
                process.Attach();
            }
        }

        private static void StopDebugger(DTE dte)
        {
            dte.Debugger.Stop();
        }

        private static void  RebuildSolution(string vsPath, string solutionPath, string buildConfig)
        {
            var arg = $"{QuotePathIfNeeded(solutionPath)} /Rebuild {buildConfig}";
            BuildSolutionInternal(vsPath, arg);
        }

        private static void BuildSolution(string vsPath, string solutionPath, string buildConfig)
        {
            var arg = $"{QuotePathIfNeeded(solutionPath)} /Build {buildConfig}";
            BuildSolutionInternal(vsPath, arg);
        }

        private static void BuildSolutionInternal(string vsPath, string arguemnts)
        {
            var processInfo = new System.Diagnostics.ProcessStartInfo(vsPath, arguemnts);
            processInfo.CreateNoWindow = true;
            processInfo.UseShellExecute = false;
            processInfo.RedirectStandardError = true;
            processInfo.RedirectStandardOutput = true;

            var devenv = System.Diagnostics.Process.Start(processInfo);
            devenv.ErrorDataReceived += (object sender, System.Diagnostics.DataReceivedEventArgs e) =>
            {
                if (!string.IsNullOrEmpty(e.Data))
                {
                    Console.WriteLine($"Debug>> {e.Data}");
                }
            };
            devenv.BeginErrorReadLine();

            devenv.OutputDataReceived += (object sender, System.Diagnostics.DataReceivedEventArgs e) =>
            {
                if (!string.IsNullOrEmpty(e.Data))
                {
                    Console.WriteLine($"Debug>> {e.Data}");
                }
            };
            devenv.BeginOutputReadLine();
            devenv.WaitForExit();
        }
    }
}
