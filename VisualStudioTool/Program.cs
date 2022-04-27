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
        static void Main(string[] args)
        {
            if (args.Length < 5)
            {
                return;
            }

            bool openFile = args[0].Equals("open", StringComparison.OrdinalIgnoreCase);
            string vsPath = args[1];
            string solutionPath = args[2];

            try
            {
                var dte = FindRunningVSProWithOurSolution(solutionPath);
                if (dte == null)
                {
                    dte = CreateNewRunningVSProWithOurSolution(vsPath, solutionPath);
                }
                if (openFile && int.TryParse(args[4], out int fileLine))
                {
                    string filePath = args[3];
                    HaveRunningVSProOpenFile(dte, filePath, fileLine);
                }
                else if (int.TryParse(args[3], out int processID))
                {
                    AttachDebugger(dte, processID);
                }
            }
            catch (Exception e)
            {
                Console.Write(e.Message);
            }
        }

        [DllImport("ole32.dll")]
        private static extern int CreateBindCtx(uint reserved, out IBindCtx ppbc);

        static DTE FindRunningVSProWithOurSolution(string solutionPath)
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
                    var dte2 = runningObject as DTE;
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

        static DTE FindRunningVSProWithOurProcess(int processId)
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

        static DTE CreateNewRunningVSProWithOurSolution(string vsPath, string solutionPath)
        {
            if (!File.Exists(vsPath))
            {
                return null;
            }

            var devenv = System.Diagnostics.Process.Start(vsPath, solutionPath);

            DTE dte = null;
            do
            {
                System.Threading.Thread.Sleep(2000);
                dte = FindRunningVSProWithOurProcess(devenv.Id);
            }
            while (dte == null);

            do
            {
                System.Threading.Thread.Sleep(1000);
            } while (dte.ItemOperations == null);
            return dte;
        }

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        static void HaveRunningVSProOpenFile(DTE dte, string filePath, int fileLine)
        {
            if (dte == null)
            {
                return;
            }

            dte.MainWindow.Activate();
            SetForegroundWindow(new IntPtr(dte.MainWindow.HWnd));

            var window = dte.ItemOperations.OpenFile(filePath);
            var textSelection = (TextSelection)window.Selection;
            textSelection.GotoLine(fileLine, true);
            Marshal.ReleaseComObject(dte);
        }

        static void AttachDebugger(DTE dte, int processID)
        {
            if (dte == null)
            {
                return;
            }

            IEnumerable<Process> processes = dte.Debugger.LocalProcesses.OfType<Process>();
            var process = processes.SingleOrDefault(x => x.ProcessID == processID);
            if (process != null)
            {
                process.Attach();
            }
        }
    }
}
