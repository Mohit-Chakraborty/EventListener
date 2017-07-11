using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace VSIXProject
{
    internal sealed class SolutionEventHandler : IVsSolutionEvents, IVsSolutionLoadEvents
    {
        private const string LslwindowPaneCaption = "Lightweight solution load";
        private IVsOutputWindowPane windowPane;
        private string windowPaneData;
        private Package package;

        internal SolutionEventHandler(Package package)
        {
            this.package = package;
        }

        public int OnBeforeOpenSolution(string pszSolutionFilename)
        {
            this.WriteToOutputPane("OnBeforeOpenSolution");
            this.WriteToOutputPane("\t" + pszSolutionFilename);

            return VSConstants.S_OK;
        }

        public int OnBeforeBackgroundSolutionLoadBegins()
        {
            this.WriteToOutputPane("OnBeforeBackgroundSolutionLoadBegins");
            return VSConstants.S_OK;
        }

        public int OnQueryBackgroundLoadProjectBatch(out bool pfShouldDelayLoadToNextIdle)
        {
            pfShouldDelayLoadToNextIdle = false;

            this.WriteToOutputPane("OnQueryBackgroundLoadProjectBatch");
            return VSConstants.S_OK;
        }

        public int OnBeforeLoadProjectBatch(bool fIsBackgroundIdleBatch)
        {
            this.WriteToOutputPane("OnBeforeLoadProjectBatch");
            return VSConstants.S_OK;
        }

        public int OnAfterLoadProjectBatch(bool fIsBackgroundIdleBatch)
        {
            this.WriteToOutputPane("OnAfterLoadProjectBatch");
            return VSConstants.S_OK;
        }

        public int OnAfterBackgroundSolutionLoadComplete()
        {
            this.WriteToOutputPane("OnAfterBackgroundSolutionLoadComplete");
            this.RealWriteToOutputPane(this.windowPaneData);
            this.windowPaneData = null;

            return VSConstants.S_OK;
        }

        public int OnAfterOpenProject(IVsHierarchy pHierarchy, int fAdded)
        {
            this.WriteToOutputPane("OnAfterOpenProject");

            //int hr = pHierarchy.GetProperty((uint)VSConstants.VSITEMID.Root, (int)__VSHPROPID.VSHPROPID_Caption, out object caption);
            //int hr = pHierarchy.GetProperty((uint)VSConstants.VSITEMID.Root, (int)__VSHPROPID.VSHPROPID_Name, out object caption);
            //this.WriteToOutputPane("\t" +  caption);

            int hr = pHierarchy.GetCanonicalName((uint)VSConstants.VSITEMID.Root, out string name);
            this.WriteToOutputPane("\t" + name);

            hr = pHierarchy.GetProperty((uint)VSConstants.VSITEMID.Root, (int)__VSHPROPID9.VSHPROPID_IsDeferred, out object isDeferred);
            if (ErrorHandler.Succeeded(hr))
            {
                this.WriteToOutputPane("\tDeferred: " + (bool)isDeferred);
            }

            return VSConstants.S_OK;
        }

        public int OnQueryCloseProject(IVsHierarchy pHierarchy, int fRemoving, ref int pfCancel)
        {
            this.WriteToOutputPane("OnQueryCloseProject");

            int hr = pHierarchy.GetProperty((uint)VSConstants.VSITEMID.Root, (int)__VSHPROPID.VSHPROPID_Caption, out object caption);
            this.WriteToOutputPane("\t" + caption);

            return VSConstants.S_OK;
        }

        public int OnBeforeCloseProject(IVsHierarchy pHierarchy, int fRemoved)
        {
            this.WriteToOutputPane("OnBeforeCloseProject");

            int hr = pHierarchy.GetProperty((uint)VSConstants.VSITEMID.Root, (int)__VSHPROPID.VSHPROPID_Caption, out object caption);
            this.WriteToOutputPane("\t" + caption);

            return VSConstants.S_OK;
        }

        public int OnAfterLoadProject(IVsHierarchy pStubHierarchy, IVsHierarchy pRealHierarchy)
        {
            this.WriteToOutputPane("OnAfterLoadProject");

            int hr = pStubHierarchy.GetProperty((uint)VSConstants.VSITEMID.Root, (int)__VSHPROPID.VSHPROPID_Caption, out object caption);
            this.WriteToOutputPane("\t" + caption);

            hr = pRealHierarchy.GetProperty((uint)VSConstants.VSITEMID.Root, (int)__VSHPROPID.VSHPROPID_Caption, out caption);
            this.WriteToOutputPane("\t" + caption);

            return VSConstants.S_OK;
        }

        public int OnQueryUnloadProject(IVsHierarchy pRealHierarchy, ref int pfCancel)
        {
            this.WriteToOutputPane("OnQueryUnloadProject");

            int hr = pRealHierarchy.GetProperty((uint)VSConstants.VSITEMID.Root, (int)__VSHPROPID.VSHPROPID_Caption, out object caption);
            this.WriteToOutputPane("\t" + caption);

            return VSConstants.S_OK;
        }

        public int OnBeforeUnloadProject(IVsHierarchy pRealHierarchy, IVsHierarchy pStubHierarchy)
        {
            this.WriteToOutputPane("OnBeforeUnloadProject");

            int hr = pRealHierarchy.GetProperty((uint)VSConstants.VSITEMID.Root, (int)__VSHPROPID.VSHPROPID_ProjectName, out object caption);
            this.WriteToOutputPane("\t" + caption);

            return VSConstants.S_OK;
        }

        public int OnAfterOpenSolution(object pUnkReserved, int fNewSolution)
        {
            this.WriteToOutputPane("OnAfterOpenSolution");
            return VSConstants.S_OK;
        }

        public int OnQueryCloseSolution(object pUnkReserved, ref int pfCancel)
        {
            this.WriteToOutputPane("OnQueryCloseSolution");
            return VSConstants.S_OK;
        }

        public int OnBeforeCloseSolution(object pUnkReserved)
        {
            this.WriteToOutputPane("OnBeforeCloseSolution");
            return VSConstants.S_OK;
        }

        public int OnAfterCloseSolution(object pUnkReserved)
        {
            this.WriteToOutputPane("OnAfterCloseSolution");
            this.RealWriteToOutputPane(this.windowPaneData);
            this.windowPaneData = null;

            return VSConstants.S_OK;
        }

        private void WriteToOutputPane(string data)
        {
            this.windowPaneData += data + System.Environment.NewLine;
        }

        private void RealWriteToOutputPane(string data)
        {
            if (this.windowPane == null)
            {
                this.windowPane = this.package.GetOutputPane(VSConstants.OutputWindowPaneGuid.GeneralPane_guid, LslwindowPaneCaption);
            }

            int? hr = this.windowPane?.Activate();
            //hr = this.windowPane?.OutputString(data + System.Environment.NewLine);
            hr = this.windowPane?.OutputStringThreadSafe(data + System.Environment.NewLine);

            MessageBox.Show(data + System.Environment.NewLine);
        }
    }
}
