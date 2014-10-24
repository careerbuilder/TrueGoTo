using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

namespace Careerbuilder.TrueGoTo
{
    public class SolutionListener : IVsSolutionEvents, IDisposable
    {
        private IVsSolution activeSolution;
        private uint cookie;

        public event Action AfterSolutionLoaded;
        public event Action BeforeSolutionClosed;

        public SolutionListener(IVsSolution solution, Projects projects)
        {
            AfterSolutionLoaded += () => { SolutionNavigator.Navigate(projects); };
            BeforeSolutionClosed += () => { };

            activeSolution = solution;
            if (activeSolution != null)
            {
                activeSolution.AdviseSolutionEvents(this, out cookie);
            }
        }

        int IVsSolutionEvents.OnAfterCloseSolution(object pUnkReserved)
        { return VSConstants.S_OK; }

        int IVsSolutionEvents.OnAfterLoadProject(IVsHierarchy pStubHierarchy, IVsHierarchy pRealHierarchy)
        { return VSConstants.S_OK; }

        int IVsSolutionEvents.OnAfterOpenProject(IVsHierarchy pHierarchy, int fAdded)
        { return VSConstants.S_OK; }

        int IVsSolutionEvents.OnAfterOpenSolution(object pUnkReserved, int fNewSolution)
        { AfterSolutionLoaded(); return VSConstants.S_OK; }

        int IVsSolutionEvents.OnBeforeCloseProject(IVsHierarchy pHierarchy, int fRemoved)
        { return VSConstants.S_OK; }

        int IVsSolutionEvents.OnBeforeCloseSolution(object pUnkReserved)
        { BeforeSolutionClosed(); return VSConstants.S_OK; }

        int IVsSolutionEvents.OnBeforeUnloadProject(IVsHierarchy pRealHierarchy, IVsHierarchy pStubHierarchy)
        { return VSConstants.S_OK; }

        int IVsSolutionEvents.OnQueryCloseProject(IVsHierarchy pHierarchy, int fRemoving, ref int pfCancel)
        { return VSConstants.S_OK; }

        int IVsSolutionEvents.OnQueryCloseSolution(object pUnkReserved, ref int pfCancel)
        { return VSConstants.S_OK; }

        int IVsSolutionEvents.OnQueryUnloadProject(IVsHierarchy pRealHierarchy, ref int pfCancel)
        { return VSConstants.S_OK; }

        public void Dispose()
        {
            if (activeSolution != null && cookie != 0)
            {
                GC.SuppressFinalize(this);
                activeSolution.UnadviseSolutionEvents(cookie);
                AfterSolutionLoaded = null;
                BeforeSolutionClosed = null;
                cookie = 0;
                activeSolution = null;
            }
        }
    }
}