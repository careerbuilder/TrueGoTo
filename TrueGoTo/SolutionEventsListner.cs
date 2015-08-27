/*
 Copyright 2015 CareerBuilder, LLC
Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    http://www.apache.org/licenses/LICENSE-2.0
Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and limitations under the License.
 */

using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell.Interop;

namespace Careerbuilder.TrueGoTo
{
    public sealed class SolutionListener : IVsSolutionEvents
    {
        private uint cookie;
        private DTE2 _dte;
        private CodeModelEvents _codeEvents;

        public SolutionListener(IVsSolution solution, DTE2 dte)
        {
            if (solution != null)
            {
                solution.AdviseSolutionEvents(this, out cookie);
            }
            if (dte != null)
            {
                _dte = dte;
                _codeEvents = ((Events2)dte.Events).get_CodeModelEvents();
            }
        }

        int IVsSolutionEvents.OnAfterCloseSolution(object pUnkReserved)
        { 
            SolutionNavigator.getInstance().IsNavigated = false;
            RemoveCodeElementHandlers();
            return VSConstants.S_OK; 
        }

        int IVsSolutionEvents.OnAfterLoadProject(IVsHierarchy pStubHierarchy, IVsHierarchy pRealHierarchy)
        { return VSConstants.S_OK; }

        int IVsSolutionEvents.OnAfterOpenProject(IVsHierarchy pHierarchy, int fAdded)
        { return VSConstants.S_OK; }

        int IVsSolutionEvents.OnAfterOpenSolution(object pUnkReserved, int fNewSolution)
        { 
            SolutionNavigator.Navigate(_dte.Solution.Projects);
            AddCodeElementHandlers();
            return VSConstants.S_OK; 
        }

        int IVsSolutionEvents.OnBeforeCloseProject(IVsHierarchy pHierarchy, int fRemoved)
        { return VSConstants.S_OK; }

        int IVsSolutionEvents.OnBeforeCloseSolution(object pUnkReserved)
        { return VSConstants.S_OK; }

        int IVsSolutionEvents.OnBeforeUnloadProject(IVsHierarchy pRealHierarchy, IVsHierarchy pStubHierarchy)
        { return VSConstants.S_OK; }

        int IVsSolutionEvents.OnQueryCloseProject(IVsHierarchy pHierarchy, int fRemoving, ref int pfCancel)
        { return VSConstants.S_OK; }

        int IVsSolutionEvents.OnQueryCloseSolution(object pUnkReserved, ref int pfCancel)
        { return VSConstants.S_OK; }

        int IVsSolutionEvents.OnQueryUnloadProject(IVsHierarchy pRealHierarchy, ref int pfCancel)
        { return VSConstants.S_OK; }

        private void AddCodeElementHandlers()
        {
            _codeEvents.ElementAdded += new _dispCodeModelEvents_ElementAddedEventHandler(AddedEventHandler);
            _codeEvents.ElementChanged += new _dispCodeModelEvents_ElementChangedEventHandler(ChangedEventHandler);
            _codeEvents.ElementDeleted += new _dispCodeModelEvents_ElementDeletedEventHandler(DeletedEventHandler);
        }

        private void AddedEventHandler(CodeElement Element)
        {
            SolutionNavigator.AddElement(Element);
        }

        private void ChangedEventHandler(CodeElement Element, vsCMChangeKind Change)
        {
            if (Change == vsCMChangeKind.vsCMChangeKindRename || Change == vsCMChangeKind.vsCMChangeKindUnknown)
            {
                SolutionNavigator.AddElement(Element);
            }
        }

        private void DeletedEventHandler(object Parent, CodeElement Element)
        {
            SolutionNavigator.RemoveElement(Element);
        }

        private void RemoveCodeElementHandlers()
        {
            _codeEvents.ElementAdded -= new _dispCodeModelEvents_ElementAddedEventHandler(AddedEventHandler);
            _codeEvents.ElementChanged -= new _dispCodeModelEvents_ElementChangedEventHandler(ChangedEventHandler);
            _codeEvents.ElementDeleted -= new _dispCodeModelEvents_ElementDeletedEventHandler(DeletedEventHandler);
        }

    }
}