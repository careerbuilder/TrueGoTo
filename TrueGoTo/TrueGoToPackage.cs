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

using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

namespace Careerbuilder.TrueGoTo
{
    [PackageRegistration(UseManagedResourcesOnly = true)]
    [ProvideAutoLoad(UIContextGuids80.SolutionHasSingleProject)]
    [ProvideAutoLoad(UIContextGuids80.SolutionHasMultipleProjects)]
    [InstalledProductRegistration("#110", "#112", "1.2", IconResourceID = 400)]
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [Guid(GuidList.guidTrueGoToPkgString)]
    public sealed class TrueGoToPackage : Package
    {
        private DTE2 _dte;
        
        public TrueGoToPackage() { }

        protected override void Initialize()
        {
            base.Initialize();
            _dte = (DTE2)GetService(typeof(DTE));
            SolutionListener _solutionEvents = new SolutionListener(GetService(typeof(SVsSolution)) as IVsSolution, _dte);

            OleMenuCommandService mcs = GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (null != mcs)
            {
                CommandID menuCommandID = new CommandID(GuidList.guidTrueGoToCmdSet, (int)PkgCmdIDList.cmdTrueGoTo);
                MenuCommand menuItem = new MenuCommand(MenuItemCallback, menuCommandID);
                mcs.AddCommand(menuItem);
            }
        }

        private void MenuItemCallback(object sender, EventArgs e)
        {
            if (SolutionNavigator.getInstance().IsNavigated && _dte.ActiveDocument != null && _dte.ActiveDocument.Selection != null)
            {
                if (!(SolutionNavigator.getInstance().Elements.Count > 0))
                {
                    SolutionNavigator.Navigate(_dte.Solution.Projects);
                }
                HackThatDef();
            }
        }

        private void HackThatDef()
        {
            string startWord = Utilities.GetWordFromSelection((TextSelection)_dte.ActiveDocument.Selection);
            string currentDocument = _dte.ActiveDocument.Name;
            Window objectBrowser = GetObjectBrowser();
            bool OBWasOpen = (objectBrowser != null && objectBrowser.Visible);
            _dte.ExecuteCommand("Edit.GoToDefinition");
            string name = _dte.ActiveDocument.ActiveWindow.Caption;
            string elementPath = String.Empty;
            CodeElement targetElement = null;

            if (name.Equals(currentDocument))   // VB to C# and some reference types
            {
                objectBrowser = GetObjectBrowser();
                if (objectBrowser != null && objectBrowser.Visible)
                {
                    _dte.ExecuteCommand("Edit.Copy");
                    elementPath = Clipboard.GetText();
                    elementPath = elementPath.Substring(0, elementPath.LastIndexOf('.'));

                    if (!OBWasOpen)
                    {
                        objectBrowser.Close();
                    }
                }
                
                targetElement = Utilities.ReduceResultSet(_dte, SolutionNavigator.getInstance().Elements, elementPath, startWord);

                if (targetElement != null)
                {
                    ChangeActiveDocument(targetElement);
                }
                
            }
            else if (name.Contains("from metadata"))    // C# to VB and the rest of the reference types
            {
                elementPath = _dte.ActiveDocument.Name.Substring(0, _dte.ActiveDocument.Name.Length - 3);
                targetElement = Utilities.ReduceResultSet(_dte, SolutionNavigator.getInstance().Elements, elementPath, startWord);
                
                if (targetElement != null)
                {
                    _dte.ActiveWindow.Close();
                    ChangeActiveDocument(targetElement);
                }
            }
        }

        private Window GetObjectBrowser()
        {
            IEnumerable<Window> windows = Utilities.ConvertToElementArray<Window>(_dte.Windows).Where(x => x.Caption.ToLower().Contains("object browser"));
            return windows.FirstOrDefault();
        }

        private void ChangeActiveDocument(CodeElement definingElement)
        {
            Window window = definingElement.ProjectItem.Open(EnvDTE.Constants.vsViewKindCode);
            window.Activate();
            TextSelection currentPoint = window.Document.Selection as TextSelection;
            currentPoint.MoveToDisplayColumn(definingElement.StartPoint.Line, definingElement.StartPoint.DisplayColumn);
        }      
    }
}