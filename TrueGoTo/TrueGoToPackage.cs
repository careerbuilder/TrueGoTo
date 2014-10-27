using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

namespace Careerbuilder.TrueGoTo
{
    [PackageRegistration(UseManagedResourcesOnly = true)]
    [ProvideAutoLoad(UIContextGuids80.SolutionHasSingleProject)]
    [ProvideAutoLoad(UIContextGuids80.SolutionHasMultipleProjects)]
    [InstalledProductRegistration("#110", "#112", "1.1", IconResourceID = 400)]
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
            string startWord = HelperElves.GetWordFromSelection((TextSelection)_dte.ActiveDocument.Selection);
            _dte.ExecuteCommand("Edit.GoToDefinition");
            string name = _dte.ActiveDocument.Name;
            string elementPath = name.Substring(0, name.Length - 3);
            name = _dte.ActiveDocument.ActiveWindow.Caption;
            CodeElement targetElement = null;
            if (name.Contains("from metadata"))
            {
                targetElement = HelperElves.ReduceResultSet(_dte, SolutionNavigator.getInstance().Elements, elementPath, startWord);
            }
            if (targetElement != null)
            {
                _dte.ActiveWindow.Close();
                ChangeActiveDocument(targetElement);
            }
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