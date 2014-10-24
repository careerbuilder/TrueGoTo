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
    [ProvideAutoLoad(UIContextGuids80.SolutionExists)]
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)]
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [Guid(GuidList.guidTrueGoToPkgString)]
    public sealed class TrueGoToPackage : Package
    {
        private DTE2 _dte;
        private CodeModelEvents _codeEvents;
        private SolutionListener _solutionEvents;

        public TrueGoToPackage() { }

        protected override void Initialize()
        {
            base.Initialize();
            _dte = (DTE2)GetService(typeof(DTE));
            _solutionEvents = new SolutionListener(GetService(typeof(SVsSolution)) as IVsSolution, _dte.Solution.Projects);

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
            if (_dte.Solution.IsOpen && _dte.ActiveDocument != null && _dte.ActiveDocument.Selection != null)
            {
                if (!(SolutionNavigator.getInstance().Elements.Count > 0))
                {
                    SolutionNavigator.Navigate(_dte.Solution.Projects);
                }
                HackThatDef();
                return;
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

        private void AddHandlers()
        {
            EnvDTE80.Events2 events2;
            events2 = (EnvDTE80.Events2)_dte.Events;
            _codeEvents = events2.get_CodeModelEvents();

            _codeEvents.ElementAdded += new _dispCodeModelEvents_ElementAddedEventHandler(AddedEventHandler);
            _codeEvents.ElementChanged += new _dispCodeModelEvents_ElementChangedEventHandler(ChangedEventHandler);
            _codeEvents.ElementDeleted += new _dispCodeModelEvents_ElementDeletedEventHandler(DeletedEventHandler);
        }

        private void AddedEventHandler(CodeElement Element)
        {
            SolutionNavigator.getInstance().AddElement(Element);
        }

        private void ChangedEventHandler(CodeElement Element, vsCMChangeKind Change)
        {
            if (Change == vsCMChangeKind.vsCMChangeKindRename || Change == vsCMChangeKind.vsCMChangeKindUnknown)
            {
                SolutionNavigator.getInstance().AddElement(Element);
            }
        }

        private void DeletedEventHandler(object Parent, CodeElement Element)
        {
            SolutionNavigator.getInstance().RemoveElement(Element);
        }

        private void RemoveHandlers()
        {
            _codeEvents.ElementAdded -= new _dispCodeModelEvents_ElementAddedEventHandler(AddedEventHandler);
            _codeEvents.ElementChanged -= new _dispCodeModelEvents_ElementChangedEventHandler(ChangedEventHandler);
            _codeEvents.ElementDeleted -= new _dispCodeModelEvents_ElementDeletedEventHandler(DeletedEventHandler);
        }
    }
}