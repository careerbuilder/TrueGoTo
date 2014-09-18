using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;

namespace Careerbuilder.TrueGoTo
{
    /// <summary>
    /// This is the class that implements the package exposed by this assembly.
    ///
    /// The minimum requirement for a class to be considered a valid package for Visual Studio
    /// is to implement the IVsPackage interface and register itself with the shell.
    /// This package uses the helper classes defined inside the Managed Package Framework (MPF)
    /// to do it: it derives from the Package class that provides the implementation of the 
    /// IVsPackage interface and uses the registration attributes defined in the framework to 
    /// register itself and its components with the shell.
    /// </summary>
    // This attribute tells the PkgDef creation utility (CreatePkgDef.exe) that this class is
    // a package.
    [PackageRegistration(UseManagedResourcesOnly = true)]
    // This attribute is used to register the information needed to show this package
    // in the Help/About dialog of Visual Studio.
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)]
    // This attribute is needed to let the shell know that this package exposes some menus.
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [Guid(GuidList.guidTrueGoToPkgString)]
    public sealed class TrueGoToPackage : Package
    {
        private DTE2 _DTE;
        private CodeModelEvents _codeEvents;

        /// <summary>
        /// Default constructor of the package.
        /// Inside this method you can place any initialization code that does not require 
        /// any Visual Studio service because at this point the package object is created but 
        /// not sited yet inside Visual Studio environment. The place to do all the other 
        /// initialization is the Initialize method.
        /// </summary>
        public TrueGoToPackage()
        {
        }
        
        #region Overridden Package Members

        /// <summary>
        /// Initialization of the package; this method is called right after the package is sited, so this is the place
        /// where you can put all the initialization code that rely on services provided by VisualStudio.
        /// </summary>
        protected override void Initialize()
        {
            base.Initialize();

            _DTE = (DTE2)GetService(typeof(DTE));

            // Add our command handlers for menu (commands must exist in the .vsct file)
            OleMenuCommandService mcs = GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (null != mcs)
            {
                // Create the command for the menu item.
                CommandID menuCommandID = new CommandID(GuidList.guidTrueGoToCmdSet, (int)PkgCmdIDList.cmdTrueGoTo);
                MenuCommand menuItem = new MenuCommand(MenuItemCallback, menuCommandID);
                mcs.AddCommand(menuItem);
            }
        }
        #endregion

        /// <summary>
        /// This function is the callback used to execute a command when the a menu item is clicked.
        /// See the Initialize method to see how the menu item is associated to this function using
        /// the OleMenuCommandService service and the MenuCommand class.
        /// </summary>
        private void MenuItemCallback(object sender, EventArgs e)
        {
            if (_DTE.Solution.IsOpen && _DTE.ActiveDocument != null && _DTE.ActiveDocument.Selection != null)
            {
                EnvDTE.TextSelection selectedText = (EnvDTE.TextSelection)_DTE.ActiveDocument.Selection;
                string target = selectedText.Text;
                CodeElement targetElement = selectedText.ActivePoint.get_CodeElement(vsCMElement.vsCMElementDeclareDecl);
                if (!String.IsNullOrWhiteSpace(target))
                {
                    List<CodeElement> codeElements = getAllTypesInSolution(_DTE.Solution.Projects);
                    List<CodeElement> targetElements = codeElements.Where(t => t.get_FullName().Contains(target)).ToList();
                    //_applicationObject.ExecuteCommand("Edit.NavigateTo");
                    //_applicationObject.ExecuteCommand("Edit.GoToDefinition");
                }
                return;
            }
        }

        private List<CodeElement> getAllTypesInSolution(Projects items)
        {
            List<CodeElement> types = new List<CodeElement>();
            foreach (Project p in items)
            {
                types.AddRange(getAllTypesInSolution(p.ProjectItems));
            }
            return types;
        }

        private CodeElement[] getAllTypesInSolution(ProjectItems items)
        {
            List<CodeElement> elements = new List<CodeElement>();
            foreach (ProjectItem p in items)
            {
                if (p.ContainingProject.CodeModel != null)
                {
                    foreach (CodeElement c in p.ContainingProject.CodeModel.CodeElements)
                    {
                        //if (c.Kind == vsCMElement.vsCMElementDeclareDecl)
                        //{
                        elements.Add(c);
                        //}
                    }
                }
            }
            return elements.ToArray();
        }

        private void AddHandlers()
        {
            EnvDTE80.Events2 events2;
            events2 = (EnvDTE80.Events2)_DTE.Events;
            _codeEvents = events2.get_CodeModelEvents();

            _codeEvents.ElementAdded += new _dispCodeModelEvents_ElementAddedEventHandler(AddedEventHandler);
            _codeEvents.ElementChanged += new _dispCodeModelEvents_ElementChangedEventHandler(ChangedEventHandler);
            _codeEvents.ElementDeleted += new _dispCodeModelEvents_ElementDeletedEventHandler(DeletedEventHandler);
        }

        private void AddedEventHandler(CodeElement Element)
        {
            throw new NotImplementedException();
        }

        private void ChangedEventHandler(CodeElement Element, vsCMChangeKind Change)
        {
            throw new NotImplementedException();
        }

        private void DeletedEventHandler(object Parent, CodeElement Element)
        {
            throw new NotImplementedException();
        }

        private void RemoveHandlers()
        {
            _codeEvents.ElementAdded -= new _dispCodeModelEvents_ElementAddedEventHandler(AddedEventHandler);
            _codeEvents.ElementChanged -= new _dispCodeModelEvents_ElementChangedEventHandler(ChangedEventHandler);
            _codeEvents.ElementDeleted -= new _dispCodeModelEvents_ElementDeletedEventHandler(DeletedEventHandler);
        }
    }
}
