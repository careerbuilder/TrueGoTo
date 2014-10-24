using System.Collections.Generic;
using System.Linq;
using EnvDTE;

namespace Careerbuilder.TrueGoTo
{
    public class SolutionNavigator
    {
        private static SolutionNavigator singleton;

        private readonly vsCMElement[] _blackList = new vsCMElement[] { vsCMElement.vsCMElementImportStmt, vsCMElement.vsCMElementUsingStmt, vsCMElement.vsCMElementAttribute, vsCMElement.vsCMElementParameter };
        private List<CodeElement> elements = new List<CodeElement>();

        private SolutionNavigator() {}

        public static SolutionNavigator getInstance()
        {
            if (singleton == null)
            {
                singleton = new SolutionNavigator();
            }
            return singleton;
        }

        public static List<CodeElement> Navigate(Projects projects)
        {
            SolutionNavigator.getInstance().NavigateProjects(projects);
            return singleton.Elements;
        }

        public List<CodeElement> Elements
        {
            get
            {
                return elements;
            }
        }

        public void AddElement(CodeElement toAdd)
        {
            elements.Add(toAdd);
        }

        public void RemoveElement(CodeElement toRemove)
        {
            if (elements.Contains(toRemove))
                elements.Remove(toRemove);
        }

        private void NavigateProjects(Projects projects)
        {
            List<CodeElement> types = new List<CodeElement>();

            foreach (Project p in projects)
            {
                types.AddRange(NavigateProjectItems(p.ProjectItems));
            }

            elements = types;
        }

        private CodeElement[] NavigateProjectItems(ProjectItems items)
        {
            List<CodeElement> codeElements = new List<CodeElement>();

            if (items != null)
            {
                foreach (ProjectItem item in items)
                {
                    if (item.SubProject != null)
                        codeElements.AddRange(NavigateProjectItems(item.SubProject.ProjectItems));
                    else
                        codeElements.AddRange(NavigateProjectItems(item.ProjectItems));
                    if (item.FileCodeModel != null)
                        codeElements.AddRange(NavigateCodeElements(item.FileCodeModel.CodeElements));
                }
            }

            return codeElements.ToArray();
        }

        private CodeElement[] NavigateCodeElements(CodeElements elements)
        {
            List<CodeElement> codeElements = new List<CodeElement>();
            CodeElements members = null;

            if (elements != null)
            {
                foreach (CodeElement element in elements)
                {
                    if (element.Kind != vsCMElement.vsCMElementDelegate)
                    {
                        members = GetMembers(element);
                        if (members != null)
                            codeElements.AddRange(NavigateCodeElements(members));
                    }

                    if (!_blackList.Contains(element.Kind))
                        codeElements.Add(element);
                }
            }

            return codeElements.ToArray();
        }

        private CodeElements GetMembers(CodeElement element)
        {
            if (element is CodeNamespace)
                return ((CodeNamespace)element).Members;
            else if (element is CodeType)
                return ((CodeType)element).Members;
            else if (element is CodeFunction)
                return ((CodeFunction)element).Parameters;
            else
                return null;
        }
    }
}
