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
using System.Collections.Generic;
using System.Linq;
using EnvDTE;

namespace Careerbuilder.TrueGoTo
{
    public class SolutionNavigator
    {
        private static SolutionNavigator singleton;

        private readonly vsCMElement[] _blackList;
        private List<CodeElement> _elements;
        private bool _isNavigated;

        private SolutionNavigator() 
        {
            _blackList = new vsCMElement[] { vsCMElement.vsCMElementImportStmt, vsCMElement.vsCMElementUsingStmt, vsCMElement.vsCMElementAttribute, vsCMElement.vsCMElementParameter };
            _elements = new List<CodeElement>();
            _isNavigated = false;
        }

        #region "Static Functions"
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
            SolutionNavigator.getInstance().IsNavigated = true;
            return singleton.Elements;
        }

        public static void AddElement(CodeElement toAdd)
        {
            if (!SolutionNavigator.getInstance().BlackList.Contains(toAdd.Kind))
                SolutionNavigator.getInstance().Elements.Add(toAdd);
        }

        public static void RemoveElement(CodeElement toRemove)
        {
            if (SolutionNavigator.getInstance().Elements.Contains(toRemove))
                SolutionNavigator.getInstance().Elements.Remove(toRemove);
        }
        #endregion

        #region "Properties"
        public vsCMElement[] BlackList
        {
            get
            {
                return _blackList;
            }
        }

        public List<CodeElement> Elements
        {
            get
            {
                return _elements;
            }
        }

        public bool IsNavigated
        {
            get
            {
                return _isNavigated;
            }

            set
            {
                _isNavigated = value && _elements.Count > 0;
            }
        }
        #endregion

        #region "Singleton Functions"
        private void NavigateProjects(Projects projects)
        {
            List<CodeElement> types = new List<CodeElement>();

            foreach (Project p in projects)
            {
                types.AddRange(NavigateProjectItems(p.ProjectItems));
            }

            _elements = types;
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
    #endregion
}
