using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using EnvDTE;
using EnvDTE80;

namespace Careerbuilder.TrueGoTo
{
    public class HelperElves
    {
        public static IEnumerable<T> ConvertToElementArray<T>(IEnumerable list)
        {
            foreach (T element in list)
                yield return element;
        }

        public static string GetWordFromSelection(TextSelection selection)
        {
            string target = selection.Text ?? String.Empty;
            int line = selection.TopPoint.Line;
            int offset = selection.TopPoint.LineCharOffset;

            selection.WordLeft(true);
            string leftWord = selection.Text;
            selection.WordRight(true);
            string rightWord = selection.Text;

            if (!(String.IsNullOrWhiteSpace(leftWord) || String.IsNullOrWhiteSpace(rightWord)))
            {
                string selectedWord = leftWord + rightWord;
                if (String.IsNullOrWhiteSpace(target) || Regex.Match(selectedWord, target, RegexOptions.IgnoreCase).Success)
                {
                    ResetSelection(selection, line, offset, target.Count());
                    return selectedWord.Trim();
                }
            }

            ResetSelection(selection, line, offset, target.Count());
            return target;
        }

        private static void ResetSelection(TextSelection selection, int lineNumber, int offset, int length)
        {
            selection.MoveToLineAndOffset(lineNumber, offset);
            selection.CharRight(true, length);
        }

        public static CodeElement ReduceResultSet(DTE2 dte, List<CodeElement> elements, string targetPath, string targetName)
        {
            List<CodeElement> codeElements = new List<CodeElement>();
            List<string> activeNamespaces = new List<string>();
            vsCMElement[] whiteList = new vsCMElement[] { vsCMElement.vsCMElementImportStmt, vsCMElement.vsCMElementUsingStmt, vsCMElement.vsCMElementIncludeStmt };
            string target = targetPath.ToLower() + targetName .ToLower();
            
            foreach (CodeElement ele in elements)
            {
                if (ele.FullName.ToLower() == targetPath.ToLower() || ele.FullName.ToLower() == target)
                    codeElements.Add(ele);
            }            

            if (codeElements != null && codeElements.Count > 0)
            {
                if (codeElements.Count == 1)
                    return codeElements[0];
                activeNamespaces = ConvertToElementArray<CodeElement>(dte.ActiveDocument.ProjectItem.FileCodeModel.CodeElements)
                    .Where(e => whiteList.Contains(e.Kind)).Select(e => ((CodeImport)e).Namespace).ToList();
                return HandleFunctionResultSet(codeElements.Where(e => activeNamespaces.Any(a => e.FullName.Contains(a))));
            }
            return null;
        }
        
        public static CodeElement HandleFunctionResultSet(IEnumerable<CodeElement> elements)
        {
            if (elements.All(e => e.Kind != vsCMElement.vsCMElementFunction))
                return elements.FirstOrDefault();
            else
            {
                // Need to fix this!
                return elements.FirstOrDefault();
            }
        }
    }
}
