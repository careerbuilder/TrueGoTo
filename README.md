TrueGoTo
===========

A GoToDefinition Extension for Visual Studio

This tool will GoToDefinition in Visual Studio in .Net files across languages. 
Today, Visual Studio loads differing languages in the CLR as assemblies and, as such, cannot GoToDefinition between
them. By Installing this extension, you are able to seamlessly move from C# to VB.Net and back again.


###Installation

Download the .vsix package and double click. Follow the on-screen prompts.

###Usage

You have a number of options for using TrueGoTo. First, highlight or select a piece of code, then:
* Press Alt+F1
* Right click and select TrueGoTo
* Select "Tools" in the toolbar and select TrueGoTo

TrueGoTo will then GoToDefinition on the code, opening a file if necessary, or prompting an error if it is not possible.

###Notes
Due to the nature of Visual Studio, TrueGoTo must build up a cache at the time a Solution file is loaded. Because of this,
you may experience longer than normal wait times when first opening code.
