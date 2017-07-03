using Microsoft.CSharp;
using System;
using System.CodeDom;
using System.CodeDom.Compiler;
using System.Drawing;
using System.IO;
using System.Resources;
using System.Resources.Tools;

namespace CustomIconAddIn
{
    public class MyResourceBuilder
    {
        public static void Main()
        {
            StreamWriter sw = new StreamWriter(@".\Resources.cs");
            string[] errors = null;
            CSharpCodeProvider provider = new CSharpCodeProvider();
            CodeCompileUnit code = StronglyTypedResourceBuilder.Create(
                "Resources.resx", //resource file
                "Resources", //base
                "CustomIconAddIn", //namespace
                provider,
                false,
                out errors);

            if (errors.Length > 0)
                foreach (var error in errors)
                    Console.WriteLine(error);

            provider.GenerateCodeFromCompileUnit(code, sw, new CodeGeneratorOptions());
            sw.Close();
        }
    }
}
