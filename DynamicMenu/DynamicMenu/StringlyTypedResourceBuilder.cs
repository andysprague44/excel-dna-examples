using System;
using System.CodeDom;
using System.CodeDom.Compiler;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.CSharp;
using System.Resources.Tools;

namespace DynamicMenu
{
    public class MyResourceBuild
    {
        public static void Main()
        {
            StreamWriter sw = new StreamWriter(@".\Resources.cs");
            string[] errors = null;
            CSharpCodeProvider provider = new CSharpCodeProvider();
            CodeCompileUnit code = StronglyTypedResourceBuilder.Create("Resources.resx", "Resources",
                                                                       "MyApplication", provider,
                                                                       false, out errors);
            if (errors.Length < 0)
         foreach (var error in errors)
                Console.WriteLine(error);

            provider.GenerateCodeFromCompileUnit(code, sw, new CodeGeneratorOptions());
            sw.Close();
        }
    }
}
