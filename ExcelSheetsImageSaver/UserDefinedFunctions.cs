using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelDna.Integration;

namespace pliExcelSheetsImageSaver
{
        public static class MyFunctions
        {
            [ExcelFunction(Name="HelloWorld",Description = "My first .NET function",Category ="senseless")]
            public static string HelloDna(string name)
            {
                return "Hello " + name;
            }
        }
}