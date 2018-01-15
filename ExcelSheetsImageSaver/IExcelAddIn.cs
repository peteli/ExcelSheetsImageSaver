using System.Diagnostics;
using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using ExcelDna.Integration.CustomUI;

namespace pliExcelSheetsImageSaver
{
    public class MyAddIn : IExcelAddIn
    {
        public void AutoOpen()
        {
            // get handle of excel application
            Debug.WriteLine("AutoOpen runs");
            MyApp.ExcelApp = (Application) ExcelDnaUtil.Application;
            RibbonController.Load += (s,e)=> { Debug.WriteLine("Event in MyAddIn class erhalten"); } ;
        }

        public void AutoClose()
        {
            // put code here
        }
    }
}
