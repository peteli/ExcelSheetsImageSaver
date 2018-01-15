using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfceExcel = Microsoft.Office.Interop.Excel;
using ExcelDna.Integration;
using sysImg = System.Drawing;
using System.Windows.Forms;
using System.Diagnostics;


namespace pliExcelSheetsImageSaver
{
    public static class MyApp
    {
        public static OfceExcel.Application ExcelApp { get; set; }

        public static bool SaveImage(string fullpathname, sysImg.Imaging.ImageFormat imageFormat)
        {
            if (Clipboard.ContainsImage())
            {
                   sysImg.Image image = Clipboard.GetImage();
                    image.Save(fullpathname, sysImg.Imaging.ImageFormat.Png);
                return true;
            }
            else if (Clipboard.ContainsData(DataFormats.MetafilePict))
            {
                sysImg.Image what = (sysImg.Image)Clipboard.GetData(DataFormats.MetafilePict);

                Debug.WriteLine(what.GetType());
                
                //image.Save(fullpathname, sysImg.Imaging.ImageFormat.Png);
                return true;
            } else
            {
                Debug.WriteLine("NO IMAGE");
                return false;
            }
        }
        public static string Fullpathname { get; set; }

        public static void SaveWithWB(string filename)
        {
            MyApp.Fullpathname = MyApp.ExcelApp.ActiveWorkbook.Path + "\\" + filename + ".png";
            sysImg.Imaging.ImageFormat PngFormat = sysImg.Imaging.ImageFormat.Png;
            bool result = MyApp.SaveImage(MyApp.Fullpathname, PngFormat);
        }
        public static void ExportNamedRanges()
        {
            OfceExcel.Workbook WB = MyApp.ExcelApp.ActiveWorkbook;

            foreach(OfceExcel.Name NamedRange in WB.Names)
            {
                try
                {
                    OfceExcel.Range CurRange = NamedRange.RefersToRange;
                    CurRange.CopyPicture(OfceExcel.XlPictureAppearance.xlScreen, OfceExcel.XlCopyPictureFormat.xlBitmap);
                    MyApp.Fullpathname = MyApp.ExcelApp.ActiveWorkbook.Path + "\\" + WB.Name + "_" + NamedRange.Name + ".png";
                    sysImg.Imaging.ImageFormat PngFormat = sysImg.Imaging.ImageFormat.Png;
                    bool result = MyApp.SaveImage(MyApp.Fullpathname, PngFormat);

                }
                catch { }
                finally { }

            }



            foreach (OfceExcel.Worksheet WS in WB.Worksheets)
            {
                foreach (OfceExcel.Name NamedRange in WS.Names)
                {

                }
            }


        }

    }

    /*        If Not System.Windows.Forms.Clipboard.GetDataObject() Is Nothing Then
                Dim oDataObj As IDataObject = System.Windows.Forms.Clipboard.GetDataObject()
                If oDataObj.GetDataPresent(System.Windows.Forms.DataFormats.Bitmap) Then
                    Dim oImgObj As System.Drawing.Image = oDataObj.GetData(DataFormats.Bitmap, True)
                    'To Save as Bitmap
                    oImgObj.Save("c:\Test.bmp", System.Drawing.Imaging.ImageFormat.Bmp)
                    'To Save as Jpeg
                    oImgObj.Save("c:\Test.jpeg", System.Drawing.Imaging.ImageFormat.Jpeg)
                    'To Save as Gif
                    oImgObj.Save("c:\Test.gif", System.Drawing.Imaging.ImageFormat.Gif)
                End If
            End If
        End Sub
    */
}
