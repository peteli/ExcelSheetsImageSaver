using System;
using System.Runtime.InteropServices;
using ExcelDna.Integration.CustomUI;
using myResources = pliExcelSheetsImageSaver.Properties;
using System.Diagnostics;
using ExcelSheetsImageSaver;
using Microsoft.Office.Core;
//using stdole; for  IPictureDisp  only
using System.Drawing;

namespace pliExcelSheetsImageSaver
{
    [ComVisible(true)]
    public partial class RibbonController : ExcelRibbon, IRibbonCallbacks
    {
        #region functions
        public override string GetCustomUI(string RibbonID)
        {
            //MessageBox.Show((string)myResources.Resources.Ribbon1);
            return (string)myResources.Resources.Ribbon1;
            /*          return @"
                  <customUI xmlns='http://schemas.microsoft.com/office/2006/01/customui'>
                  <ribbon>
                    <tabs>
                      <tab id='tab1' label='My Tab'>
                        <group id='group1' label='My Group'>
                          <button id='button1' label='My Button' onAction='OnButtonPressed'/>
                        </group >
                      </tab>
                    </tabs>
                  </ribbon>
                </customUI>";
            */
         }
        
        #endregion
        #region properties
        public static Microsoft.Office.Core.IRibbonUI AppRibbon { get; protected set; }
        #endregion

        #region callbacks
        #endregion
        #region OnButtonPressed
        public static event EventHandler ButtonPressed;
        // ribbon callback
        public void OnButtonPressed(Microsoft.Office.Core.IRibbonControl control)
        {
            //string ExcelWBName = MyApp.ExcelApp.ActiveWorkbook.Name;
            //Debug.WriteLine(ExcelWBName);
            MyApp.ExportNamedRanges();
            ButtonPressedEventArgs args = new ButtonPressedEventArgs { ControlId = control };
            OnButtonPressedEvent(args);
        }
        // raise event
        protected virtual void OnButtonPressedEvent(EventArgs e)
        {
            EventHandler handler = ButtonPressed;
            handler?.Invoke(this, e);
        }

        public string GetAltText(Microsoft.Office.Core.IRibbonControl control)
        {
            throw new NotImplementedException();
        }

        public string GetContent(Microsoft.Office.Core.IRibbonControl control)
        {
            throw new NotImplementedException();
        }

        public string GetDescription(Microsoft.Office.Core.IRibbonControl control)
        {
            throw new NotImplementedException();
        }

        public bool GetEnabled(Microsoft.Office.Core.IRibbonControl control)
        {
            throw new NotImplementedException();
        }

        public string GetHelperText(Microsoft.Office.Core.IRibbonControl control, int itemIndex)
        {
            throw new NotImplementedException();
        }

        public string GetHelperText(Microsoft.Office.Core.IRibbonControl control)
        {
            throw new NotImplementedException();
        }

        public Bitmap GetImage(Microsoft.Office.Core.IRibbonControl control)
        {
            throw new NotImplementedException();
        }

        public string GetImageString(Microsoft.Office.Core.IRibbonControl control)
        {
            throw new NotImplementedException();
        }

        public string GetImageMso(Microsoft.Office.Core.IRibbonControl control)
        {
            throw new NotImplementedException();
        }

        public int GetItemCount(Microsoft.Office.Core.IRibbonControl control)
        {
            throw new NotImplementedException();
        }

        public string GetItemCountString(Microsoft.Office.Core.IRibbonControl control)
        {
            throw new NotImplementedException();
        }

        public int getItemHeight(Microsoft.Office.Core.IRibbonControl control)
        {
            throw new NotImplementedException();
        }

        public string GetItemID(Microsoft.Office.Core.IRibbonControl control, int itemIndex)
        {
            throw new NotImplementedException();
        }

        public string GetItemID(Microsoft.Office.Core.IRibbonControl control)
        {
            throw new NotImplementedException();
        }

        public Bitmap GetItemImage(Microsoft.Office.Core.IRibbonControl control, int index)
        {
            throw new NotImplementedException();
        }

        public string GetItemLabel(Microsoft.Office.Core.IRibbonControl control, int itemIndex)
        {
            throw new NotImplementedException();
        }

        public string GetItemLabel(Microsoft.Office.Core.IRibbonControl control)
        {
            throw new NotImplementedException();
        }

        public string GetItemScreentip(Microsoft.Office.Core.IRibbonControl control, int itemIndex)
        {
            throw new NotImplementedException();
        }

        public string GetItemSupertip(Microsoft.Office.Core.IRibbonControl control, int itemIndex)
        {
            throw new NotImplementedException();
        }

        public int getItemWidth(Microsoft.Office.Core.IRibbonControl control)
        {
            throw new NotImplementedException();
        }

        public string GetKeytip(Microsoft.Office.Core.IRibbonControl control)
        {
            throw new NotImplementedException();
        }

        public string GetKeytip(Microsoft.Office.Core.IRibbonControl control, int itemIndex)
        {
            throw new NotImplementedException();
        }

        public string GetLabel(Microsoft.Office.Core.IRibbonControl control)
        {
            throw new NotImplementedException();
        }

        public string GetLabel(Microsoft.Office.Core.IRibbonControl control, int itemIndex)
        {
            throw new NotImplementedException();
        }

        public bool GetPressed(Microsoft.Office.Core.IRibbonControl control)
        {
            throw new NotImplementedException();
        }

        public string GetScreentip(Microsoft.Office.Core.IRibbonControl control)
        {
            throw new NotImplementedException();
        }

        public string GetScreentip(Microsoft.Office.Core.IRibbonControl control, int itemIndex)
        {
            throw new NotImplementedException();
        }

        public string GetSelectedItemID(Microsoft.Office.Core.IRibbonControl control, int itemIndex)
        {
            throw new NotImplementedException();
        }

        public int GetSelectedItemIndex(Microsoft.Office.Core.IRibbonControl control)
        {
            throw new NotImplementedException();
        }

        public string GetSelectedItemIndex(Microsoft.Office.Core.IRibbonControl control, int itemIndex)
        {
            throw new NotImplementedException();
        }

        public bool GetShowImage(Microsoft.Office.Core.IRibbonControl control)
        {
            throw new NotImplementedException();
        }

        public bool GetShowLabel(Microsoft.Office.Core.IRibbonControl control)
        {
            throw new NotImplementedException();
        }

        public BackstageGroupStyle GetStyle(Microsoft.Office.Core.IRibbonControl control)
        {
            throw new NotImplementedException();
        }

        public string GetSupertip(Microsoft.Office.Core.IRibbonControl control)
        {
            throw new NotImplementedException();
        }

        public string GetSupertip(Microsoft.Office.Core.IRibbonControl control, int itemIndex)
        {
            throw new NotImplementedException();
        }

        public string GetTarget(Microsoft.Office.Core.IRibbonControl control)
        {
            throw new NotImplementedException();
        }

        public string GetText(Microsoft.Office.Core.IRibbonControl control)
        {
            throw new NotImplementedException();
        }

        public string GetTitle(Microsoft.Office.Core.IRibbonControl control)
        {
            throw new NotImplementedException();
        }

        public bool GetVisible(Microsoft.Office.Core.IRibbonControl control)
        {
            throw new NotImplementedException();
        }

        Bitmap IRibbonCallbacks.LoadImage(string image_id)
        {
            throw new NotImplementedException();
        }

        public void OnAction(Microsoft.Office.Core.IRibbonControl control)
        {
            Debug.WriteLine(control.Id);
            MyApp.ExportNamedRanges();
        }

        public void OnAction(Microsoft.Office.Core.IRibbonControl control, string itemID, int itemIndex)
        {
            throw new NotImplementedException();
        }

        public void OnChange(Microsoft.Office.Core.IRibbonControl control, string text)
        {
            throw new NotImplementedException();
        }

        public void OnHide(object contextObject)
        {
            throw new NotImplementedException();
        }

        public void OnLoad(Microsoft.Office.Core.IRibbonUI ribbon)
        {
            AppRibbon = ribbon;
            OnLoadEvent(new LoadEventArgs { Ribbon = AppRibbon });
        }

        public void OnShow(object contextObject)
        {
            throw new NotImplementedException();
        }
        #endregion




        #region events
        internal static event EventHandler<LoadEventArgs> Load;
        protected virtual void OnLoadEvent(LoadEventArgs e)
        {
            EventHandler<LoadEventArgs> handler = Load;
            handler?.Invoke(this, e);
        }
        #endregion



        #region event arguments
        public class ButtonPressedEventArgs : EventArgs
        {
            public Microsoft.Office.Core.IRibbonControl ControlId { get; set; }
        }
        public class LoadEventArgs : EventArgs
        {
            public Microsoft.Office.Core.IRibbonUI Ribbon { get; set; }
        }
        #endregion

    }
}