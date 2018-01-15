using System;
using System.Runtime.InteropServices;
//using System.Windows.Forms;
using ExcelDna.Integration.CustomUI;
using myResources = pliExcelSheetsImageSaver.Properties;
using System.Diagnostics;


namespace pliExcelSheetsImageSaver
{
    [ComVisible(true)]
    public class RibbonController : ExcelRibbon
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
        public static IRibbonUI AppRibbon{ get; protected set; }
        #endregion

        #region callbacks
        #region OnButtonPressed
        public static event EventHandler ButtonPressed;
        // ribbon callback
        public void OnButtonPressed(IRibbonControl control)
        {
            //string ExcelWBName = MyApp.ExcelApp.ActiveWorkbook.Name;
            //Debug.WriteLine(ExcelWBName);
            //MyApp.ExportNamedRanges();
            ButtonPressedEventArgs args = new ButtonPressedEventArgs { ControlId = control };
            OnButtonPressedEvent(args);

        }
        // raise event
        protected virtual void OnButtonPressedEvent(EventArgs e)
        {
            EventHandler handler = ButtonPressed;
            handler?.Invoke(this, e);
        }
        #endregion
        #region OnLoad
        public static event EventHandler Load;
        // ribbon callback
        public void OnLoad(IRibbonUI ribbon)
        {
            Debug.WriteLine("Ribbon geladen");
            AppRibbon = ribbon;
            LoadEventArgs args = new LoadEventArgs
            {
                Ribbon = ribbon
            };
            
            OnLoadEvent(args);
        }
        // raise event
        protected virtual void OnLoadEvent(EventArgs e)
        {
            Debug.WriteLine("Event will fire");
            EventHandler handler = Load;
            handler?.Invoke(this, e);
            Debug.WriteLine("Event Fired");

        }
        #endregion

        #endregion

        #region event arguments
        public class ButtonPressedEventArgs : EventArgs
        {
            public IRibbonControl ControlId { get; set; }
        }
        public class LoadEventArgs : EventArgs
        {
            public IRibbonUI Ribbon { get; set; }
        }
        #endregion

    }
}