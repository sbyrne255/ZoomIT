using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using System.Drawing;
using InspectorWrapperExplained.Properties;

// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new Ribbon1();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace InspectorWrapperExplained
{
    [ComVisible(true)]
    public class Ribbon1 : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public Ribbon1()
        {
        }

        private string x = "";
        private bool zoomPanes = Properties.Settings.Default.zoomPanes;

        public Bitmap imageSuper_GetImage(Office.IRibbonControl control)
        {
            try
            {
                return Resources.Zoom_in_512;
            }
            catch (Exception er) { MessageBox.Show(er.ToString()); return null; }
        }
        #region IRibbonExtensibility Members
        public void buttonAction(Office.IRibbonControl control, bool isPressed)
        {
            Properties.Settings.Default.zoomPanes = isPressed;
            zoomPanes = isPressed;
        }
        public bool check_changed(Office.IRibbonControl control)
        {
            return zoomPanes;
        }
        public void Button_Click(Office.IRibbonControl control)
        {
            try
            {
                if (x.Length <= 0)
                {
                    Properties.Settings.Default.zoomLevel = 150;
                    Properties.Settings.Default.Save();
                }
                else
                {
                    int q = Convert.ToInt32(x);

                    if (q > 500)
                    {
                        Properties.Settings.Default.zoomLevel = 500;
                        Properties.Settings.Default.Save();
                        return;
                    }
                    if (q < 10)
                    {
                        Properties.Settings.Default.zoomLevel = 10;
                        Properties.Settings.Default.Save();
                        return;
                    }
                    else
                    {
                        Properties.Settings.Default.zoomLevel = q;
                        Properties.Settings.Default.Save();
                    }
                }
            }
            catch (Exception) { MessageBox.Show("Please only enter numbers. \n"); }
        }
        public void RecupDonnee(Office.IRibbonControl control, String text)
        {
            try
            {
                x = text;
            }
            catch (Exception er) { MessageBox.Show(er.ToString()); }

        }
        #endregion

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("InspectorWrapperExplained.Ribbon1.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, select the Ribbon XML item in Solution Explorer and then press F1

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
