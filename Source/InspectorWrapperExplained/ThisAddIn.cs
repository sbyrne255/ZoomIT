﻿using System;
using System.Collections.Generic;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Windows.Forms;
using Microsoft.Office.Core;
using System.Runtime.InteropServices;
using System.Threading;
using System.Drawing;


namespace InspectorWrapperExplained
{
    public partial class ThisAddIn
    {
        private string lastItem = "";
        private bool isDone = true;
        private Outlook.Explorer currentExplorer = null;

        [DllImport("user32.dll", SetLastError = true)]
        static extern void keybd_event(byte bVk, byte bScan, int dwFlags, int dwExtraInfo);
        public const int KEYEVENTF_EXTENDEDKEY = 0x0001; //Key down flag
        public const int KEYEVENTF_KEYUP = 0x0002; //Key up flag
        public const int VK_CONTROL = 0x11; //Control key code
        public const int VK_TAB = 0x09; //tab key code
        public const int VK_SHIFT = 0x10; //shift key code
        private Outlook.Inspectors _inspectors;
        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new Ribbon1();
        }
        /// <summary>
        /// A dictionary that holds a reference to the Inspectors handled by the add-in
        /// </summary>
        private Dictionary<Guid, InspectorWrapper> _wrappedInspectors;
        /// <summary>
        /// Startup method is called when the add-in is loaded by Outlook
        /// </summary>
        /// 

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            try
            {
                _wrappedInspectors = new Dictionary<Guid, InspectorWrapper>();
                _inspectors = Globals.ThisAddIn.Application.Inspectors;
                _inspectors.NewInspector += new Outlook.InspectorsEvents_NewInspectorEventHandler(WrapInspector);

                currentExplorer = this.Application.ActiveExplorer();
                currentExplorer.SelectionChange += new Outlook.ExplorerEvents_10_SelectionChangeEventHandler(CurrentExplorer_Event);
                //currentExplorer.ViewSwitch += new Outlook.ExplorerEvents_10_ViewSwitchEventHandler(CurrentExplorer_Event);

                // Handle also already existing Inspectors
                // (e.g. Double clicking a .msg file)
                foreach (Outlook.Inspector inspector in _inspectors)
                {
                    WrapInspector(inspector);
                }
            }
            catch (Exception er) { MessageBox.Show(er.ToString()); }
        }
        private void CurrentExplorer_Event()
        {
            try
            {
                if (Properties.Settings.Default.zoomPanes)
                {
                    if (this.Application.ActiveExplorer().Selection.Count > 0 && this.Application.ActiveExplorer().Selection.Count < 2)
                    {
                        Object selObject = this.Application.ActiveExplorer().Selection[1];

                        if (selObject is Outlook.MailItem)
                        {
                            Outlook.MailItem mailItem = (selObject as Outlook.MailItem);
                            if (lastItem != mailItem.EntryID)
                            {
                                lastItem = mailItem.EntryID;
                                currentExplorer = this.Application.ActiveExplorer();
                                int h = currentExplorer.Height;
                                int w = currentExplorer.Width;


                                PointConverter pc = new PointConverter();
                                Point pt = new Point();
                                if (isDone)
                                {
                                    isDone = false;
                                    int sleep = 0;
                                    if (Keyboard.IsKeyDown(Keys.Up) || Keyboard.IsKeyDown(Keys.Down))
                                    {
                                        sleep = 800;
                                    }
                                    new Thread(() =>{
                                        //https://stackoverflow.com/questions/32405387/exiting-c-sharp-function-execution-if-one-of-the-variable-value-during-execution Look at some threading options
                                        //Problem when scrolling fast, emails don't zoom.
                                        Thread.Sleep(sleep);//Delay so the form can load/set selected item and be active before the scroll attempt...
                                        //Thread.CurrentThread.IsBackground = true;
                                        pt = (Point)pc.ConvertFromString(w.ToString() + "," + h.ToString());
                                        int posX = Cursor.Position.X;
                                        int posY = Cursor.Position.Y;
                                        Cursor.Position = pt;
                                        keybd_event(VK_TAB, 0x9d, 0, 0); // tab Press
                                        keybd_event(VK_CONTROL, 0x9d, 0, 0); // Ctrl Press
                                        //Set proper scroll...
                                        InspectorWrapperExplained.NativeMethods.MouseInput.ScrollWheel((Properties.Settings.Default.zoomLevel - 100) / 10);//num * 10% IE, 5 = +150% zoom
                                        keybd_event(VK_TAB, 0x9d, KEYEVENTF_KEYUP, 0); // Tab Release
                                        keybd_event(VK_CONTROL, 0x9d, KEYEVENTF_KEYUP, 0); // Ctrl Release

                                        pt = (Point)pc.ConvertFromString(posX.ToString() + "," + posY.ToString());
                                        Cursor.Position = pt;

                                        //Return tab to previous location (work around, doesn't work for tabs but does center messages so that uses can scroll through their emails with the arrow keys.).
                                        keybd_event(VK_SHIFT, 0x9d, 0, 0); // Shift Press
                                        keybd_event(VK_TAB, 0x9d, 0, 0); // tab Press

                                        keybd_event(VK_TAB, 0x9d, KEYEVENTF_KEYUP, 0); // Tab Release
                                        keybd_event(VK_SHIFT, 0x9d, KEYEVENTF_KEYUP, 0); // Shift Release
                                        isDone = true;

                                    }).Start();
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception)
            {
                MessageBox.Show("There was an error trying to zoom the preview, if this persists please email our support team.");
            }
        }
        /// <summary>
        /// Wraps an Inspector if required and remember it in memory to get events of the wrapped Inspector
        /// </summary>
        /// <param name="inspector">The Outlook Inspector instance</param>
        void WrapInspector(Outlook.Inspector inspector) {
            try
            {
                InspectorWrapper wrapper = InspectorWrapper.GetWrapperFor(inspector);
                if (wrapper != null)
                {
                    // register for the closed event
                    wrapper.Closed += new InspectorWrapperClosedEventHandler(wrapper_Closed);
                    // remember the inspector in memory
                    _wrappedInspectors[wrapper.Id] = wrapper;
                }
            }
            catch (Exception er) { MessageBox.Show(er.ToString()); }
        }
        /// <summary>
        /// Method is called when an inspector has been closed
        /// Removes reference from memory
        /// </summary>
        /// <param name="id">The unique id of the closed inspector</param>
        void wrapper_Closed(Guid id) {
            try
            {
                _wrappedInspectors.Remove(id);
            }
            catch (Exception er) { MessageBox.Show(er.ToString()); }
        }
        /// <summary>
        /// Shutdown method is called when Outlook is unloading the add-in
        /// </summary>
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            try
            {
                //MessageBox.Show("HERE 5");
                // do the homework and cleanup
                _wrappedInspectors.Clear();
                _inspectors.NewInspector -= new Outlook.InspectorsEvents_NewInspectorEventHandler(WrapInspector);
                _inspectors = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            catch (Exception er) { MessageBox.Show(er.ToString()); }

        }
        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}


