using System;
using System.Collections.Generic;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Windows.Forms;

namespace InspectorWrapperExplained
{
    public partial class ThisAddIn
    {

        /// <summary>
        /// Holds a reference to the Application.Inspectors collection
        /// Required to get notifications for NewInspector events.
        /// </summary>
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
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            try
            {
                _wrappedInspectors = new Dictionary<Guid, InspectorWrapper>();
                _inspectors = Globals.ThisAddIn.Application.Inspectors;
                _inspectors.NewInspector += new Outlook.InspectorsEvents_NewInspectorEventHandler(WrapInspector);
                //MessageBox.Show("HERE 3");

                // Handle also already existing Inspectors
                // (e.g. Double clicking a .msg file)
                foreach (Outlook.Inspector inspector in _inspectors)
                {
                    WrapInspector(inspector);
                }
            }
            catch (Exception er) { MessageBox.Show(er.ToString()); }
            //MessageBox.Show("HERE 4");
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


