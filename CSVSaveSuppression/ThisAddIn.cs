using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using XlFileFormat = Microsoft.Office.Interop.Excel.XlFileFormat;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

namespace CSVSaveSuppression
{
    public partial class ThisAddIn
    {
        /// <summary>
        /// Adds a custom event handler to the Application calling the custom Application_WorkbookBeforeSave method when a file is saved
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.Application.WorkbookBeforeSave += Application_WorkbookBeforeSave;
        }

        /// <summary>
        /// Must exist to clean up the AddIn, but this AddIn doesn't need any special cleanup and the method is left empty.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        /// <summary>
        /// Handels saving CSV files differently by programatically saving under special conditions and canceling the origial save.
        /// 
        /// Does nothing is a SaveAs is being performed.  Otherwise, checks the file types for any of Excels CSV file formats, and
        /// if found performs a programatic save, marks the file as saved, and cancels the original save so as to not save twice.
        /// </summary>
        /// <param name="Wb"></param>
        /// <param name="SaveAsUi"></param>
        /// <param name="Cancel"></param>
        private void Application_WorkbookBeforeSave(Excel.Workbook Wb, bool SaveAsUi, ref bool Cancel)
        {
            // Check if a Save As is being performed
            if (!SaveAsUi)
            {
                if ((new[] { XlFileFormat.xlCSV, XlFileFormat.xlCSVMac, XlFileFormat.xlCSVMSDOS, XlFileFormat.xlCSVWindows }).Contains(Wb.FileFormat))
                {
                    // Temporarily un-register this event handler to prevent recursion
                    this.Application.WorkbookBeforeSave -= Application_WorkbookBeforeSave;
                    Wb.Save();
                    // Re-register this event handler
                    this.Application.WorkbookBeforeSave += Application_WorkbookBeforeSave;

                    // Mark the file as being saved
                    Wb.Saved = true;

                    // Cancel the orginal save
                    Cancel = true;
                }
            }
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
