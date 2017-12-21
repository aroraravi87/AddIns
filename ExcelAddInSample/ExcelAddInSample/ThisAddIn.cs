using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

namespace ExcelAddInSample
{
    public partial class ThisAddIn
    {
        private Microsoft.Office.Tools.CustomTaskPane customPane;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            //this.Application.WorkbookBeforeSave += new Excel.AppEvents_WorkbookBeforeSaveEventHandler(MyExcelAddIn_BeforeSave);

            ShowUserControl();
        }

        private void ShowUserControl()
        {
            var txtObject = new CustomAddInControl();
            customPane = this.CustomTaskPanes.Add(txtObject, "Enter Details");
            customPane.Width = txtObject.Width;
            customPane.Visible = true;
        }

        //private void MyExcelAddIn_BeforeSave(Excel.Workbook Wb, bool SaveAsUI, ref bool Cancel)
        //{
        //    Excel.Worksheet currentsheet = ((Excel.Worksheet)Application.ActiveSheet);
        //    Excel.Range SheetFirstRow = currentsheet.get_Range("A1");
        //    SheetFirstRow.EntireRow.Insert(Excel.XlInsertShiftDirection.xlShiftDown);
        //    Excel.Range NewSheetRow = currentsheet.get_Range("A1");
        //    NewSheetRow.Value2 = "Visual Studio 2013 Running 123";
        //}

       

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
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
