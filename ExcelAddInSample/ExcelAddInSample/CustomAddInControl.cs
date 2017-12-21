using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
namespace ExcelAddInSample
{
    public partial class CustomAddInControl : UserControl
    {
        private static int counter = 1;
        public CustomAddInControl()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
          //  Range SheetFirstRow = (Globals.ThisAddIn.Application.ActiveCell);
            Worksheet currentsheet = ((Worksheet)Globals.ThisAddIn.Application.ActiveSheet);
            Range SheetFirstRow = currentsheet.get_Range("A" + counter + "");
            if (string.IsNullOrWhiteSpace(SheetFirstRow.Value2))
            {
               
                SheetFirstRow.Interior.Color = Color.Yellow; //Active Cell back Color
                SheetFirstRow.Borders.Color = Color.Black; // Active Cell border Color
                SheetFirstRow.Borders.LineStyle = XlLineStyle.xlContinuous;
                SheetFirstRow.Value = string.Format("{0} reside in {1} and current status is {2}", txtName.Text,
                    txtaddress.Text, chkactive.Checked); //Active Cell Text Add
                SheetFirstRow.Columns.AutoFit();
                counter++;
            }
           
        }


    }
}
