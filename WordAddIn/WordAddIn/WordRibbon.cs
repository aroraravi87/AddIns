using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;

namespace WordAddIn
{
    public partial class WordRibbon
    {
        private void WordRibbon_Load(object sender, RibbonUIEventArgs e)
        {


        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            SaveFileDialog objFileDialog = new SaveFileDialog();
            objFileDialog.FileName = "*";
            objFileDialog.DefaultExt = "pdf";
            objFileDialog.ValidateNames = true;
            if (objFileDialog.ShowDialog() == DialogResult.OK)
            {
                Globals.ThisAddIn.Application.ActiveDocument.ExportAsFixedFormat(objFileDialog.FileName,
                    WdExportFormat.wdExportFormatPDF, OpenAfterExport: true);
            }
        }

        private void btnAddImage_Click(object sender, RibbonControlEventArgs e)
        {
            OpenFileDialog objDialog = new OpenFileDialog();
            objDialog.FileName = "*";
            objDialog.DefaultExt = "png";
            objDialog.ValidateNames = true;
            objDialog.Filter =
                "Bitmap Image (.bmp)|*.bmp|Gif Image (.gif)|*.gif|JPEG Image (.jpeg)|*.jpeg|Png Image (.png)|*.png";
            if (objDialog.ShowDialog() == DialogResult.OK)
            {
                Globals.ThisAddIn.Application.ActiveDocument.Shapes.AddPicture(objDialog.FileName);

            }
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Application.ActiveDocument.Tables.Add(
                Globals.ThisAddIn.Application.ActiveDocument.Range(0, 0), 3, 4);
            //Globals.ThisAddIn.Application.ActiveDocument.Tables[1].Range.Shading.BackgroundPatternColor=WdColor.wdColorBlueGray;
            Globals.ThisAddIn.Application.ActiveDocument.Tables[1].Range.Font.Size = 12;
            Globals.ThisAddIn.Application.ActiveDocument.Tables[1].Range.Rows.Borders.Enable = 1;
  
        }

    }
}