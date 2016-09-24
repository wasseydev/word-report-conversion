using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

using Converter;
using System.Windows.Forms;

namespace ReportConversionAddIn
{
    public partial class MainRibbon
    {
        private void MainRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void CreateJRXMLReport()
        {
            Jasper jasper;
            try
            {
                jasper = new Jasper(Globals.ThisAddIn.Application);
                jasper.ProcessDocument();
#if DEBUG
                // Output the generated JRXML to the current document if debugging
                var sel = Globals.ThisAddIn.Application.Selection;
                sel.WholeStory();
                sel.Delete();
                sel.TypeText(jasper.JRXML);
                // Mark as saved so we don't get asked all the time...
                Globals.ThisAddIn.Application.ActiveDocument.Saved = true;
#endif
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error encountered: " + ex.Message);
            }
        }

        private void CreateJRXML_Click(object sender, RibbonControlEventArgs e)
        {
            CreateJRXMLReport();
        }
    }
}
