﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

using Converter;
using System.Windows.Forms;

namespace WordReportingTool
{
    public partial class WordReportingRibbon
    {
        private void WordReportingRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void CreateJRXMLReport()
        {
            Jasper jasper;
            try
            {
                jasper = new Jasper(Globals.ThisAddIn.Application);
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
