using ExcelAddin.Commands;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelAddin
{
    public partial class RibbonPanel
    {
        private void RibbonPanel_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnMultipleCSV_Click(object sender, RibbonControlEventArgs e)
        {
            ExcelUtilities.ImportMultipleCsv();
        }

        private void btnMultipleCsvSettings_Click(object sender, RibbonControlEventArgs e)
        {
            ExcelUtilities.ChangeDelimiter();
        }
    }
}
