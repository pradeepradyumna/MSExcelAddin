using System;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;

namespace CustomAddin
{
    public partial class MyRibbon
    {
        private void MyRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btn_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Worksheet currentWorkSheet = Globals.ThisAddIn.GetCurrentWorkSheet();

            currentWorkSheet.Range["E8"].Value = $"Hey { Environment.UserName}! What's up!";
            currentWorkSheet.Columns.AutoFit();
        }
    }
}
