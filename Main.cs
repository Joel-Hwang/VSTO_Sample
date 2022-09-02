using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace VSTO_Sample
{
    public partial class Main : UserControl
    {
        public Main()
        {
            InitializeComponent();
        }

        private void btnDataLoad_Click(object sender, EventArgs e)
        {
            Worksheet ws = Globals.ThisAddIn.Application.ActiveSheet;
            ws.Range["A1"].Value2 = "Test Data1";
            ws.Range["A2"].Value2 = "Test Data2";
            ws.Range["A3"].Value2 = "Test Data3";
            ws.Range["A4"].Value2 = "Test Data4";

            ws.Range["B1"].Value2 = "1";
            ws.Range["B2"].Value2 = "2";
            ws.Range["B3"].Value2 = "3";
            ws.Range["B4"].Value2 = "B4";
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            Worksheet ws = Globals.ThisAddIn.Application.ActiveSheet;

            for (int i = 1; i<=100; i++)
            {
                string val = GetCellValue(ws,"B"+i);

                if (string.IsNullOrEmpty(val)) continue;

                int res = 0;
                if (int.TryParse(val,out res) == false)
                {
                    MessageBox.Show(val + " is not a number.");
                    return;
                }
            }

            MessageBox.Show(string.Format("Done ( Call the upload API Here )\n\n {0} {1} {2} {3}", 
                GetCellValue(ws, "A1"), GetCellValue(ws, "A2"), GetCellValue(ws, "A3"), GetCellValue(ws, "A4")));


        }


        private string GetCellValue(Worksheet ws, string addr)
        {
            string val = ws.Range[addr].Value2 == null ? "" : ws.Range[addr].Value2.ToString();
            return val;
        }
    }
}
