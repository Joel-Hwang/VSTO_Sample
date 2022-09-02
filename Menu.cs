using Microsoft.Office.Tools;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace VSTO_Sample
{
    public partial class Menu
    {
        private void Menu_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            Main main = new Main();

            CustomTaskPane taskPane = Globals.ThisAddIn.CustomTaskPanes.Add(main, "Hello World");
            taskPane.Width = 500;
            taskPane.Visible = true;
        }
    }
}
