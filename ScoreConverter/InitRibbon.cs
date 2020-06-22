using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace ScoreConverter
{
    public partial class InitRibbon
    {
        private void InitRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void OpenFormButton_Click(object sender, RibbonControlEventArgs e)
        {
            var form = new SelectFileForm();
            form.Show();
        }
    }
}
