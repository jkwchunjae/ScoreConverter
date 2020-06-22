using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ScoreConverter
{
    public partial class SelectFileForm : Form
    {
        Excel.Application ExcelApp => Globals.ThisAddIn.Application;

        public SelectFileForm()
        {
            InitializeComponent();
        }

        private void SelectFileForm_Load(object sender, EventArgs e)
        {
            SourceWorkbook.Items.AddRange(ExcelApp.GetWorkbooks().Select(x => x.Name).ToArray());
            TargetWorkbook.Items.AddRange(ExcelApp.GetWorkbooks().Select(x => x.Name).ToArray());
        }

        private void SourceWorkbook_SelectedIndexChanged(object sender, EventArgs e)
        {
            var sourceWorkbookName = SourceWorkbook.Text;
            if (ExcelApp.TryGetWorkbook(x => x.Name == sourceWorkbookName, out var workbook))
            {
                SourceWorksheet.Items.AddRange(workbook.GetWorksheets().Select(x => x.Name).ToArray());
            }
            else
            {
                SourceWorksheet.Items.Clear();
            }
        }

        private void ValidateButton_Click(object sender, EventArgs e)
        {
            var sourceWorkbookName = SourceWorkbook.Text;
            if (ExcelApp.TryGetWorkbook(x => x.Name == sourceWorkbookName, out var sourceWorkbook))
            {
                if (sourceWorkbook.TryGetWorksheet(x => x.Name == SourceWorksheet.Text, out var sourceWorksheet))
                {
                    Validator.Validate(sourceWorksheet, null);
                }
            }
        }
    }
}
