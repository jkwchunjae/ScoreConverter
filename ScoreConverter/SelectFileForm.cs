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
                SourceWorksheet.Items.Clear();
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
            try
            {
                if (ExcelApp.TryGetWorkbook(x => x.Name == sourceWorkbookName, out var sourceWorkbook))
                {
                    if (sourceWorkbook.TryGetWorksheet(x => x.Name == SourceWorksheet.Text, out var sourceWorksheet))
                    {
                        if (ExcelApp.TryGetWorkbook(x => TargetWorkbook.Text == x.Name, out var targetWorkbook))
                        {
                            var result = Converter.Validate(sourceWorksheet, targetWorkbook);
                            if (result)
                            {
                                MessageBox.Show("검사 통과하였습니다.");
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + Environment.NewLine + ex.StackTrace);
            }
        }

        private void ExecuteButton_Click(object sender, EventArgs e)
        {
            var sourceWorkbookName = SourceWorkbook.Text;
            if (ExcelApp.TryGetWorkbook(x => x.Name == sourceWorkbookName, out var sourceWorkbook))
            {
                if (sourceWorkbook.TryGetWorksheet(x => x.Name == SourceWorksheet.Text, out var sourceWorksheet))
                {
                    if (ExcelApp.TryGetWorkbook(x => TargetWorkbook.Text == x.Name, out var targetWorkbook))
                    {
                        Converter.Execute(sourceWorksheet, targetWorkbook);
                        MessageBox.Show("완료");
                    }
                }
            }
        }
    }
}
