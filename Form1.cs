using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using ClosedXML.Excel;

namespace ExcelUtls
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            this.dataGridView1.Dock = DockStyle.Fill;

            // import dialog 
            importDialog.Filter = "Excel File(*.xlsx)|*.xlsx";
            importDialog.CheckFileExists = true;
            importDialog.CheckPathExists = true;
            importDialog.ValidateNames = true;
            importDialog.Multiselect = true;

            // export dialog
            exportDialog.Filter = "Excel File(*.xlsx)|*.xlsx";

            //
            this.statusStrip1.Items[0].Text = "Welcome :)";
        }

        private string GetValueString(object obj)
        {
            return obj == null ? "" : obj.ToString();
        }

        private double ParseToDouble(object obj)
        {
            try
            {
                return double.Parse(obj.ToString());
            }
            catch
            {
                return 0;
            }
        }

        private void textBoxColumn_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == System.Windows.Forms.MouseButtons.Left)
            {
                TextBox tb = sender as TextBox;
                if (tb != null)
                {
                    tb.SelectAll();
                }
            }
        }

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                string file = dataGridView1.Rows[e.RowIndex].Tag as string;
                try
                {
                    System.Diagnostics.Process.Start(file);
                }
                catch
                {
                    MessageBox.Show(file + "无法打开！");
                }
            }
        }

        private void Show(int sheetNo, string cellName)
        {
            dataGridView1.Rows.Clear();
            dataGridView1.Columns["sheet"].HeaderText = "Sheet (" + sheetNo.ToString() + ")";
            dataGridView1.Columns["cell"].HeaderText = "Cell (" + cellName + ")";
            foreach (KeyValuePair<string, XLWorkbook> wb in wbs)
            {
                var ws = wb.Value.Worksheet(sheetNo);
                var cell = ws.Cell(cellName);

                DataGridViewRow row = new DataGridViewRow();
                row.Tag = wb.Key;
                row.CreateCells(dataGridView1, new object[] { Path.GetFileName(wb.Key), ws.Name, cell.Value });
                dataGridView1.Rows.Add(row);
            }
        }

        private void Calc()
        {
            string cellName = this.textBoxColumn.Text.Trim() + this.textBoxRow.Text.Trim();
            Show(1, cellName);

            List<double> vals = new List<double>();
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                vals.Add(ParseToDouble(row.Cells["cell"].Value));
            }

            this.textBoxSum.Text = vals.Count == 0 ? "" : vals.Sum().ToString();
            this.textBoxAvg.Text = vals.Count == 0 ? "" : vals.Average().ToString();
            this.textBoxMax.Text = vals.Count == 0 ? "" : vals.Max().ToString();
            this.textBoxMin.Text = vals.Count == 0 ? "" : vals.Min().ToString();
        }

        private void ParseExcels()
        {
            foreach (string f in files)
            {
                wbs[f] = new XLWorkbook(f);
            }
            Calc();
        }
        
        private void importToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (importDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                files = new List<string>(importDialog.FileNames);
                ParseExcels();
            }
        }

        private void exportToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (exportDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                var wb = new XLWorkbook();
                var ws = wb.AddWorksheet("Sheet1");

                for (int i = 0; i < dataGridView1.Columns.Count; ++i)
                {
                    ws.Cell(1, i + 1).Value = dataGridView1.Columns[i].HeaderText;
                }

                for (int i = 0; i < dataGridView1.Rows.Count; ++i)
                {
                    for (int j = 0; j < dataGridView1.Columns.Count; ++j)
                    {
                        ws.Cell(i + 2, j + 1).Value = GetValueString(dataGridView1.Rows[i].Cells[j].Value);
                    }
                }

                string file = exportDialog.FileName;
                wb.SaveAs(file);
            }

        }

        private void RefreshExcels()
        {
            foreach (string f in files)
            {
                if (!File.Exists(f))
                {
                    files.Remove(f);
                    wbs.Remove(f);
                }
            }
            ParseExcels();
        }

        private void refreshToolStripMenuItem_Click(object sender, EventArgs e)
        {
            RefreshExcels();
        }

        private OpenFileDialog importDialog = new OpenFileDialog();
        private SaveFileDialog exportDialog = new SaveFileDialog();

        private List<string> files = new List<string>();
        private Dictionary<string, XLWorkbook> wbs = new Dictionary<string, XLWorkbook>();

        private void textBoxColumn_TextChanged(object sender, EventArgs e)
        {
            Calc(); 
        }

        private void textBoxRow_TextChanged(object sender, EventArgs e)
        {
            Calc();
        }

        private void textBoxColumn_Validating(object sender, CancelEventArgs e)
        {
        }

        private void textBoxColumn_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsLetter(e.KeyChar))
            {
                e.Handled = true;
            }
            else
            {
                e.KeyChar = char.ToUpper(e.KeyChar);
            }
        }

        private void textBoxRow_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsDigit(e.KeyChar))
            {
                e.Handled = true;
            }
        }

        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("           Cell One 1.0.1              \n \n               For Lisa ", "About", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void toolStripStatusLabel2_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("http://www.weibo.com/hutusi/");
        }

        private void addToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (importDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                files.AddRange(importDialog.FileNames);
                RefreshExcels();
            }
        }

        private void removeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow row in dataGridView1.SelectedRows)
            {
                string file = row.Tag as string;
                files.Remove(file);
                wbs.Remove(file);
            }
            RefreshExcels();
        }
    }
}
