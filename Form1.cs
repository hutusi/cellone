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

        private int ParseToInt(object obj)
        {
            try
            {
                return int.Parse(obj.ToString());
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
                string dir = Directory.GetCurrentDirectory();
                string file = dir + "\\" + dataGridView1.Rows[e.RowIndex].Cells["Excel"].Value.ToString();
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
                DataGridViewRow row = new DataGridViewRow();
                var ws = wb.Value.Worksheet(sheetNo);
                var cell = ws.Cell(cellName);

                dataGridView1.Rows.Add(new object[] { Path.GetFileName(wb.Key), ws.Name, cell.Value });
            }
        }

        private void Calc()
        {
            string cellName = this.textBoxColumn.Text.Trim() + this.textBoxRow.Text.Trim();
            Show(1, cellName);

            List<int> vals = new List<int>();
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                vals.Add(ParseToInt(row.Cells["cell"].Value));
            }

            this.textBoxSum.Text = vals.Sum().ToString();
            this.textBoxAvg.Text = vals.Average().ToString();
            this.textBoxMax.Text = vals.Max().ToString();
            this.textBoxMin.Text = vals.Min().ToString();
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

        private void refreshToolStripMenuItem_Click(object sender, EventArgs e)
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
            MessageBox.Show("           Cell One 1.0.0              \n \n               For Lisa ", "About", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void toolStripStatusLabel2_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("http://www.weibo.com/hutusi/");
        }
    }
}
