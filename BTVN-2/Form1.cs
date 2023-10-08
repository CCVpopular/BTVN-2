using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.Remoting.Contexts;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;

namespace BTVN_2
{
    public partial class Form1 : Form
    {
        private DataGridView backupDataGridView;
        Microsoft.Office.Interop.Excel.Application excel = null;
        Workbook workbook = null;
        Worksheet sheet = null;
        Range range = null;
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog ofd = new OpenFileDialog();
                ofd.Filter = "Excel Files|*.xls;*.xlsx";
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    textBoxfilepath.Text = ofd.FileName;
                    excel = new Microsoft.Office.Interop.Excel.Application();
                    workbook = excel.Workbooks.Open(ofd.FileName);
                    sheet = workbook.Worksheets[1];
                    range = sheet.UsedRange;
                    dataGridView1.Rows.Clear();
                    dataGridView1.Columns.Clear();
                    dataGridView1.Columns.Add("Column 1", "TT");
                    for (int col = 1; col <= range.Columns.Count; col++)
                    {
                        string columnstext = range.Cells[1, col].Value.ToString();
                        dataGridView1.Columns.Add("Column " + (col + 1), columnstext);
                    }

                    for (int row = 2; row <= range.Rows.Count; row++)
                    {
                        DataGridViewRow dataGridViewRow = new DataGridViewRow();
                        dataGridView1.Rows.Add(dataGridViewRow);
                        for (int col = 1; col <= range.Columns.Count; col++)
                        {
                            dataGridView1.Rows[row - 2].Cells[col].Value = range.Cells[row, col].Value;
                        }
                    }
                    excel.Workbooks.Close();
                    excel.Quit();
                }
            }
            catch (Exception ex)
            {
                dataGridView1.Rows.Clear();
                dataGridView1.Columns.Clear();
                textBoxfilepath.Text = "";
                excel.Workbooks.Close();
                excel.Quit();
                MessageBox.Show(ex.Message);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            backupDataGridView = new DataGridView();
            foreach (DataGridViewColumn col in dataGridView1.Columns)
            {
                backupDataGridView.Columns.Add(col.Clone() as DataGridViewColumn);
            }
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (row.IsNewRow) continue;
                DataGridViewRow newRow = (DataGridViewRow)row.Clone();
                for (int i = 0; i < row.Cells.Count; i++)
                {
                    newRow.Cells[i].Value = row.Cells[i].Value;
                }
                backupDataGridView.Rows.Add(newRow);
            }
            try
                {
                OpenFileDialog ofd = new OpenFileDialog();
                ofd.Filter = "Excel Files|*.xls;*.xlsx";
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    excel = new Microsoft.Office.Interop.Excel.Application();
                    workbook = excel.Workbooks.Open(ofd.FileName);
                    sheet = workbook.Worksheets[1];
                    range = sheet.UsedRange;
                    for(int i = 1; i < (dataGridView1.Columns.Count); i++) 
                    {
                        if (dataGridView1.Columns[i].HeaderText.ToString() != range.Cells[1,i].Value.ToString())
                        {
                            throw new Exception("File import kiểu dữ liệu với file đã có sẵn");
                        }
                    }

                    for (int row = 2; row <= range.Rows.Count; row++)
                    {
                        int existingRowIndex = FindRowIndexByValue(dataGridView1, range.Cells[row, 1].Value.ToString());
                        if (existingRowIndex != -1)
                        {
                            dataGridView1.Rows[existingRowIndex].Cells[0].Value = "Cập Nhật";
                            for (int col = 1; col <= range.Columns.Count; col++)
                            {
                                dataGridView1.Rows[existingRowIndex].Cells[col].Value = range.Cells[row, col].Value;
                            }
                        }
                        else
                        {
                            DataGridViewRow dataGridViewRow = new DataGridViewRow();
                            dataGridView1.Rows.Add(dataGridViewRow);
                            int newIndex = dataGridView1.Rows.Count - 1;
                            dataGridView1.Rows[newIndex].Cells[0].Value = "Thêm Mới";

                            for (int col = 1; col <= range.Columns.Count; col++)
                            {
                                dataGridView1.Rows[newIndex].Cells[col].Value = range.Cells[row, col].Value;
                            }
                        }
                    }
                        excel.Workbooks.Close();
                        excel.Quit();
                }
            }
            catch (Exception ex)
            {
                dataGridView1.Rows.Clear();
                dataGridView1.Columns.Clear();
                foreach (DataGridViewColumn col in backupDataGridView.Columns)
                {
                    dataGridView1.Columns.Add(col.Clone() as DataGridViewColumn);
                }
                foreach (DataGridViewRow row in backupDataGridView.Rows)
                {
                    if (row.IsNewRow) continue;
                    DataGridViewRow newRow = (DataGridViewRow)row.Clone();
                    for (int i = 0; i < row.Cells.Count; i++)
                    {
                        newRow.Cells[i].Value = row.Cells[i].Value;
                    }
                    dataGridView1.Rows.Add(newRow);
                }
                excel.Workbooks.Close();
                excel.Quit();
                MessageBox.Show(ex.Message);
            }
        }

        private int FindRowIndexByValue(DataGridView dataGridView, string value)
        {
            for (int i = 0; i < dataGridView.RowCount; i++)
            {
                if (dataGridView.Rows[i].Cells[1].Value != null && dataGridView.Rows[i].Cells[1].Value.ToString() == value)
                {
                    return i;
                }
            }
            return -1;
        }
    }
}
