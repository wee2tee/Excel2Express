using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel2Express.SubForm;
using CC;

namespace Excel2Express
{
    public enum FORM_MODE
    {
        READ,
        EDIT,
        PROCESS
    }

    public partial class Form1 : Form
    {
        public string selected_excel_file_name = null;
        public DataTable data;
        private FORM_MODE form_mode;
        private DataGridViewColumn editing_column = null;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.ResetFormState(FORM_MODE.READ);
            this.xDatagrid1.AllowUserToResizeColumns = true;
            this.HideInlineForm();
        }

        private void ResetFormState(FORM_MODE form_mode)
        {
            this.form_mode = form_mode;

            this.xDatagrid1.SetControlState(new FORM_MODE[] { FORM_MODE.READ }, this.form_mode);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog fd = new OpenFileDialog();
            if(fd.ShowDialog() == DialogResult.OK)
            {
                this.textBox1.Text = fd.FileName;
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            this.selected_excel_file_name = ((TextBox)sender).Text;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            using (ExcelPackage xls = new ExcelPackage(new FileInfo(this.selected_excel_file_name)))
            {
                this.data = this.GetDataTable(xls.Workbook.Worksheets.First());
                this.xDatagrid1.DataSource = this.data;
            }
        }

        private DataTable GetDataTable(ExcelWorksheet wrk_sheet)
        {
            int total_row = wrk_sheet.Dimension.End.Row;
            int total_col = wrk_sheet.Dimension.End.Column;
            
            List<List<string>> list_data = new List<List<string>>();
            
            for (int i = 1; i <= total_row; i++)
            {
                List<string> row = new List<string>();

                wrk_sheet.Cells[i, 1, i, total_col].ToList().ForEach(c =>
                {
                    string val = c.Value != null ? c.Value.ToString() : string.Empty;
                    row.Add(val);
                });

                list_data.Add(row);
            }

            DataTable dt = new DataTable();

            // get max column
            int max_col = 0;
            foreach (var arr in list_data)
            {
                if (arr.Count > max_col)
                {
                    max_col = arr.Count;
                }
            }
            // add column to datatable
            for (int i = 0; i < max_col; i++)
            {
                dt.Columns.Add();
            }

            // fill data from List<string> to DataRow
            foreach (var arr in list_data)
            {
                dt.Rows.Add(arr.ToArray<string>());
            }

            return dt;
        }

        private void xDatagrid1_MouseClick(object sender, MouseEventArgs e)
        {
            if(e.Button == MouseButtons.Right)
            {
                int row_index = ((XDatagrid)sender).HitTest(e.X, e.Y).RowIndex;
                int col_index = ((XDatagrid)sender).HitTest(e.X, e.Y).ColumnIndex;
                if (row_index == -1) // specify field name
                {
                    DataGridViewColumn col = ((XDatagrid)sender).Columns.Cast<DataGridViewColumn>().Where(c => c.Index == col_index).FirstOrDefault();

                    ContextMenu cm = new ContextMenu();
                    MenuItem mnu_edit = new MenuItem("แก้ไขชื่อฟิลด์");
                    mnu_edit.Click += delegate
                    {
                        this.ShowInlineForm(col);
                    };
                    cm.MenuItems.Add(mnu_edit);

                    MenuItem mnu_prop = new MenuItem("แก้ไขคุณสมบัติของฟิลด์");
                    mnu_prop.Click += delegate
                    {
                        FormPropertyEditor pe = new FormPropertyEditor(col);
                        pe.ShowDialog();
                    };
                    cm.MenuItems.Add(mnu_prop);

                    cm.Show(((XDatagrid)sender), new Point(e.X, e.Y));
                }
                else
                {
                    ((XDatagrid)sender).Rows[row_index].Cells[((XDatagrid)sender).FirstDisplayedScrollingColumnIndex].Selected = true;

                    ContextMenu cm = new ContextMenu();
                    MenuItem mnu_del = new MenuItem("ลบ");
                    mnu_del.Click += delegate
                    {
                        ((XDatagrid)sender).Rows[row_index].DrawDeletingRowOverlay();
                        if(MessageBox.Show("ลบรายการที่เลือกนี้ ทำต่อหรือไม่?", "", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
                        {
                            this.data.Rows.RemoveAt(row_index);
                        }
                        else
                        {
                            ((XDatagrid)sender).Rows[row_index].ClearDeletingRowOverlay();
                        }

                    };
                    cm.MenuItems.Add(mnu_del);

                    cm.Show(((XDatagrid)sender), new Point(e.X, e.Y));
                }
            }
        }

        private void xDatagrid1_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            this.xDatagrid1.Columns.Cast<DataGridViewColumn>().ToList().ForEach(c => c.SortMode = DataGridViewColumnSortMode.NotSortable);

        }

        private void ShowInlineForm(DataGridViewColumn col)
        {
            Rectangle rect = col.DataGridView.GetCellDisplayRectangle(col.Index, -1, true);
            this.inlineFieldName.SetBounds(rect.X, rect.Y, rect.Width, rect.Height);
            this.inlineFieldName.Text = col.HeaderText;
            this.inlineFieldName.Focus();
            this.editing_column = col;
        }

        private void HideInlineForm()
        {
            this.inlineFieldName.SetBounds(-9999, -9999, this.inlineFieldName.Width, this.inlineFieldName.Height);
        }

        private void inlineFieldName_KeyPress(object sender, KeyPressEventArgs e)
        {
            if(e.KeyChar == (char)13)
            {
                if(((TextBox)sender).Text.Trim().Length == 0)
                {
                    MessageBox.Show("กรุณาระบุชื่อฟิลด์", "", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    return;
                }

                if(this.xDatagrid1.Columns.Cast<DataGridViewColumn>().Where(c => c.HeaderText.Trim() == ((TextBox)sender).Text.Trim()).Count() > 0)
                {
                    MessageBox.Show("ชื่อฟิลด์ " + ((TextBox)sender).Text + " นี้มีอยู่แล้ว, กรุณาเปลี่ยนใหม่", "", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    return;
                }

                this.editing_column.HeaderText = ((TextBox)sender).Text;
                this.HideInlineForm();
                this.ResetFormState(FORM_MODE.READ);
                this.editing_column = null;
                return;
            }

            if(e.KeyChar == (char)27)
            {
                this.HideInlineForm();
                this.ResetFormState(FORM_MODE.READ);
                this.editing_column = null;
                return;
            }
        }
    }

    public static class HelperClass
    {
        public static void SetControlState(this Component comp, FORM_MODE[] allow_active_mode, FORM_MODE current_mode)
        {
            if(comp is DataGridView)
            {
                ((DataGridView)comp).Enabled = allow_active_mode.Where(m => m.ToString() == current_mode.ToString()).Count() > 0 ? true : false;
            }
        }
    }
}
