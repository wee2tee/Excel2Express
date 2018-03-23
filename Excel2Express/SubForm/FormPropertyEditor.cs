using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Excel2Express.SubForm
{
    public partial class FormPropertyEditor : Form
    {
        public DataGridViewColumn editing_column = null;

        public FormPropertyEditor(DataGridViewColumn col)
        {
            InitializeComponent();
            this.editing_column = col;
            Enum.GetValues(typeof(FIELD_TYPE)).Cast<FIELD_TYPE>().ToList().ForEach(f =>
            {
                this.txtFieldType.Items.Add(new XComboboxItem { text = f.ToString(), type = f });
            });
        }

        private void FormPropertyEditor_Load(object sender, EventArgs e)
        {
            this.txtFieldName.Text = this.editing_column.HeaderText;

            Console.WriteLine(" ==> column type : " + this.editing_column.ValueType.ToString());
        }

        private void btnOK_Click(object sender, EventArgs e)
        {

        }

        private void txtFieldType_SelectedIndexChanged(object sender, EventArgs e)
        {
            XComboboxItem selected_item = (XComboboxItem)((ComboBox)sender).SelectedItem;
            if (selected_item.type == FIELD_TYPE.DOUBLE)
            {
                this.editing_column.ValueType = typeof(Double);
            }
        }
    }

    public enum FIELD_TYPE
    {
        CHARACTER,
        DOUBLE,
        NUMERIC,
        DATE,
    }

    public class XComboboxItem
    {
        public string text { get; set; }
        public FIELD_TYPE type { get; set; }
        public override string ToString()
        {
            return this.text;
        }
    }

}
