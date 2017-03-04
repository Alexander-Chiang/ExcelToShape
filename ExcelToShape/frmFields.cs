using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ExcelToShape
{
    public partial class frmFields : Form
    {
        DataTable dt;
        bool isChange = false;
        public frmFields(DataTable dt)
        {
            InitializeComponent();
            this.dt = dt;
        }

        private void frmFields_Load(object sender, EventArgs e)
        {
            dgFields.Columns["fields"].DataPropertyName = "fields";
            dgFields.Columns["types"].DataPropertyName = "types";
            dgFields.Columns["length"].DataPropertyName = "length";

            dgFields.DataSource = dt;

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                if (dgFields.Rows[i].Cells["length"].Value == DBNull.Value)
                    dgFields.Rows[i].Cells["length"].Style.BackColor = Color.Gray;
            }
        }

        int rowindex=-1;
        DataGridViewEditingControlShowingEventArgs cell;
        private void dgFields_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            cell = e;
            ComboBox combo = e.Control as ComboBox;
            if (combo != null)
            {
                // Remove an existing event-handler, if present, to avoid 
                // adding multiple handlers when the editing control is reused.
                combo.SelectedIndexChanged -= new EventHandler(ComboBox_SelectedIndexChanged);

                // Add the event handler. 
                combo.SelectedIndexChanged += new EventHandler(ComboBox_SelectedIndexChanged);
            }
            if (e.Control.GetType().BaseType.Name == "TextBox"&&dgFields.CurrentCell.ColumnIndex==2)
            {
                TextBox txbox = e.Control as TextBox;
                txbox.KeyPress+=new KeyPressEventHandler(txbox_KeyPress);
            }
            rowindex = dgFields.CurrentCell.RowIndex;

        }

        void txbox_KeyPress(object sender, KeyPressEventArgs e)
        {
            //限制只能输入-9的数字，退格键，小数点和回车
            if (((int)e.KeyChar >= 48 && (int)e.KeyChar <= 57) || e.KeyChar == 13 || e.KeyChar == 8 || e.KeyChar == 46)
            {
                e.Handled = false;
            }
            else
            {
                e.Handled = true;
                //MessageBox.Show("只能输入数字！");
            }
        }

       

        private void ComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComboBox combox = sender as ComboBox;
            combox.Leave += new EventHandler(combox_Leave);
            if (((ComboBox)sender).Text == "TEXT")
            {
                dgFields.Rows[rowindex].Cells["length"].Value = 50;
                dgFields.Rows[rowindex].Cells["length"].Style.BackColor = Color.White;
            }
            else
            {
                dgFields.Rows[rowindex].Cells["length"].Value = DBNull.Value;
                dgFields.Rows[rowindex].Cells["length"].Style.BackColor = Color.Gray;
            }
        }

        public void combox_Leave(object sender, EventArgs e)
        {
            ComboBox combox = sender as ComboBox;
            //做完处理，须撤销动态事件
            combox.SelectedIndexChanged -= new EventHandler(ComboBox_SelectedIndexChanged);
        }

        private void dgFields_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            isChange = true;
            if (dgFields.Columns[e.ColumnIndex].Name == "length")
                if (dgFields.Rows[e.RowIndex].Cells[e.ColumnIndex].Value == DBNull.Value)
                    dgFields.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.BackColor = Color.Gray;
                else dgFields.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.BackColor = Color.White;
        }

        private void dgFields_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
            if (dgFields.Columns[e.ColumnIndex].Name == "fields")
            {
                if (e.FormattedValue.ToString().Length > 8)
                {
                    dgFields.Rows[e.RowIndex].ErrorText = "字段长度过长！";
                    e.Cancel = true;
                    return;
                }
                if (e.FormattedValue.ToString().Length == 0)
                {
                    dgFields.Rows[e.RowIndex].ErrorText = "字段名不能为空！";
                    e.Cancel = true;
                    return;
                }
            
            }
            if (dgFields.Columns[e.ColumnIndex].Name == "types")
            {

            }
            if (dgFields.Columns[e.ColumnIndex].Name == "length")
            {
                if (dgFields.Rows[e.RowIndex].Cells["types"].Value.ToString() != "TEXT")
                {
                    if (e.FormattedValue.ToString().Length > 0)
                    {
                        dgFields.Rows[e.RowIndex].ErrorText = "该类型字段不设置长度！";
                        e.Cancel = true;
                        return;
                    }
                }
            }
            dgFields.Rows[e.RowIndex].ErrorText = string.Empty;
        }

        private void dgFields_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            System.Drawing.Rectangle rectangle = new System.Drawing.Rectangle(e.RowBounds.Location.X,
            e.RowBounds.Location.Y,dgFields.RowHeadersWidth - 4,e.RowBounds.Height);

            TextRenderer.DrawText(e.Graphics, (e.RowIndex + 1).ToString(),
                dgFields.RowHeadersDefaultCellStyle.Font,rectangle,
                dgFields.RowHeadersDefaultCellStyle.ForeColor,
                TextFormatFlags.VerticalCenter | TextFormatFlags.Right);
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            if (isChange)
            {
                dt = (DataTable)dgFields.DataSource;
            }
            this.Close();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }







    }
}
