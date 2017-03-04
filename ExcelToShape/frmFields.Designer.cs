namespace ExcelToShape
{
    partial class frmFields
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.dgFields = new System.Windows.Forms.DataGridView();
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnSave = new System.Windows.Forms.Button();
            this.fields = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.types = new System.Windows.Forms.DataGridViewComboBoxColumn();
            this.length = new System.Windows.Forms.DataGridViewTextBoxColumn();
            ((System.ComponentModel.ISupportInitialize)(this.dgFields)).BeginInit();
            this.SuspendLayout();
            // 
            // dgFields
            // 
            this.dgFields.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dgFields.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgFields.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.fields,
            this.types,
            this.length});
            this.dgFields.Location = new System.Drawing.Point(12, 12);
            this.dgFields.Name = "dgFields";
            this.dgFields.RowTemplate.Height = 23;
            this.dgFields.Size = new System.Drawing.Size(482, 250);
            this.dgFields.TabIndex = 0;
            this.dgFields.CellValidating += new System.Windows.Forms.DataGridViewCellValidatingEventHandler(this.dgFields_CellValidating);
            this.dgFields.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgFields_CellValueChanged);
            this.dgFields.EditingControlShowing += new System.Windows.Forms.DataGridViewEditingControlShowingEventHandler(this.dgFields_EditingControlShowing);
            this.dgFields.RowPostPaint += new System.Windows.Forms.DataGridViewRowPostPaintEventHandler(this.dgFields_RowPostPaint);
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(259, 273);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 23);
            this.btnCancel.TabIndex = 1;
            this.btnCancel.Text = "取消";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // btnSave
            // 
            this.btnSave.Location = new System.Drawing.Point(141, 273);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(75, 23);
            this.btnSave.TabIndex = 2;
            this.btnSave.Text = "保存";
            this.btnSave.UseVisualStyleBackColor = true;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // fields
            // 
            this.fields.HeaderText = "字段";
            this.fields.Name = "fields";
            // 
            // types
            // 
            this.types.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.types.HeaderText = "类型";
            this.types.Items.AddRange(new object[] {
            "TEXT",
            "FLOAT",
            "DOUBLE",
            "SHORT",
            "LONG",
            "DATE",
            "BLOB",
            "BASTER",
            "GUID",
            ""});
            this.types.Name = "types";
            this.types.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.types.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            // 
            // length
            // 
            this.length.HeaderText = "长度";
            this.length.Name = "length";
            // 
            // frmFields
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(506, 308);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.dgFields);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "frmFields";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "属性设置";
            this.Load += new System.EventHandler(this.frmFields_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dgFields)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView dgFields;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.DataGridViewTextBoxColumn fields;
        private System.Windows.Forms.DataGridViewComboBoxColumn types;
        private System.Windows.Forms.DataGridViewTextBoxColumn length;
    }
}