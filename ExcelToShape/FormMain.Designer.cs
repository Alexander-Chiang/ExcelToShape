namespace ExcelToShape
{
    partial class FormMain
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.btnTransform = new System.Windows.Forms.Button();
            this.btnSetInput = new System.Windows.Forms.Button();
            this.btnSetOutput = new System.Windows.Forms.Button();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.lblProgress = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.txbX = new System.Windows.Forms.TextBox();
            this.txbY = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.文件ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.设置输入ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.设置输出ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.打开输出文件夹ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.打开日志ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.退出ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.设置ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.设置属性ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.数据预处理ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.帮助ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.关于ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnTransform
            // 
            this.btnTransform.Location = new System.Drawing.Point(368, 206);
            this.btnTransform.Name = "btnTransform";
            this.btnTransform.Size = new System.Drawing.Size(75, 23);
            this.btnTransform.TabIndex = 0;
            this.btnTransform.Text = "开始转换";
            this.btnTransform.UseVisualStyleBackColor = true;
            this.btnTransform.Click += new System.EventHandler(this.button1_Click);
            // 
            // btnSetInput
            // 
            this.btnSetInput.Font = new System.Drawing.Font("微软雅黑", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnSetInput.Location = new System.Drawing.Point(368, 90);
            this.btnSetInput.Name = "btnSetInput";
            this.btnSetInput.Size = new System.Drawing.Size(29, 23);
            this.btnSetInput.TabIndex = 1;
            this.btnSetInput.Text = "...";
            this.btnSetInput.UseVisualStyleBackColor = true;
            this.btnSetInput.Click += new System.EventHandler(this.button2_Click);
            // 
            // btnSetOutput
            // 
            this.btnSetOutput.Font = new System.Drawing.Font("微软雅黑", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnSetOutput.Location = new System.Drawing.Point(368, 135);
            this.btnSetOutput.Name = "btnSetOutput";
            this.btnSetOutput.Size = new System.Drawing.Size(29, 23);
            this.btnSetOutput.TabIndex = 2;
            this.btnSetOutput.Text = "...";
            this.btnSetOutput.UseVisualStyleBackColor = true;
            this.btnSetOutput.Click += new System.EventHandler(this.button3_Click);
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(96, 90);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(266, 21);
            this.textBox1.TabIndex = 3;
            // 
            // textBox2
            // 
            this.textBox2.Location = new System.Drawing.Point(96, 135);
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(266, 21);
            this.textBox2.TabIndex = 4;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(29, 93);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(65, 12);
            this.label1.TabIndex = 5;
            this.label1.Text = "输入路径：";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(29, 140);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(65, 12);
            this.label2.TabIndex = 6;
            this.label2.Text = "输出路径：";
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(31, 206);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(331, 23);
            this.progressBar1.TabIndex = 7;
            // 
            // lblProgress
            // 
            this.lblProgress.AutoSize = true;
            this.lblProgress.Location = new System.Drawing.Point(31, 187);
            this.lblProgress.Name = "lblProgress";
            this.lblProgress.Size = new System.Drawing.Size(0, 12);
            this.lblProgress.TabIndex = 8;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(41, 46);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(53, 12);
            this.label3.TabIndex = 9;
            this.label3.Text = "横坐标：";
            // 
            // txbX
            // 
            this.txbX.Location = new System.Drawing.Point(96, 42);
            this.txbX.Name = "txbX";
            this.txbX.Size = new System.Drawing.Size(87, 21);
            this.txbX.TabIndex = 10;
            this.txbX.Text = "x";
            // 
            // txbY
            // 
            this.txbY.Location = new System.Drawing.Point(275, 43);
            this.txbY.Name = "txbY";
            this.txbY.Size = new System.Drawing.Size(87, 21);
            this.txbY.TabIndex = 12;
            this.txbY.Text = "y";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(221, 46);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(53, 12);
            this.label4.TabIndex = 11;
            this.label4.Text = "纵坐标：";
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.文件ToolStripMenuItem,
            this.设置ToolStripMenuItem,
            this.帮助ToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(480, 25);
            this.menuStrip1.TabIndex = 13;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // 文件ToolStripMenuItem
            // 
            this.文件ToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.设置输入ToolStripMenuItem,
            this.设置输出ToolStripMenuItem,
            this.打开输出文件夹ToolStripMenuItem,
            this.打开日志ToolStripMenuItem,
            this.退出ToolStripMenuItem});
            this.文件ToolStripMenuItem.Name = "文件ToolStripMenuItem";
            this.文件ToolStripMenuItem.Size = new System.Drawing.Size(58, 21);
            this.文件ToolStripMenuItem.Text = "文件(&F)";
            // 
            // 设置输入ToolStripMenuItem
            // 
            this.设置输入ToolStripMenuItem.Name = "设置输入ToolStripMenuItem";
            this.设置输入ToolStripMenuItem.Size = new System.Drawing.Size(160, 22);
            this.设置输入ToolStripMenuItem.Text = "设置输入";
            this.设置输入ToolStripMenuItem.Click += new System.EventHandler(this.设置输入ToolStripMenuItem_Click);
            // 
            // 设置输出ToolStripMenuItem
            // 
            this.设置输出ToolStripMenuItem.Name = "设置输出ToolStripMenuItem";
            this.设置输出ToolStripMenuItem.Size = new System.Drawing.Size(160, 22);
            this.设置输出ToolStripMenuItem.Text = "设置输出";
            this.设置输出ToolStripMenuItem.Click += new System.EventHandler(this.设置输出ToolStripMenuItem_Click);
            // 
            // 打开输出文件夹ToolStripMenuItem
            // 
            this.打开输出文件夹ToolStripMenuItem.Name = "打开输出文件夹ToolStripMenuItem";
            this.打开输出文件夹ToolStripMenuItem.Size = new System.Drawing.Size(160, 22);
            this.打开输出文件夹ToolStripMenuItem.Text = "打开输出文件夹";
            this.打开输出文件夹ToolStripMenuItem.Click += new System.EventHandler(this.打开输出文件夹ToolStripMenuItem_Click);
            // 
            // 打开日志ToolStripMenuItem
            // 
            this.打开日志ToolStripMenuItem.Name = "打开日志ToolStripMenuItem";
            this.打开日志ToolStripMenuItem.Size = new System.Drawing.Size(160, 22);
            this.打开日志ToolStripMenuItem.Text = "打开日志";
            this.打开日志ToolStripMenuItem.Click += new System.EventHandler(this.打开日志ToolStripMenuItem_Click);
            // 
            // 退出ToolStripMenuItem
            // 
            this.退出ToolStripMenuItem.Name = "退出ToolStripMenuItem";
            this.退出ToolStripMenuItem.Size = new System.Drawing.Size(160, 22);
            this.退出ToolStripMenuItem.Text = "退出";
            this.退出ToolStripMenuItem.Click += new System.EventHandler(this.退出ToolStripMenuItem_Click);
            // 
            // 设置ToolStripMenuItem
            // 
            this.设置ToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.设置属性ToolStripMenuItem,
            this.数据预处理ToolStripMenuItem});
            this.设置ToolStripMenuItem.Name = "设置ToolStripMenuItem";
            this.设置ToolStripMenuItem.Size = new System.Drawing.Size(59, 21);
            this.设置ToolStripMenuItem.Text = "设置(&S)";
            // 
            // 设置属性ToolStripMenuItem
            // 
            this.设置属性ToolStripMenuItem.Name = "设置属性ToolStripMenuItem";
            this.设置属性ToolStripMenuItem.Size = new System.Drawing.Size(136, 22);
            this.设置属性ToolStripMenuItem.Text = "设置属性";
            this.设置属性ToolStripMenuItem.Click += new System.EventHandler(this.设置属性ToolStripMenuItem_Click);
            // 
            // 数据预处理ToolStripMenuItem
            // 
            this.数据预处理ToolStripMenuItem.Name = "数据预处理ToolStripMenuItem";
            this.数据预处理ToolStripMenuItem.Size = new System.Drawing.Size(136, 22);
            this.数据预处理ToolStripMenuItem.Text = "数据预处理";
            this.数据预处理ToolStripMenuItem.Click += new System.EventHandler(this.数据预处理ToolStripMenuItem_Click);
            // 
            // 帮助ToolStripMenuItem
            // 
            this.帮助ToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.关于ToolStripMenuItem});
            this.帮助ToolStripMenuItem.Name = "帮助ToolStripMenuItem";
            this.帮助ToolStripMenuItem.Size = new System.Drawing.Size(61, 21);
            this.帮助ToolStripMenuItem.Text = "帮助(&H)";
            // 
            // 关于ToolStripMenuItem
            // 
            this.关于ToolStripMenuItem.Name = "关于ToolStripMenuItem";
            this.关于ToolStripMenuItem.Size = new System.Drawing.Size(100, 22);
            this.关于ToolStripMenuItem.Text = "关于";
            this.关于ToolStripMenuItem.Click += new System.EventHandler(this.关于ToolStripMenuItem_Click);
            // 
            // FormMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(480, 258);
            this.Controls.Add(this.txbY);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.txbX);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.lblProgress);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.textBox2);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.btnSetOutput);
            this.Controls.Add(this.btnSetInput);
            this.Controls.Add(this.btnTransform);
            this.Controls.Add(this.menuStrip1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MainMenuStrip = this.menuStrip1;
            this.MaximizeBox = false;
            this.Name = "FormMain";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Excel2Shape ";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FormMain_FormClosing);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnTransform;
        private System.Windows.Forms.Button btnSetInput;
        private System.Windows.Forms.Button btnSetOutput;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Label lblProgress;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txbX;
        private System.Windows.Forms.TextBox txbY;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem 文件ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 设置ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 设置属性ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 帮助ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 设置输入ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 设置输出ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 打开输出文件夹ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 打开日志ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 退出ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 关于ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 数据预处理ToolStripMenuItem;
    }
}

