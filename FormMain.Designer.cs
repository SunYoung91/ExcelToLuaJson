namespace ExcelExport
{
    partial class FormMain
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
            this.panel1 = new System.Windows.Forms.Panel();
            this.btn_exchange_select = new System.Windows.Forms.Button();
            this.btn_select_none = new System.Windows.Forms.Button();
            this.btn_select_all = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.xlsFileList = new System.Windows.Forms.CheckedListBox();
            this.textLog = new System.Windows.Forms.TextBox();
            this.btn_Export = new System.Windows.Forms.Button();
            this.export_path = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.panel1.Controls.Add(this.btn_exchange_select);
            this.panel1.Controls.Add(this.btn_select_none);
            this.panel1.Controls.Add(this.btn_select_all);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.xlsFileList);
            this.panel1.Location = new System.Drawing.Point(0, 3);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(477, 732);
            this.panel1.TabIndex = 0;
            // 
            // btn_exchange_select
            // 
            this.btn_exchange_select.Location = new System.Drawing.Point(366, 59);
            this.btn_exchange_select.Name = "btn_exchange_select";
            this.btn_exchange_select.Size = new System.Drawing.Size(75, 23);
            this.btn_exchange_select.TabIndex = 4;
            this.btn_exchange_select.Text = "反选";
            this.btn_exchange_select.UseVisualStyleBackColor = true;
            this.btn_exchange_select.Click += new System.EventHandler(this.btn_exchange_select_Click);
            // 
            // btn_select_none
            // 
            this.btn_select_none.Location = new System.Drawing.Point(239, 59);
            this.btn_select_none.Name = "btn_select_none";
            this.btn_select_none.Size = new System.Drawing.Size(75, 23);
            this.btn_select_none.TabIndex = 3;
            this.btn_select_none.Text = "全不选";
            this.btn_select_none.UseVisualStyleBackColor = true;
            this.btn_select_none.Click += new System.EventHandler(this.btn_select_none_Click);
            // 
            // btn_select_all
            // 
            this.btn_select_all.Location = new System.Drawing.Point(119, 59);
            this.btn_select_all.Name = "btn_select_all";
            this.btn_select_all.Size = new System.Drawing.Size(75, 23);
            this.btn_select_all.TabIndex = 2;
            this.btn_select_all.Text = "全选";
            this.btn_select_all.UseVisualStyleBackColor = true;
            this.btn_select_all.Click += new System.EventHandler(this.btn_select_all_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("微软雅黑", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label1.Location = new System.Drawing.Point(3, 63);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(79, 22);
            this.label1.TabIndex = 1;
            this.label1.Text = "文件列表:";
            // 
            // xlsFileList
            // 
            this.xlsFileList.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.xlsFileList.Font = new System.Drawing.Font("微软雅黑", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.xlsFileList.FormattingEnabled = true;
            this.xlsFileList.Location = new System.Drawing.Point(3, 88);
            this.xlsFileList.Name = "xlsFileList";
            this.xlsFileList.Size = new System.Drawing.Size(471, 628);
            this.xlsFileList.TabIndex = 0;
            this.xlsFileList.ThreeDCheckBoxes = true;
            // 
            // textLog
            // 
            this.textLog.Location = new System.Drawing.Point(480, 136);
            this.textLog.Multiline = true;
            this.textLog.Name = "textLog";
            this.textLog.Size = new System.Drawing.Size(562, 520);
            this.textLog.TabIndex = 1;
            // 
            // btn_Export
            // 
            this.btn_Export.Location = new System.Drawing.Point(690, 666);
            this.btn_Export.Name = "btn_Export";
            this.btn_Export.Size = new System.Drawing.Size(154, 53);
            this.btn_Export.TabIndex = 2;
            this.btn_Export.Text = "导出";
            this.btn_Export.UseVisualStyleBackColor = true;
            this.btn_Export.Click += new System.EventHandler(this.btn_Export_Click);
            // 
            // export_path
            // 
            this.export_path.Location = new System.Drawing.Point(483, 52);
            this.export_path.Name = "export_path";
            this.export_path.Size = new System.Drawing.Size(557, 21);
            this.export_path.TabIndex = 3;
            this.export_path.TextChanged += new System.EventHandler(this.client_path_TextChanged);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("微软雅黑", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label2.Location = new System.Drawing.Point(485, 18);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(79, 22);
            this.label2.TabIndex = 2;
            this.label2.Text = "导出目录:";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("微软雅黑", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label3.Location = new System.Drawing.Point(485, 111);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(79, 22);
            this.label3.TabIndex = 4;
            this.label3.Text = "日志信息:";
            // 
            // FormMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1052, 747);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.export_path);
            this.Controls.Add(this.btn_Export);
            this.Controls.Add(this.textLog);
            this.Controls.Add(this.panel1);
            this.Name = "FormMain";
            this.Text = "FormMain";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FormMain_FormClosing);
            this.Load += new System.EventHandler(this.FormMain_Load);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.CheckedListBox xlsFileList;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btn_exchange_select;
        private System.Windows.Forms.Button btn_select_none;
        private System.Windows.Forms.Button btn_select_all;
        private System.Windows.Forms.TextBox textLog;
        private System.Windows.Forms.Button btn_Export;
        private System.Windows.Forms.TextBox export_path;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
    }
}