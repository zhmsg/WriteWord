namespace WriteWord
{
    partial class CreateReport
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(CreateReport));
            this.lab_TaskId = new System.Windows.Forms.Label();
            this.tb_TaskId = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.btn = new System.Windows.Forms.Button();
            this.lab_SaveDir = new System.Windows.Forms.Label();
            this.cb_ExportWord = new System.Windows.Forms.CheckBox();
            this.cb_ExportPDF = new System.Windows.Forms.CheckBox();
            this.label2 = new System.Windows.Forms.Label();
            this.tb_PathRule = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.btn_StartExport = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // lab_TaskId
            // 
            this.lab_TaskId.AutoSize = true;
            this.lab_TaskId.Font = new System.Drawing.Font("宋体", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lab_TaskId.Location = new System.Drawing.Point(58, 28);
            this.lab_TaskId.Name = "lab_TaskId";
            this.lab_TaskId.Size = new System.Drawing.Size(69, 19);
            this.lab_TaskId.TabIndex = 0;
            this.lab_TaskId.Text = "TaskID";
            // 
            // tb_TaskId
            // 
            this.tb_TaskId.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.tb_TaskId.Location = new System.Drawing.Point(133, 25);
            this.tb_TaskId.MaxLength = 32;
            this.tb_TaskId.Name = "tb_TaskId";
            this.tb_TaskId.Size = new System.Drawing.Size(266, 26);
            this.tb_TaskId.TabIndex = 1;
            this.tb_TaskId.TextChanged += new System.EventHandler(this.tb_TaskId_TextChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("宋体", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label1.Location = new System.Drawing.Point(58, 79);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(123, 19);
            this.label1.TabIndex = 2;
            this.label1.Text = "文件保存目录";
            // 
            // btn
            // 
            this.btn.Font = new System.Drawing.Font("宋体", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btn.Location = new System.Drawing.Point(232, 71);
            this.btn.Name = "btn";
            this.btn.Size = new System.Drawing.Size(95, 34);
            this.btn.TabIndex = 3;
            this.btn.Text = "浏览";
            this.btn.UseVisualStyleBackColor = true;
            this.btn.Click += new System.EventHandler(this.btn_Click);
            // 
            // lab_SaveDir
            // 
            this.lab_SaveDir.AutoSize = true;
            this.lab_SaveDir.Location = new System.Drawing.Point(62, 110);
            this.lab_SaveDir.Name = "lab_SaveDir";
            this.lab_SaveDir.Size = new System.Drawing.Size(0, 12);
            this.lab_SaveDir.TabIndex = 4;
            // 
            // cb_ExportWord
            // 
            this.cb_ExportWord.AutoSize = true;
            this.cb_ExportWord.Checked = true;
            this.cb_ExportWord.CheckState = System.Windows.Forms.CheckState.Checked;
            this.cb_ExportWord.Location = new System.Drawing.Point(119, 144);
            this.cb_ExportWord.Name = "cb_ExportWord";
            this.cb_ExportWord.Size = new System.Drawing.Size(72, 16);
            this.cb_ExportWord.TabIndex = 5;
            this.cb_ExportWord.Text = "导出Word";
            this.cb_ExportWord.UseVisualStyleBackColor = true;
            // 
            // cb_ExportPDF
            // 
            this.cb_ExportPDF.AutoSize = true;
            this.cb_ExportPDF.Checked = true;
            this.cb_ExportPDF.CheckState = System.Windows.Forms.CheckState.Checked;
            this.cb_ExportPDF.Location = new System.Drawing.Point(261, 144);
            this.cb_ExportPDF.Name = "cb_ExportPDF";
            this.cb_ExportPDF.Size = new System.Drawing.Size(66, 16);
            this.cb_ExportPDF.TabIndex = 6;
            this.cb_ExportPDF.Text = "导出PDF";
            this.cb_ExportPDF.UseVisualStyleBackColor = true;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("宋体", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label2.Location = new System.Drawing.Point(58, 183);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(133, 19);
            this.label2.TabIndex = 7;
            this.label2.Text = "文件命名规则 ";
            // 
            // tb_PathRule
            // 
            this.tb_PathRule.Font = new System.Drawing.Font("宋体", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.tb_PathRule.Location = new System.Drawing.Point(64, 218);
            this.tb_PathRule.MaxLength = 40;
            this.tb_PathRule.Name = "tb_PathRule";
            this.tb_PathRule.Size = new System.Drawing.Size(335, 29);
            this.tb_PathRule.TabIndex = 8;
            this.tb_PathRule.Text = "[N]_[T]_报告";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label3.Location = new System.Drawing.Point(197, 190);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(143, 12);
            this.label3.TabIndex = 9;
            this.label3.Text = "[N]受检人姓名 [T]TaskID";
            // 
            // btn_StartExport
            // 
            this.btn_StartExport.Font = new System.Drawing.Font("宋体", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btn_StartExport.Location = new System.Drawing.Point(184, 321);
            this.btn_StartExport.Name = "btn_StartExport";
            this.btn_StartExport.Size = new System.Drawing.Size(95, 34);
            this.btn_StartExport.TabIndex = 10;
            this.btn_StartExport.Text = "开始导出";
            this.btn_StartExport.UseVisualStyleBackColor = true;
            this.btn_StartExport.Click += new System.EventHandler(this.btn_StartExport_Click);
            // 
            // CreateReport
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(464, 367);
            this.Controls.Add(this.btn_StartExport);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.tb_PathRule);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.cb_ExportPDF);
            this.Controls.Add(this.cb_ExportWord);
            this.Controls.Add(this.lab_SaveDir);
            this.Controls.Add(this.btn);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.tb_TaskId);
            this.Controls.Add(this.lab_TaskId);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "CreateReport";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "生成报告";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lab_TaskId;
        private System.Windows.Forms.TextBox tb_TaskId;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btn;
        private System.Windows.Forms.Label lab_SaveDir;
        private System.Windows.Forms.CheckBox cb_ExportWord;
        private System.Windows.Forms.CheckBox cb_ExportPDF;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox tb_PathRule;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button btn_StartExport;
    }
}

