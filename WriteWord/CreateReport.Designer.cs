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
            this.lab_SaveDIr = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // lab_TaskId
            // 
            this.lab_TaskId.AutoSize = true;
            this.lab_TaskId.Font = new System.Drawing.Font("宋体", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lab_TaskId.Location = new System.Drawing.Point(66, 28);
            this.lab_TaskId.Name = "lab_TaskId";
            this.lab_TaskId.Size = new System.Drawing.Size(69, 19);
            this.lab_TaskId.TabIndex = 0;
            this.lab_TaskId.Text = "TaskID";
            // 
            // tb_TaskId
            // 
            this.tb_TaskId.Font = new System.Drawing.Font("宋体", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.tb_TaskId.Location = new System.Drawing.Point(161, 25);
            this.tb_TaskId.MaxLength = 32;
            this.tb_TaskId.Name = "tb_TaskId";
            this.tb_TaskId.Size = new System.Drawing.Size(187, 29);
            this.tb_TaskId.TabIndex = 1;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("宋体", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label1.Location = new System.Drawing.Point(66, 72);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(123, 19);
            this.label1.TabIndex = 2;
            this.label1.Text = "文件保存目录";
            // 
            // btn
            // 
            this.btn.Font = new System.Drawing.Font("宋体", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btn.Location = new System.Drawing.Point(208, 64);
            this.btn.Name = "btn";
            this.btn.Size = new System.Drawing.Size(95, 34);
            this.btn.TabIndex = 3;
            this.btn.Text = "浏览";
            this.btn.UseVisualStyleBackColor = true;
            // 
            // lab_SaveDIr
            // 
            this.lab_SaveDIr.AutoSize = true;
            this.lab_SaveDIr.Location = new System.Drawing.Point(70, 110);
            this.lab_SaveDIr.Name = "lab_SaveDIr";
            this.lab_SaveDIr.Size = new System.Drawing.Size(0, 12);
            this.lab_SaveDIr.TabIndex = 4;
            // 
            // CreateReport
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(450, 256);
            this.Controls.Add(this.lab_SaveDIr);
            this.Controls.Add(this.btn);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.tb_TaskId);
            this.Controls.Add(this.lab_TaskId);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "CreateReport";
            this.Text = "生成报告";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lab_TaskId;
        private System.Windows.Forms.TextBox tb_TaskId;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btn;
        private System.Windows.Forms.Label lab_SaveDIr;
    }
}

