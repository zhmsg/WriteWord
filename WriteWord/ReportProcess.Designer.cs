namespace WriteWord
{
    partial class ReportProcess
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
            this.pb_Export = new System.Windows.Forms.ProgressBar();
            this.btn_StartExport = new System.Windows.Forms.Button();
            this.lab_info = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // pb_Export
            // 
            this.pb_Export.Location = new System.Drawing.Point(0, 1);
            this.pb_Export.Name = "pb_Export";
            this.pb_Export.Size = new System.Drawing.Size(557, 23);
            this.pb_Export.TabIndex = 0;
            // 
            // btn_StartExport
            // 
            this.btn_StartExport.Font = new System.Drawing.Font("宋体", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btn_StartExport.Location = new System.Drawing.Point(241, 72);
            this.btn_StartExport.Name = "btn_StartExport";
            this.btn_StartExport.Size = new System.Drawing.Size(76, 34);
            this.btn_StartExport.TabIndex = 11;
            this.btn_StartExport.Text = "取消";
            this.btn_StartExport.UseVisualStyleBackColor = true;
            this.btn_StartExport.Click += new System.EventHandler(this.btn_StartExport_Click);
            // 
            // lab_info
            // 
            this.lab_info.AutoSize = true;
            this.lab_info.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lab_info.Location = new System.Drawing.Point(1, 28);
            this.lab_info.Name = "lab_info";
            this.lab_info.Size = new System.Drawing.Size(32, 16);
            this.lab_info.TabIndex = 12;
            this.lab_info.Text = "lab";
            // 
            // ReportProcess
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(557, 118);
            this.Controls.Add(this.lab_info);
            this.Controls.Add(this.btn_StartExport);
            this.Controls.Add(this.pb_Export);
            this.Name = "ReportProcess";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "ReportProcess";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ProgressBar pb_Export;
        private System.Windows.Forms.Button btn_StartExport;
        private System.Windows.Forms.Label lab_info;
    }
}