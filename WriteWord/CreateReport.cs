using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WriteWord
{
    public partial class CreateReport : Form
    {
        Regex reg = new Regex(@"[^\w]"); 
        public CreateReport()
        {
            InitializeComponent();
            lab_SaveDir.Text = Environment.CurrentDirectory;
        }

        private void btn_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog FBD = new FolderBrowserDialog();
            if (FBD.ShowDialog() == DialogResult.OK)
            {
                lab_SaveDir.Text = FBD.SelectedPath;
            }
        }

        private void tb_TaskId_TextChanged(object sender, EventArgs e)
        {
            tb_TaskId.Text = reg.Replace(tb_TaskId.Text, "");
        }

        private void btn_StartExport_Click(object sender, EventArgs e)
        {
            string TaskId = tb_TaskId.Text;
            if (TaskId.Length != 32)
            {
                MessageBox.Show("TaskID不正确");
                return;
            }
            bool SaveWord, SavePDF;
            SaveWord = cb_ExportWord.Checked;
            SavePDF = cb_ExportPDF.Checked;
            string SaveDir = lab_SaveDir.Text;
            string FilePathRule = tb_PathRule.Text;
            this.Hide();
            ReportProcess RP = new ReportProcess(TaskId, SaveDir, FilePathRule);
            RP.ShowDialog();
            this.Show();
        }
    }
}
