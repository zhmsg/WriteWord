using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WriteWord
{
    public partial class ReportProcess : Form
    {
        string TaskID, SaveDir, PathRule;
        public ReportProcess(string TaskID, string SaveDir, string PathRule)
        {
            this.TaskID = TaskID;
            this.SaveDir = SaveDir;
            this.PathRule = PathRule;
            InitializeComponent();
        }

        private void btn_StartExport_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
