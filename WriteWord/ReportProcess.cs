using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WriteWord
{
    public partial class ReportProcess : Form
    {
        string TaskID, SaveDir, PathRule;
        _Application wordApp;
        _Document wordDoc;
        Object Nothing;
        string FontName = "造字工房悦黑体验版常规体";
        string GreenPicPath;
        string RedPicPath;
        string YellowPicPath;
        float RiskPicWidth = 1.6f;
        float RiskPicHeight = 0.4f;
        float PageHeight = 26f;
        float PageLeftMargin = 1.21f;
        float PageRightMargin = 1.21f;
        object bs = WdBreakType.wdSectionBreakNextPage;
        public ReportProcess(string TaskID, string SaveDir, string PathRule)
        {
            this.TaskID = TaskID;
            this.SaveDir = SaveDir;
            this.PathRule = PathRule;
            InitializeComponent();
            bw_Export.WorkerSupportsCancellation = true;
            bw_Export.RunWorkerAsync();
            
        }

        private void btn_StartExport_Click(object sender, EventArgs e)
        {
            if(bw_Export.IsBusy == true)
                bw_Export.CancelAsync();
            this.Close();
        }

        private void bw_Export_DoWork(object sender, DoWorkEventArgs e)
        {
            ReportResourcesRequired rrr = new ReportResourcesRequired();
            Type t = rrr.GetType();
            MemberInfo[] MIs = t.GetMembers();
            foreach (MemberInfo mi in MIs)
            {
                bw_Export.ReportProgress(250, mi.Name);
            }
        }

        private void bw_Export_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            pb_Export.Value = (int)(e.ProgressPercentage);
            string state = (string)e.UserState;
            lab_info.Text = "ssd";
            if (state.Length > 0)
                lab_info.Text = state;
        }
    }
}
