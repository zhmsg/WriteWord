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
    public partial class CreateReport : Form
    {
        public CreateReport()
        {
            InitializeComponent();
            lab_SaveDIr.Text = Environment.CurrentDirectory;
        }
    }
}
