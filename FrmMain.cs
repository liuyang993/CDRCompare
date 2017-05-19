using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace CDRcompare
{
    public partial class FrmMain : Form
    {
        public FrmMain()
        {
            InitializeComponent();
        }

        private void formattedCompareToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            FrmFormattedCompare fc = new FrmFormattedCompare();
            fc.Show();
            fc.MdiParent = this;
        }


    }
}
