using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using DevExpress.XtraEditors;

namespace WeekReports
{
    public partial class MsgForm : DevExpress.XtraEditors.XtraForm
    {
        int FTimeCount = 0;
        int FLimitTime = 0;
        public MsgForm(string xMsg = "", string xCaption = "訊息" ,int xSeconds = -1)
        {
            InitializeComponent();

            me01.Text = xMsg;
            Text = xCaption;
            if(xSeconds != -1 )
            {
                FLimitTime = xSeconds;
                timer1.Enabled = true;
            }

        }

        public MsgForm(string xMsg = "", string xCaption = "訊息", int xSeconds = -1 , int xWidth = 462,int xHeight=279)
        {
            InitializeComponent();

            me01.Text = xMsg;
            Text = xCaption;
            Width = xWidth;
            Height = xHeight;
            if (xSeconds != -1)
            {
                FLimitTime = xSeconds;
                timer1.Enabled = true;
            }
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            FTimeCount++;
            if(FTimeCount >= FLimitTime)
            {
                timer1.Enabled = false;
                Close();
            }
        }

        private void MsgForm_Load(object sender, EventArgs e)
        {


        }

        private void me01_EditValueChanged(object sender, EventArgs e)
        {

        }

        private void MsgForm_Shown(object sender, EventArgs e)
        {
            me01.SelectionStart = me01.Text.Length;
            me01.SelectionLength = 0;
        }
    }
}