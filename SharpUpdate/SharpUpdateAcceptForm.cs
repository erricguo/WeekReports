using System;
using System.Windows.Forms;

namespace SharpUpdate
{
    public partial class SharpUpdateAcceptForm : DevExpress.XtraEditors.XtraForm
    {
        private ISharpUpdatable applicationInfo;

        private SharpUpdateXml updateInfo;

        private SharpUpdateInfoForm updateInfoForm;


        internal SharpUpdateAcceptForm(ISharpUpdatable applicationInfo, SharpUpdateXml updateInfo)
        {
            InitializeComponent();
            fc.ErrorLog("VSharpUpdateAcceptForm");

            this.applicationInfo = applicationInfo;
            this.updateInfo = updateInfo;

            this.Text = this.applicationInfo.ApplicationName + "";

            if (this.applicationInfo.ApplicationIcon != null)
                this.Icon = this.applicationInfo.ApplicationIcon;

            this.lbNew.Text = string.Format("新版本 Ver:{0}", this.updateInfo.Version.ToString());

        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Yes;
            this.Close();
        }

        private void btnNo_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.No;
            this.Close();
        }

        private void btnDetail_Click(object sender, EventArgs e)
        {
            if (this.updateInfoForm == null)
                this.updateInfoForm = new SharpUpdateInfoForm(this.applicationInfo, this.updateInfo);

            this.updateInfoForm.ShowDialog(this);
        }


    }
}
