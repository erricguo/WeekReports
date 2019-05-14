using System;
using System.Windows.Forms;

namespace SharpUpdate
{
    public partial class SharpUpdateInfoForm : DevExpress.XtraEditors.XtraForm
    {
        public SharpUpdateInfoForm(ISharpUpdatable applicationInfo,SharpUpdateXml updatInfo)
        {
            InitializeComponent();

            if(applicationInfo.ApplicationIcon != null)
            {
                this.Icon = applicationInfo.ApplicationIcon;
            }

            System.Diagnostics.FileVersionInfo ver =
                System.Diagnostics.FileVersionInfo.GetVersionInfo(
                applicationInfo.ApplicationAssembly.Location);

            string[] tmpv = ver.FileVersion.Split('.');
            Version v = new Version(Int32.Parse(tmpv[0]),
                Int32.Parse(tmpv[1]),
                Int32.Parse(tmpv[2]),
                Int32.Parse(tmpv[3]));

            this.Text = applicationInfo.ApplicationName + " ";
            this.lbdes.Text = string.Format("更新版本 Ver:{1}\n目前版本 Ver:{0}", applicationInfo.ApplicationAssembly.GetName().Version,
              updatInfo.Version.ToString());
            this.rtbInfo.Text = updatInfo.Description;
        }

        private void btnBack_Click(object sender, System.EventArgs e)
        {
            this.Close();
        }

        private void rtbInfo_KeyDown(object sender, KeyEventArgs e)
        {
            if (!(e.Control && e.KeyCode == Keys.C))
                e.SuppressKeyPress = true;
        }
    }
}
