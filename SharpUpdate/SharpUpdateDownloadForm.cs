using System;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Net;
using System.Windows.Forms;

namespace SharpUpdate
{
    internal partial class SharpUpdateDownloadForm : DevExpress.XtraEditors.XtraForm
    {
        private WebClient webClient;

        private BackgroundWorker bgWorker;

        private string tempFile;

        private string md5;

        internal string TempFilePath
        {
            get { return this.tempFile; }
        }

        internal SharpUpdateDownloadForm(Uri locationi, string md5, Icon programIcon)
        {
            InitializeComponent();

            if (programIcon != null)
                this.Icon = programIcon;

            tempFile = Path.GetTempFileName();

            this.md5 = md5;

            webClient = new WebClient();
            webClient.DownloadProgressChanged += new DownloadProgressChangedEventHandler(webClient_DownloadProgressChanger);
            webClient.DownloadFileCompleted += new AsyncCompletedEventHandler(webClient_DownloadFileCompleted);

            bgWorker = new BackgroundWorker();
            bgWorker.DoWork += new DoWorkEventHandler(bgWorker_DoWork);
            bgWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bgWorker_RunWorkerCompleted);

            try
            { webClient.DownloadFileAsync(locationi, this.tempFile); }
            catch
            { this.DialogResult = DialogResult.No; this.Close(); }
        }

        private void webClient_DownloadProgressChanger(object sender,DownloadProgressChangedEventArgs e)
        {
            this.progressBarControl1.EditValue = e.ProgressPercentage;
            Application.DoEvents();
            this.lbProcess.Text = string.Format("下載 {0}/{1}", FormatBytes(e.BytesReceived, 1, true), FormatBytes(e.TotalBytesToReceive, 1, true));
        }

        private string FormatBytes(long bytes,int decimalPlaces,bool showByteType)
        {
            double newBytes = bytes;
            string formatString = "{0";
            string byteType = "B";

            if(newBytes > 1024 && newBytes < 1048576)
            {
                newBytes /= 1024;
                byteType = "KB";
            }
            else if (newBytes > 1048576 && newBytes < 1073741824)
            {
                newBytes /= 1048576;
                byteType = "MB";
            }
            else
            {
                newBytes /= 1073741824;
                byteType = "GB";
            }

            if(decimalPlaces > 0)
                formatString += ":0.";

            for (int i = 0; i < decimalPlaces; i++)
                formatString += "0";

            formatString += "}";

            if (showByteType)
                formatString += byteType;

            return string.Format(formatString, newBytes);

        }

        private void webClient_DownloadFileCompleted(object sender, AsyncCompletedEventArgs e)
        {
            
            if(e.Error != null)
            {
                fc.ErrorLog("更新檔案中 e.Error");
                this.DialogResult = DialogResult.No;
                this.Close();
            }
            else if(e.Cancelled)
            {
                fc.ErrorLog("更新檔案中 e.Cancelled");
                this.DialogResult = DialogResult.Abort;
                this.Close();
            }
            else
            {
                lbProcess.Text = "更新檔案中...";
                //progressBarControl1.Style = ProgressBarStyle.Marquee;
                fc.ErrorLog("更新檔案中");
                bgWorker.RunWorkerAsync(new string[] { this.tempFile, this.md5 });
            }
        }

        private void bgWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            string file = ((string[])e.Argument)[0];
            string updateMd5 = ((string[])e.Argument)[1];

            if (Hasher.HashFile(file, HashType.MD5) != updateMd5.ToLower())
            {
                 fc.ErrorLog("e.Result = DialogResult.No");
                e.Result = DialogResult.No;
            }
            else
            {
                fc.ErrorLog("e.Result = DialogResult.OK");
                e.Result = DialogResult.OK;
            }

        }

        private void bgWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            fc.ErrorLog(e.Result.ToString());
            this.DialogResult = (DialogResult)e.Result;
            this.Close();
        }

        private void SharpUpdateDownloadForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            if(webClient.IsBusy)
            {
                webClient.CancelAsync();
                this.DialogResult = DialogResult.Abort;
            }

            if (bgWorker.IsBusy)
            {
                bgWorker.CancelAsync();
                this.DialogResult = DialogResult.Abort;
            }
        }

     
    }

}
