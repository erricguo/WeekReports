using System;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;

namespace SharpUpdate
{
    public class SharpUpdater
    {
        private ISharpUpdatable applicationInfo;
        private BackgroundWorker bgWorker;
        string FLog = "";
        bool FBool;

        public SharpUpdater(ISharpUpdatable applicationInfo)
        {
            this.applicationInfo = applicationInfo;

            this.bgWorker = new BackgroundWorker();
            this.bgWorker.DoWork += new DoWorkEventHandler(bgWorker_DoWork);
            this.bgWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bgWorker_RunWorkerCompleted);
        }

        public void DoUpdate(bool IsFirst)
        {
            FBool = IsFirst;
            if (!this.bgWorker.IsBusy)
                this.bgWorker.RunWorkerAsync(this.applicationInfo);
        }

        private void bgWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            fc.ErrorLog("bgWorker_DoWork");
            ISharpUpdatable application = (ISharpUpdatable)e.Argument;

            if (!SharpUpdateXml.ExistOnServer(application.UpdateXmlLocation))
            {
                e.Cancel = true;
                fc.ErrorLog("e.Cancel = true;");
            }
            else
            {
                e.Result = SharpUpdateXml.Parse(application.UpdateXmlLocation, application.ApplicationID);
                fc.ErrorLog("application.UpdateXmlLocation=" + application.UpdateXmlLocation + " application.ApplicationID=" + application.ApplicationID);
            }
            

        }

        private void bgWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (!e.Cancelled)
            {
                SharpUpdateXml update = (SharpUpdateXml)e.Result;

                System.Diagnostics.FileVersionInfo ver =
                System.Diagnostics.FileVersionInfo.GetVersionInfo(
                this.applicationInfo.ApplicationAssembly.Location);

                string[] tmpv = ver.FileVersion.Split('.');
                Version v = new Version(Int32.Parse(tmpv[0]),
                    Int32.Parse(tmpv[1]),
                    Int32.Parse(tmpv[2]),
                    Int32.Parse(tmpv[3]));

                if (update != null && update.IsNewerThan(this.applicationInfo.ApplicationAssembly.GetName().Version))
                //if (update != null && update.IsNewerThan(v))
                {
                    fc.ErrorLog("Version=" + this.applicationInfo.ApplicationAssembly.GetName().Version);
                    //fc.ErrorLog("Version=" + v);
                    fc.ErrorLog("this.applicationInfo.Context=" + this.applicationInfo.Context);
                    Form f = this.applicationInfo.Context;
                    fc.ErrorLog("f=" + f.ToString());

                    if (FBool)
                    {
                        return;
                    }
                    else if (new SharpUpdateAcceptForm(this.applicationInfo, update).ShowDialog(f) == DialogResult.Yes)
                    {
                        this.DownloadUpdate(update);
                    }
                }
                else
                {
                    if (!FBool)
                        fc.Msg("版本:" + this.applicationInfo.ApplicationAssembly.GetName().Version + " 已是最新版!","提示");
                        //MessageBox.Show("版本:" + v + " 已是最新版!");
                    
                }
            }
        }
        private void DownloadUpdate(SharpUpdateXml update)
        {
            SharpUpdateDownloadForm form = new SharpUpdateDownloadForm(update.Uri, update.MD5, this.applicationInfo.ApplicationIcon);
            fc.ErrorLog("this.applicationInfo.Context=" + this.applicationInfo.Context);
            Form f = this.applicationInfo.Context;
            DialogResult result = form.ShowDialog();

            if(result == DialogResult.OK)
            {
                string currentPath = this.applicationInfo.ApplicationAssembly.Location;
                string newPath = Path.GetDirectoryName(currentPath) + "\\" + update.FileName;

                UpdateAoolication(form.TempFilePath, currentPath, newPath, update.LaunchArgs);
                Application.Exit();
            }
            else if(result == DialogResult.Abort)
            {
                fc.Msg("更新被取消，程式將不會異動", "更新取消");
            }
            else 
            {
                fc.Msg("更新錯誤，請重新嘗試更新", "更新錯誤");
            }
        }      

        private void UpdateAoolication(string tempFilePath, string currentPath, string newPath, string launchArgs)
        {
            string argument = "/C Choice /C Y /N /D Y /T 4 & Del /F /Q \"{0}\" & Choice /C Y /N /D Y /T 2 & Move /Y \"{1}\" \"{2}\" & Start \"\" /D \"{3}\" \"{4}\" {5}";
            
            ProcessStartInfo info = new ProcessStartInfo();
            info.Arguments = string.Format(argument, currentPath, tempFilePath, newPath, Path.GetDirectoryName(newPath), Path.GetFileName(newPath), launchArgs);
            info.WindowStyle = ProcessWindowStyle.Hidden;
            info.CreateNoWindow = true;
            info.FileName = "cmd.exe";
            Process.Start(info);
        }

    }
}