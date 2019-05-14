using DevExpress.XtraEditors.Controls;
using Microsoft.Win32;
using System;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;


namespace WeekReports
{
    public class fc
    {
        public static string MyDocumentsPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        public static string WRDIR = MyDocumentsPath + "\\WeekReports";
        public static string NodeSoftWare = "Software";
        public static string NodeWR = "WeekReports";
        public static string NodePath = @"HKEY_CURRENT_USER\Software\" + NodeWR;
        public static string INIPath = MyDocumentsPath + "\\WeekReports\\Config.ini"; 


        public static bool IsNewer(string xOldVersion, string xNewVersion)
        {
            bool mResult = false;
            string[] oldVersion = xOldVersion.Split('.');
            string[] newVersion = xNewVersion.Split('.');

            for (int i = 0; i < oldVersion.Length; i++)
            {
                if (Int32.Parse(newVersion[i]) > Int32.Parse(oldVersion[i]))
                {
                    mResult = true;
                }
            }

            return mResult;
        }

        public static object iif(bool xBool,object Obja,object Objb)
        {
            if (xBool)
                return Obja;
            else
                return Objb;
        }

        public static void Msg(string xMsg = "", string xCaption = "訊息")
        {
            MsgForm msg = new MsgForm(xMsg, xCaption, -1);
            msg.ShowDialog();
        }

        public static void Msg(string xMsg = "", string xCaption = "訊息", int xSeconds = -1)
        {
            MsgForm msg = new MsgForm(xMsg, xCaption, xSeconds);
            msg.ShowDialog();
        }

        public static void Msg(string xMsg = "", string xCaption = "訊息", int xSeconds = -1, int xWidth = 462, int xHeight = 279)
        {
            MsgForm msg = new MsgForm(xMsg, xCaption, xSeconds, xWidth, xHeight);
            msg.ShowDialog();
        }

        public static void ErrorLog(string xStr)
        {
            try
            {
                string filepath = WRDIR + "\\Log_" + DateTime.Now.ToString("yyyyMMdd") + ".txt";
                if (!File.Exists(filepath))
                {
                    if (!Directory.Exists(fc.WRDIR))
                    {
                        Directory.CreateDirectory(fc.WRDIR);
                    }
                    File.Create(filepath).Close();
                }
                StreamWriter sw = new StreamWriter(filepath, true, Encoding.Default);
                string mStr = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") + "\t\t" + xStr;
                sw.WriteLine(mStr);
                sw.Close();
            }
            catch (Exception ex)
            {
                Msg(ex.Message.ToString(),"錯誤");
                fc.ErrorLog(ex.Message);
            }
        }

        public static void Restart(string appName)
        {
            System.Threading.Thread thtmp = new System.Threading.Thread(new System.Threading.ParameterizedThreadStart(run));
            thtmp.Start(appName);
        }
        private static void run(Object obj)
        {
            System.Diagnostics.Process ps = new System.Diagnostics.Process();
            ps.StartInfo.FileName = obj.ToString();
            ps.Start();
        }

        public static string makeConnectString()
        {
            string mID = "";
            string mPW = "";
            string mIP = "";
            string mDB = "";

            //打開 子機碼 路徑。
            RegistryKey Reg = Registry.CurrentUser.OpenSubKey(NodeSoftWare, true);
            ////檢查mDB子機碼是否存在，檢查資料夾是否存在。
            if (Reg.GetSubKeyNames().Contains(NodeWR))
            {
                mID = Registry.GetValue(NodePath, "ID", "").ToString();
                mPW = Registry.GetValue(NodePath, "PW", "").ToString();
                mIP = Registry.GetValue(NodePath, "IP", "").ToString();
                mDB = Registry.GetValue(NodePath, "DB", "").ToString();
            }
            Reg.Close();

            string ConnStr = "Data Source = "+mIP+" ;Initial catalog = "+mDB+" ;" +
                             "User id = "+mID+" ; Password = "+mPW;

            return ConnStr;
            //return new SqlConnection(ConnStr);
        }

        public static bool isDirectory(string p)//目錄是否存在
        {
            if (p == "")
            {
                return false;
            }
            return System.IO.Directory.Exists(p);
        }

        public static bool IsFileLocked(string file)
        {
            try
            {
                using (File.Open(file, FileMode.Open, FileAccess.Write, FileShare.None))
                {
                    return false;
                }
            }
            catch (IOException exception)
            {
                var errorCode = Marshal.GetHRForException(exception) & 65535;
                return errorCode == 32 || errorCode == 33;
            }
            catch (Exception)
            {
                return false;
            }
        }
    }
}
