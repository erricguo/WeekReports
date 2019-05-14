using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Globalization;
using System.Security.Cryptography;
using System.Diagnostics;
using System.Reflection;
using System.Windows.Forms;
using System.Drawing;
using System.IO;
using Nini.Config;
using WeekReports;

namespace SharpUpdate
{
    public class fc
    {
        public static Dictionary<string, Dictionary<string, string>> FROOT = new Dictionary<string, Dictionary<string, string>>();
        public static fc.DBInfo FDBInfo = new fc.DBInfo();
        public static fc.DES FDes = new fc.DES();
        public static fc.UserInfo FUser = new fc.UserInfo();
        public static string MyDocumentsPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        //public static string KUPODIR = MyDocumentsPath + "\\KUPO";
        //public static string COSMOSRESDIR = MyDocumentsPath + "\\KUPO\\COSMOS_RES\\";
        public static string WEEKREPORTS = MyDocumentsPath + "\\WeekReports\\";
        public static string INIPath = MyDocumentsPath + "\\WeekReports\\Config.ini";
        public static string LOGINConfig = "LOGIN";
        public static string VerifyKey = "VERIFICATION";
        public static string DBINFOConfig = "DBINFO";

        public class UserInfo
        {
            public string ID = "";
            public string Name = "";
            public string Email = "";
        }
        public class DES
        {
            // 创建Key
            public string GenerateKey()
            {
                /*
                DESCryptoServiceProvider desCrypto = (DESCryptoServiceProvider)DESCryptoServiceProvider.Create();
                return ASCIIEncoding.ASCII.GetString(desCrypto.Key);*/
                return "這@是A密%k碼";
            }
            // 加密字符串
            public string EncryptString(string sInputString, string sKey)
            {
                byte[] data = Encoding.UTF8.GetBytes(sInputString);
                DESCryptoServiceProvider DES = new DESCryptoServiceProvider();
                DES.Key = ASCIIEncoding.ASCII.GetBytes(sKey);
                DES.IV = ASCIIEncoding.ASCII.GetBytes(sKey);
                ICryptoTransform desencrypt = DES.CreateEncryptor();
                byte[] result = desencrypt.TransformFinalBlock(data, 0, data.Length);
                return BitConverter.ToString(result);
            }
            // 解密字符串
            public string DecryptString(string sInputString, string sKey)
            {                
                string[] sInput = sInputString.Split("-".ToCharArray());
                byte[] data = new byte[sInput.Length];
                if (sInputString!="")
                {
                    for (int i = 0; i < sInput.Length; i++)
                    {
                        data[i] = byte.Parse(sInput[i], NumberStyles.HexNumber);
                    }
                }                
                DESCryptoServiceProvider DES = new DESCryptoServiceProvider();
                DES.Key = ASCIIEncoding.ASCII.GetBytes(sKey);
                DES.IV = ASCIIEncoding.ASCII.GetBytes(sKey);
                ICryptoTransform desencrypt = DES.CreateDecryptor();
                byte[] result = desencrypt.TransformFinalBlock(data, 0, data.Length);
                return Encoding.UTF8.GetString(result);
            }
        } 
        public class DBInfo
        {
            public string IP;
            public string DB;
            public string ID;
            public string PW;
            public DBInfo(string xIP, string xDB, string xID, string xPW)
            {
                // TODO: Complete member initialization
                this.IP = xIP;
                this.DB = xDB;
                this.ID = xID;
                this.PW = xPW;
            }
            public void SetDBInfo(string xIP, string xDB, string xID, string xPW)
            {
                this.IP = xIP;
                this.DB = xDB;
                this.ID = xID;
                this.PW = xPW;
            }
            public DBInfo(){}

        }
        
        public static string makeConnectString(DBInfo xDBInfo)
        {
            string mRes = "Persist Security Info=true" +
                          ";Integrated Security=" +
                          ";Data Source=" + xDBInfo.IP +
                          ";Initial Catalog=" + xDBInfo.DB +
                          ";User ID=" + xDBInfo.ID +
                          ";Password=" + xDBInfo.PW;
            return mRes;
        }
        public static string ZeroatFirst(int xIdx, int xLen)
        {
            string mFormatStr = "";
            for (int i=0;i<xLen;i++)
            {
                mFormatStr += "0";
            }
            return string.Format("{0:" + mFormatStr + "}", xIdx);
        }

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        public static extern void mouse_event(int dwFlags, int dx, int dy, int cButtons, int dwExtraInfo);

        public const int MOUSEEVENTF_LEFTDOWN = 0x02;
        public const int MOUSEEVENTF_LEFTUP = 0x04;
        public const int MOUSEEVENTF_RIGHTDOWN = 0x08;
        public const int MOUSEEVENTF_RIGHTUP = 0x10;

        public static void MouseClick(int x, int y)
        {
            mouse_event(MOUSEEVENTF_LEFTDOWN, x, y, 0, 0);
            mouse_event(MOUSEEVENTF_LEFTUP, x, y, 0, 0);
        }

        public static string GetVersion()
        {
            string s = "";
            try
            {
                //s += System.Deployment.Application.ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString();
                s += Assembly.GetExecutingAssembly().GetName().Version.ToString();
            }
            catch (Exception)
            {
                s = "開發程式階段";
            }
            return s;
        }

        /*public static void Emsg(string xMsg)
        {
            MessageBox.Show(xMsg);
        }
        public static void msg(string xMsg, string xInfo)
        {
            MessageBox.Show(xMsg, xInfo);
        }*/

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

        public class POSXKTable
        {
            public string XK001 = "";
            public string XK002 = "";
            public string XK003 = "";
            public string XK004 = "";
            public string XK005 = "";
            public string XK006 = "";
            public string XK006C = "";
            public double XH004 = 0;
        }

        /// <summary>
        /// 將 Byte 陣列轉換為 Image。
        /// </summary>
        /// <param name="Buffer">Byte 陣列。</param>        
        public static Image BufferToImage(byte[] Buffer)
        {
            if (Buffer == null || Buffer.Length == 0) { return null; }
            byte[] data = null;
            Image oImage = null;
            Bitmap oBitmap = null;
            //建立副本
            data = (byte[])Buffer.Clone();
            try
            {
                MemoryStream oMemoryStream = new MemoryStream(Buffer);
                //設定資料流位置
                oMemoryStream.Position = 0;
                oImage = System.Drawing.Image.FromStream(oMemoryStream);
                //建立副本
                oBitmap = new Bitmap(oImage);
            }
            catch (Exception ex)
            {
                fc.Msg(ex.Message.ToString(),"錯誤");
                fc.ErrorLog(ex.Message);
            }
            //return oImage;
            return oBitmap;
        }

        /// <summary>
        /// 將 Image 轉換為 Byte 陣列。
        /// </summary>
        /// <param name="Image">Image 。</param>
        /// <param name="imageFormat">指定影像格式。</param>        
        public static byte[] ImageToBuffer(Image Image, System.Drawing.Imaging.ImageFormat imageFormat)
        {
            if (Image == null) { return null; }
            byte[] data = null;
            using (MemoryStream oMemoryStream = new MemoryStream())
            {
                //建立副本
                using (Bitmap oBitmap = new Bitmap(Image))
                {
                    //儲存圖片到 MemoryStream 物件，並且指定儲存影像之格式
                    oBitmap.Save(oMemoryStream, imageFormat);
                    //設定資料流位置
                    oMemoryStream.Position = 0;
                    //設定 buffer 長度
                    data = new byte[oMemoryStream.Length];
                    //將資料寫入 buffer
                    oMemoryStream.Read(data, 0, Convert.ToInt32(oMemoryStream.Length));
                    //將所有緩衝區的資料寫入資料流
                    oMemoryStream.Flush();
                }
            }
            return data;
        }

        public static Bitmap ScaleImage(Bitmap pBmp, int pWidth, int pHeight)
        {
            try
            {
                Bitmap tmpBmp = new Bitmap(pWidth, pHeight);
                Graphics tmpG = Graphics.FromImage(tmpBmp);
                tmpG.DrawImage(pBmp,
                                           new Rectangle(0, 0, pWidth, pHeight),
                                           new Rectangle(0, 0, pBmp.Width, pBmp.Height),
                                           GraphicsUnit.Pixel);
                tmpG.Dispose();
                return tmpBmp;
            }
            catch
            {
                return null;
            }
        }

        public static void ErrorLog(string xStr)
        {
            try
            {
                string filepath = WEEKREPORTS + "Update_" + DateTime.Now.ToString("yyyyMMdd") + ".txt";
                if (!File.Exists(filepath))
                {
                    /*if (!Directory.Exists(fc.KUPODIR))
                    {
                        Directory.CreateDirectory(fc.KUPODIR);
                    }*/
                    if (!Directory.Exists(fc.WEEKREPORTS))
                    {
                        Directory.CreateDirectory(fc.WEEKREPORTS);
                    }
                    File.Create(filepath).Close();
                }
                // 建立檔案串流（@ 可取消跳脫字元 escape sequence）
                StreamWriter sw = new StreamWriter(filepath, true, Encoding.Default);
                string mStr = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") + "\t\t" + xStr;
                sw.WriteLine(mStr);
                sw.Close(); 
            }
            catch (Exception ex)
            {                ;
                fc.Msg(ex.Message.ToString(), "錯誤");
                fc.ErrorLog(ex.Message);
            }
        }

        public static void CheckIniPath()
        {
            
            if (!File.Exists(fc.INIPath))
            {
                /*if (!Directory.Exists(fc.KUPODIR))
                {
                    Directory.CreateDirectory(fc.KUPODIR);
                }*/
                if (!Directory.Exists(fc.WEEKREPORTS))
                {
                    Directory.CreateDirectory(fc.WEEKREPORTS);
                }
                File.Create(fc.INIPath).Close();
            }

            IConfigSource source = new IniConfigSource(fc.INIPath);

            foreach (KeyValuePair<string, Dictionary<string, string>> k in FROOT)
            {
                if (source.Configs[k.Key] == null)
                {
                    source.Configs.Add(k.Key);
                    fc.ErrorLog("新增結點:" + k.Key);
                    foreach (KeyValuePair<string, string> j in FROOT[k.Key])
                    {
                        if (source.Configs[k.Key].Get(j.Key, null) == null)
                        {
                            source.Configs[k.Key].Set(j.Key, j.Value);
                            fc.ErrorLog("新增子結點:" + j.Key + " 值:" + j.Value);
                        }
                    }
                }
                else
                {
                    foreach (KeyValuePair<string, string> j in FROOT[k.Key])
                    {
                        if (source.Configs[k.Key].Get(j.Key, null) == null)
                        {
                            source.Configs[k.Key].Set(j.Key, j.Value);
                            fc.ErrorLog("新增子結點:" + j.Key + " 值:" + j.Value);
                        }
                    }
                }
            }
            source.Save(); 
        }
    }
}
