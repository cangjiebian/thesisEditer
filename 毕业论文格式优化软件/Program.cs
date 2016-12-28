using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.IO;
using Microsoft.Win32;
using Microsoft.Win32.SafeHandles;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
namespace thesisEditer
{
    static class Program
    {
        [DllImport("User32.dll")]
        private static extern bool ShowWindowAsync(IntPtr hWnd, int cmdShow);
        [DllImport("User32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);
        private const int SW_RESTORE = 9;
        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        [STAThread]
        
        static void Main()
        {
            /*开始时间:2015/5/22*/

            /*单例模式运行程序*/
            Process instance = RunningInstance();
            if (instance != null)
            {
                HandleRunningInstance(instance);
                return;
            }

            /*检测office*/

            string OfficeV = "null";
            if (!OfficeIsInstall(out OfficeV))
            {
                MessageBox.Show("检测到本机未安装Microsoft Office或者安装的Microsoft Office版本不在2003-2010之内，无法使用该软件,请先下载安装。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            MainForm.officeVersion = OfficeV;

            /*创建工作目录*/
            string path = System.AppDomain.CurrentDomain.BaseDirectory;
            if (Directory.Exists(path + "office") == false)//如果不存在就创建office文件夹
            {
                Directory.CreateDirectory(path + "office");
            }

            /*检测dsoframer控件是否注册*/
            if (!IsRegistered("00460182-9E5E-11D5-B7C8-B8269041DD57"))
            {
                MessageBox.Show("首次使用,程序会进行控件的安装注册,该控件影响整个程序的使用，所以请勿使用安全软件进行阻止!!!", "欢迎使用", MessageBoxButtons.OK, MessageBoxIcon.Information);


                try
                {
                    File.Copy(path + "dsoframer.ocx", @"c:\windows\system32\dsoframer.ocx", true);
                }
                catch
                {
                }
                System.Diagnostics.Process.Start("regsvr32", @"c:\windows\system32\dsoframer.ocx /s");
                MessageBox.Show("控件注册成功！,", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                //return;

            }
            
            


            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            MainForm mainform = new MainForm();
            Application.Run(mainform);

        }
        private static bool IsRegistered(String CLSID)
        {
            if (String.IsNullOrEmpty(CLSID))
                return false;

            String key = String.Format(@"CLSID\{{{0}}}", CLSID);
            RegistryKey regKey = Registry.ClassesRoot.OpenSubKey(key);
            if (regKey != null)
                return true;
            else
                return false;
        }
        /// <summary>
        /// 通过注册表检测office版本
        /// </summary>
        /// <param name="OfficeVersion">储存office版本的字符串</param>
        /// <returns></returns>
        private static bool OfficeIsInstall(out string OfficeVersion)
        {

            OfficeVersion = "";
            Microsoft.Win32.RegistryKey regKey = null;
            Microsoft.Win32.RegistryKey regSubKey1 = null;
            Microsoft.Win32.RegistryKey regSubKey2 = null;
            Microsoft.Win32.RegistryKey regSubKey3 = null;
            
            regKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry32);
            regSubKey1 = regKey.OpenSubKey(@"SOFTWARE\Microsoft\Office\11.0\Common\InstallRoot", false);
            regSubKey2 = regKey.OpenSubKey(@"SOFTWARE\Microsoft\Office\12.0\Common\InstallRoot", false);
            regSubKey3 = regKey.OpenSubKey(@"SOFTWARE\Microsoft\Office\14.0\Common\InstallRoot", false);

            if (regSubKey3 != null && regSubKey3.GetValue("Path") != null && Directory.Exists(regSubKey3.GetValue("Path").ToString()))
            {
                OfficeVersion = "2010";
                return true;
            }
            else if (regSubKey2 != null && regSubKey2.GetValue("Path") != null && Directory.Exists(regSubKey2.GetValue("Path").ToString()))
            {
                OfficeVersion = "2007";
                return true;
            }
            else if (regSubKey1 != null && regSubKey1.GetValue("Path") != null && Directory.Exists(regSubKey1.GetValue("Path").ToString()))
            {
                OfficeVersion = "2003";
                return true;
            }

            else
            {

                regKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64);
                regSubKey1 = regKey.OpenSubKey(@"SOFTWARE\Microsoft\Office\11.0\Common\InstallRoot", false);
                regSubKey2 = regKey.OpenSubKey(@"SOFTWARE\Microsoft\Office\12.0\Common\InstallRoot", false);
                regSubKey3 = regKey.OpenSubKey(@"SOFTWARE\Microsoft\Office\14.0\Common\InstallRoot", false);
                
                if (regSubKey3 != null && regSubKey3.GetValue("Path") != null)
                {
                    OfficeVersion = "2010";
                    return true;
                }
                else if (regSubKey2 != null && regSubKey2.GetValue("Path") != null)
                {
                    OfficeVersion = "2007";
                    return true;
                }
                else if (regSubKey1 != null && regSubKey1.GetValue("Path") != null)
                {
                    OfficeVersion = "2003";
                    return true;
                }
                else
                {
                    OfficeVersion = "未知";
                    return false;
                }
            }

        }
        /// <summary>
        /// 获取正在运行的实例，没有运行的实例返回null;
        /// </summary>
        public static Process RunningInstance()
        {
            Process current = Process.GetCurrentProcess();
            Process[] processes = Process.GetProcessesByName(current.ProcessName);
            foreach (Process process in processes)
            {
                if (process.Id != current.Id)
                {
                    if (Assembly.GetExecutingAssembly().Location.Replace("/", "//") == current.MainModule.FileName)
                    {
                        return process;
                    }
                }
            }
            return null;
        }
        /// <summary>
        /// 显示已运行的程序。
        /// </summary>
        public static void HandleRunningInstance(Process instance)
        {
            ShowWindowAsync(instance.MainWindowHandle, SW_RESTORE); //显示
            SetForegroundWindow(instance.MainWindowHandle);            //放到前端
        }
    }
}
