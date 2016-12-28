using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;
using Microsoft.Win32;

namespace thesisEditer
{
    public partial class WordErrorForm : Form
    {
        const uint ERROR_ID_1 = 0xc0000005;
        const uint ERROR_ID_2 = 0xc0000374;
        public WordErrorForm()
        {
            InitializeComponent();
        }
        private void ErrorHandle(uint error_id)
        {
            switch (error_id)
            {
                case ERROR_ID_1:
                    DelFile();
                    break;
                case ERROR_ID_2:
                    break;
                default:
                    break;
            }
        }
        private void DelFile()
        {
            if (MessageBox.Show(this, "处理异常后会退出本程序，需要自行重启，是否继续？", "提示", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == System.Windows.Forms.DialogResult.Yes)
            {
                try
                {
                    Process[] myProcess = Process.GetProcessesByName("WINWORD");
                    if (myProcess.Length != 0)
                    {
                        foreach (Process p in myProcess)
                        {

                            if (p.MainWindowTitle == "")
                            {
                                try
                                {
                                    p.Kill();
                                }
                                catch { }

                            }
                        }

                    }
                    Microsoft.Win32.RegistryKey regKey = null;
                    regKey = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Registry32);
                    RenameSubKey(regKey.OpenSubKey(@"Software\Microsoft\Office\word"),"Addins111","Addins");
                    File.Delete(System.Environment.GetEnvironmentVariable("appdata") + @"\microsoft\templates\Normal.dotm");
                    
                    MessageBox.Show(this, "操作成功，程序即将退出！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    Application.Exit();
                }
                catch (Exception e)
                {
                    MessageBox.Show(this, "操作失败"+e, "提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        /// <summary>
        /// Renames a subkey of the passed in registry key since 
        /// the Framework totally forgot to include such a handy feature.
        /// </summary>
        /// <param name="regKey">The RegistryKey that contains the subkey 
        /// you want to rename (must be writeable)</param>
        /// <param name="subKeyName">The name of the subkey that you want to rename
        /// </param>
        /// <param name="newSubKeyName">The new name of the RegistryKey</param>
        /// <returns>True if succeeds</returns>
        public bool RenameSubKey(RegistryKey parentKey,
            string subKeyName, string newSubKeyName)
        {
            CopyKey(parentKey, subKeyName, newSubKeyName);
            parentKey.DeleteSubKeyTree(subKeyName);
            return true;
        }

        /// <summary>
        /// Copy a registry key.  The parentKey must be writeable.
        /// </summary>
        /// <param name="parentKey"></param>
        /// <param name="keyNameToCopy"></param>
        /// <param name="newKeyName"></param>
        /// <returns></returns>
        public bool CopyKey(RegistryKey parentKey,
            string keyNameToCopy, string newKeyName)
        {
            //Create new key
            RegistryKey destinationKey = parentKey.CreateSubKey(newKeyName);

            //Open the sourceKey we are copying from
            RegistryKey sourceKey = parentKey.OpenSubKey(keyNameToCopy);

            RecurseCopyKey(sourceKey, destinationKey);

            return true;
        }

        private void RecurseCopyKey(RegistryKey sourceKey, RegistryKey destinationKey)
        {
            //copy all the values
            foreach (string valueName in sourceKey.GetValueNames())
            {
                object objValue = sourceKey.GetValue(valueName);
                RegistryValueKind valKind = sourceKey.GetValueKind(valueName);
                destinationKey.SetValue(valueName, objValue, valKind);
            }

            //For Each subKey 
            //Create a new subKey in destinationKey 
            //Call myself 
            foreach (string sourceSubKeyName in sourceKey.GetSubKeyNames())
            {
                RegistryKey sourceSubKey = sourceKey.OpenSubKey(sourceSubKeyName);
                RegistryKey destSubKey = destinationKey.CreateSubKey(sourceSubKeyName);
                RecurseCopyKey(sourceSubKey, destSubKey);
            }
        }

        private void buttonX1_Click(object sender, EventArgs e)
        {
            ErrorHandle(ERROR_ID_1);
        }
    }
}
