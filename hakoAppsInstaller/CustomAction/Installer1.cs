using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration.Install;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Threading.Tasks;
using IWshRuntimeLibrary;

namespace CustomAction
{
    [RunInstaller(true)]
    public partial class Installer1 : System.Configuration.Install.Installer
    {
        public override void Install(System.Collections.IDictionary stateSaver)
        {
            // Install後の動作
            base.Install(stateSaver);

            // 環境変数PATHの追加
            string currentPath;
            currentPath = System.Environment.GetEnvironmentVariable("path", System.EnvironmentVariableTarget.User);
            string installPath = this.Context.Parameters["InstallPath"];
            string path = installPath + @"\hakoSim\bin;";
                        
            if (currentPath == null)
            {
                currentPath = path;
            }
            else if (currentPath.EndsWith(";"))
            {
                currentPath += path;
            }
            else
            {
                currentPath += ";"+path;
            }

#if DEBUG
            System.Windows.Forms.MessageBox.Show(currentPath);
            System.Windows.Forms.MessageBox.Show(path);
#endif
            // PYTHONPATHの設定
            string pythonPath;

            pythonPath = System.Environment.GetEnvironmentVariable("PYTHONPATH", System.EnvironmentVariableTarget.User);
            string hakopypath = installPath + @"\hakoSim\bin\drone_api\rc;";
#if DEBUG
            System.Windows.Forms.MessageBox.Show(hakopypath);
#endif
            SetUserEnvironmentPathVariable("PYTHONPATH", hakopypath);
#if DEBUG
            System.Windows.Forms.MessageBox.Show(pythonPath);
#endif
            pythonPath = System.Environment.GetEnvironmentVariable("PYTHONPATH", System.EnvironmentVariableTarget.User);
            hakopypath = installPath + @"\hakoSim\bin\drone_api\pymavlink;";
#if DEBUG
            System.Windows.Forms.MessageBox.Show(hakopypath);
#endif
            SetUserEnvironmentPathVariable("PYTHONPATH", hakopypath);

#if DEBUG
            System.Windows.Forms.MessageBox.Show(pythonPath);
#endif
            pythonPath = System.Environment.GetEnvironmentVariable("PYTHONPATH", System.EnvironmentVariableTarget.User);
            hakopypath = installPath + @"\hakoSim\bin\drone_api\libs;";
#if DEBUG
            System.Windows.Forms.MessageBox.Show(hakopypath);
#endif
            SetUserEnvironmentPathVariable("PYTHONPATH", hakopypath);

            pythonPath = System.Environment.GetEnvironmentVariable("PYTHONPATH", System.EnvironmentVariableTarget.User);
            hakopypath = installPath + @"\hakoSim\bin\drone_api\mavsdk;";
#if DEBUG
            System.Windows.Forms.MessageBox.Show(hakopypath);
#endif
            SetUserEnvironmentPathVariable("PYTHONPATH", hakopypath);

            installPath = this.Context.Parameters["InstallPath"];

#if DEBUG
            System.Windows.Forms.MessageBox.Show(installPath);
#endif


            string iniPath = System.IO.Path.Combine(installPath, @"hakoWinAppsAPI\hakoapi.ini");
#if DEBUG
            System.Windows.Forms.MessageBox.Show(iniPath);
#endif

            var updates = new Dictionary<string, string>
            {
                { "HakoWinPath", installPath + @"\hakosim\bin" },
                { "HakoAvatarApp", installPath + @"\hakoAvatar" },
                { "HakoPyPath", installPath + @"\hakosim\bin\drone_api\rc" }
              // 必要に応じて他のキーも追加
            };

            UpdateIniFile(iniPath, updates);

            string desktop = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            string desktopFolder = System.IO.Path.Combine(desktop, "hakoApps-win");

      
            string shortcutLocation = System.IO.Path.Combine(desktopFolder, "インストールフォルダを開く.lnk");

            WshShell shell = new WshShell();
            IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(shortcutLocation);
            shortcut.TargetPath = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "explorer.exe");
            shortcut.Arguments = Context.Parameters["InstallPath"]; // インストール先フォルダを開く
            shortcut.Description = "インストールフォルダを開く";
            shortcut.Save();
    }
    private void SetUserEnvironmentPathVariable(string variableName, string newPath)
    {
      string currentValue = System.Environment.GetEnvironmentVariable(variableName, System.EnvironmentVariableTarget.User);

      if (string.IsNullOrEmpty(currentValue))
      {
        currentValue = newPath;
      }
      else if (currentValue.EndsWith(";"))
      {
        currentValue += newPath;
      }
      else
      {
        currentValue += ";" + newPath;
      }

      System.Environment.SetEnvironmentVariable(variableName, currentValue, System.EnvironmentVariableTarget.User);
    }

    private void UpdateIniFile(string iniPath, Dictionary<string, string> updates)
    {
      var lines = System.IO.File.ReadAllLines(iniPath);
      for (int i = 0; i < lines.Length; i++)
      {
        foreach (var kvp in updates)
        {
          if (lines[i].StartsWith(kvp.Key + "="))
          {
            lines[i] = kvp.Key + "=" + kvp.Value;
          }
        }
      }
      System.IO.File.WriteAllLines(iniPath, lines);
    }



    //public override void Commit(System.Collections.IDictionary savedState)
    //{
    //    //コミット動作
    //    base.Commit(savedState);
    //    System.Windows.Forms.MessageBox.Show(“Commit”);
    //}

    //public override void Rollback(System.Collections.IDictionary savedState)
    //{
    //    //失敗時のロールバック動作
    //    base.Rollback(savedState);
    //    System.Windows.Forms.MessageBox.Show(“Rollback”);
    //}

    public override void Uninstall(System.Collections.IDictionary savedState)
        {
            //Un-install動作
            base.Uninstall(savedState);

            
            // 環境変数PATHを編集
            string currentPath;
            currentPath = System.Environment.GetEnvironmentVariable("path", System.EnvironmentVariableTarget.User);
            string installPath = this.Context.Parameters["InstallPath"];
            string path = installPath + @"\hakoSim\bin;";
            currentPath = currentPath.Replace(path, "");
            
            // 環境変数PATHから削除する
            System.Environment.SetEnvironmentVariable("path", currentPath, System.EnvironmentVariableTarget.User);
#if DEBUG
            System.Windows.Forms.MessageBox.Show(currentPath);
            System.Windows.Forms.MessageBox.Show(path);
#endif

            // PYTHONPATHから箱庭関連を消す
            string pythonPath = System.Environment.GetEnvironmentVariable("PYTHONPATH", System.EnvironmentVariableTarget.User);
            string hakopypath = installPath + @"\hakoSim\bin\drone_api\rc;";
            pythonPath = pythonPath.Replace(hakopypath, "");
#if DEBUG
            System.Windows.Forms.MessageBox.Show(hakopypath);
#endif
            // PYTHONPATHから箱庭関連を削除
            System.Environment.SetEnvironmentVariable("PYTHONPATH", pythonPath, System.EnvironmentVariableTarget.User);

#if DEBUG
            System.Windows.Forms.MessageBox.Show(pythonPath);
            System.Windows.Forms.MessageBox.Show(hakopypath);
#endif
            // PYTHONPATHから箱庭関連を消す
            pythonPath = System.Environment.GetEnvironmentVariable("PYTHONPATH", System.EnvironmentVariableTarget.User);
            hakopypath = installPath + @"\hakoSim\bin\drone_api\pymavlink;";
            pythonPath = pythonPath.Replace(hakopypath, "");
#if DEBUG
            System.Windows.Forms.MessageBox.Show(hakopypath);
#endif
            // PYTHONPATHから箱庭関連を削除
            System.Environment.SetEnvironmentVariable("PYTHONPATH", pythonPath, System.EnvironmentVariableTarget.User);

#if DEBUG
            System.Windows.Forms.MessageBox.Show(pythonPath);
            System.Windows.Forms.MessageBox.Show(hakopypath);
#endif
            // PYTHONPATHから箱庭関連を消す
            pythonPath = System.Environment.GetEnvironmentVariable("PYTHONPATH", System.EnvironmentVariableTarget.User);
            hakopypath = installPath + @"\hakoSim\bin\drone_api\libs;";
            pythonPath = pythonPath.Replace(hakopypath, "");
#if DEBUG
            System.Windows.Forms.MessageBox.Show(hakopypath);
#endif
            // PYTHONPATHから箱庭関連を削除
            System.Environment.SetEnvironmentVariable("PYTHONPATH", pythonPath, System.EnvironmentVariableTarget.User);

#if DEBUG
            System.Windows.Forms.MessageBox.Show(pythonPath);
            System.Windows.Forms.MessageBox.Show(hakopypath);
#endif
            // PYTHONPATHから箱庭関連を消す
            pythonPath = System.Environment.GetEnvironmentVariable("PYTHONPATH", System.EnvironmentVariableTarget.User);
            hakopypath = installPath + @"\hakoSim\bin\drone_api\mavsdk;";
            pythonPath = pythonPath.Replace(hakopypath, "");
#if DEBUG
            System.Windows.Forms.MessageBox.Show(hakopypath);
#endif
            // PYTHONPATHから箱庭関連を削除
            System.Environment.SetEnvironmentVariable("PYTHONPATH", pythonPath, System.EnvironmentVariableTarget.User);

#if DEBUG
            System.Windows.Forms.MessageBox.Show(pythonPath);
            System.Windows.Forms.MessageBox.Show(hakopypath);
#endif
            // デスクトップのパス
            string desktop = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);

            // hakoApps-win フォルダ内のショートカットのパス
            string shortcutPath = Path.Combine(desktop, "hakoApps-win", "インストールフォルダを開く.lnk");

            // ショートカットが存在すれば削除
            if (System.IO.File.Exists(shortcutPath))
            {
              try
              {
                System.IO.File.Delete(shortcutPath);
              }
              catch (Exception ex)
              {
                // ログ出力など必要に応じて追加可能
                System.Diagnostics.Debug.WriteLine("ショートカット削除失敗: " + ex.Message);
              }
            }


    }
  }
}
