using System;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Collections.Generic;

namespace TaskRun
{
    class BackupProgram
    {
        public static string baseDirectory = "";
        public static string ConfigFile = "backup_config.xlsx";
        public static string UserFile = "backup_user.xlsx";
        public static string LogDirectory = Path.Combine(baseDirectory, "log");        
        static string iniFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "taskrun.ini"); // 將 .ini 檔名設定為專案名稱
    
        public static void Main(string[] args)
        {

            try
            { 

                // 從 .ini 檔案讀取 BaseDirectory
                if (File.Exists(iniFilePath))
                {
                    baseDirectory = ConfigMng.ReadIniFile(iniFilePath, "Paths", "BaseDirectory");
                    if (string.IsNullOrWhiteSpace(baseDirectory))
                    {
                        baseDirectory = AppDomain.CurrentDomain.BaseDirectory; // 預設為當前目錄
                        Log("未在 .ini 檔案中找到 BaseDirectory 設定，使用預設值。");
                    }
                }
                else
                {
                    baseDirectory = AppDomain.CurrentDomain.BaseDirectory; // 預設為當前目錄
                    Log("TaskRun.ini 不存在，使用預設 BaseDirectory。");
                }
                LogDirectory = Path.Combine(baseDirectory, "log");   
                LogMng.Initialize(LogDirectory);  
                Log("=====================================");
                Log("TaskRun 開始執行。");
                Log($"  版本：{System.Reflection.Assembly.GetExecutingAssembly().GetName().Version}");
                //Log($"  運行目錄：{AppDomain.CurrentDomain.BaseDirectory}");
                Log($"  運行目錄：{baseDirectory}");
                Log($"  運行時間：{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")}");
                Log($"  運行參數：{string.Join(" ", args)}");
                Log($"  運行環境：{Environment.OSVersion.VersionString}");
                Log($"  運行機器名稱：{Dns.GetHostName()}");
                Log($"  運行 IP 位址：{Dns.GetHostEntry(Dns.GetHostName()).AddressList.FirstOrDefault(ip => ip.AddressFamily == AddressFamily.InterNetwork)}");
                Log($"  運行使用者：{Environment.UserName}");

                

                
                ConfigFile = ConfigMng.ReadIniFile(iniFilePath, "Paths", "ConfigFile") ?? ConfigFile;       
                UserFile = ConfigMng.ReadIniFile(iniFilePath, "Paths", "UserFile") ?? UserFile;
                

                // 檢查參數是否為空
                if (args == null || args.Length == 0)
                {
                    Log("  使用方式：");
                    Log("    /aes=<Excel檔路徑> - 加密Excel檔案中的密碼");
                    Log("    /cfg=<Excel檔路徑> - 使用Excel檔案作為備份設定");
                    Log("    /force_close - 結束後不等待按鍵輸入");


                    RunEncrypt(UserFile);
                   
                    // 執行備份任務
                    RunBackupTask(ConfigFile);
                    return;
                }
                // 解析命令行參數
                string aesPath = args.FirstOrDefault(arg => arg.StartsWith("/aes="))?.Substring(5);
                if (aesPath != null)
                {
                    RunEncrypt(aesPath);
                    return;
                } 
                
              
                // 處理備份請求 
                try
                {
                    string cfgPath = args.FirstOrDefault(arg => arg.StartsWith("/cfg="))?.Substring(5);
                    ConfigFile = !string.IsNullOrEmpty(cfgPath) ? cfgPath : ConfigFile;
                    // 執行備份任務
                    RunBackupTask(ConfigFile);
                }
                catch (Exception ex)
                {
                    Log($"載入備份配置時發生錯誤：{ex.Message}",true);
                }
            }
            catch (Exception ex)
            {
                Log($"主程序發生錯誤：{ex.Message}");
            }
            finally
            {
                bool force_close = ConfigMng.ReadIniFile(iniFilePath,"Windows","force_close").ToLower() == "true" ;
                if (!force_close && !args.Contains("/force_close"))
                {
                    Log("按任意鍵退出...");
                    Console.ReadKey();                   
                }
            } 
           
        }

        /// <summary>
        /// 執行備份任務，根據提供的配置執行。
        /// </summary>
        public static void RunBackupTask(string ConfigFile)
        {   
            
            bool isLogConfig = ConfigMng.ReadIniFile(iniFilePath,"Windows","isLogConfig").ToLower() == "true";
            
            // 載入或初始化備份配置
            var configMng = new ConfigMng();
            string  configPath = Path.Combine(baseDirectory, ConfigFile);
            // 檢查配置檔案是否存在
            if (!File.Exists(configPath))
            {
                Log($"配置檔案不存在：{configPath}，將創建新的設定。", true);
            }

            BackupConfig config = configMng.Initialize(configPath);
            var copyMng = new CopyMng();
            foreach (var task in config.BackupTasks)
            {
                try
                {
                    copyMng.ExecuteBackup(task); 
                    Log($"備份任務完成：{task.SourcePath} -> {task.TargetPath}",true);
                    if(isLogConfig) 
                    {
                        Log($"備份任務詳細信息:");
                        Log(task.ToString());
                    }
                }
                catch (Exception ex)
                {
                    Log($"備份任務失敗：{task.SourcePath} {ex.Message} :{ex.StackTrace}",true);
                    if(isLogConfig) 
                    {
                        Log($"備份任務詳細信息:",true);
                        Log(task.ToString(),true);
                    }
                    continue;
                }
            }
        }

        public static void RunEncrypt(string UserFile){
             var encryptMng = new EncryptMng();
            // 檢查是否有使用者檔案
             string UserFilePath = Path.Combine(baseDirectory, UserFile);
            if (File.Exists(UserFilePath)){
                try
                {
                    encryptMng.EncryptExcelPasswords(UserFilePath);
                    Log($"Excel 檔案中的密碼加密完成：{UserFile}",true);
                }
                catch (Exception ex)
                {
                    Log($"加密 Excel 檔案中的密碼過程中發生錯誤：{ex.Message}");
                }                
            }
        }

        /// <summary>
        /// 記錄訊息到控制台和日誌文件中。
        /// </summary>
        public static void Log(string message,bool isLog = false)
        {
            LogMng.Log(message,isLog);
        }
    }

}