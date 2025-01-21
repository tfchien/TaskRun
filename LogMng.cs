using System;
using System.IO;

namespace TaskRun
{
    public class LogMng
    {
        private static string logDirectory;

        public LogMng(string logDirectory)
        {
            LogMng.logDirectory = logDirectory;

            // 確保日誌目錄存在
            EnsureLogDirectoryExists();
        }
        /// <summary>
        /// 初始化靜態日誌管理器，必須在使用之前調用。
        /// </summary>
        public static void Initialize(string directory)
        {
            logDirectory = directory;
            EnsureLogDirectoryExists();
        }

        /// <summary>
        /// 確保日誌目錄存在，如果不存在則創建。
        /// </summary>
        private static void EnsureLogDirectoryExists()
        {
            if (!Directory.Exists(logDirectory))
            {
                Directory.CreateDirectory(logDirectory);
                Console.WriteLine($"已創建日誌目錄：{logDirectory}");
            }
        }

        /// <summary>
        /// 記錄訊息到控制台和日誌文件中。
        /// </summary>
        public static void Log(string message,bool isLog = false)
        {
            Console.WriteLine(message); // 顯示到控制台
            if(!isLog) return;
            string logFilePath = Path.Combine(logDirectory, $"backup_{DateTime.Now:yyyyMMdd}.log");
            string logEntry = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - {message}{Environment.NewLine}";
            File.AppendAllText(logFilePath, logEntry); // 寫入日誌文件
        }
    }
}