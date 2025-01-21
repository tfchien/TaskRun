using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
namespace TaskRun
{
    public class BackupTask
    {    
        public List<string> SourcePaths { get; set; } = new List<string>();
        public string SourcePath { get; set; } = "";
        public List<string> SourceFolders { get; set; } = new List<string>();
        public string SourceUser { get; set; } = "";
        public string TargetPath { get; set; }
        public string TargetFolder { get; set; } = "";
        public string TargetFullPath { get; set; } = "";
        public string TargetUser { get; set; } = "";
        public string HostIp { get; set; } = "";
        public string FormDate { get; set; } = "";
        public string ToDate { get; set; } = "";
        public List<string> IncludeFilter { get; set; } = new List<string>();
        public List<string> ExcludeFilter { get; set; } = new List<string>();
        public bool Overwrite { get; set; } = false;
        public int? BufSize { get; set; } = 64;
        public bool Acl { get; set; } = false;
        public string Tool { get; set; }

        /// <summary>
        /// 將以換行符號分隔的字符串轉換為來源路徑列表
        /// </summary>
        /// <param name="sourcePaths">包含多個來源路徑的字符串，使用換行符號進行分隔</param>
        public void SetSourcePaths(string sourcePaths)
        {
            // 清空現有的來源路徑
            SourcePaths.Clear();
            // 使用正則表達式將字符串按換行分割
            SourcePaths.AddRange(Regex.Split(sourcePaths, @"\r?\n|\r"));
        }

        /// <summary>
        /// 將來源路徑列表轉換為以換行符號分隔的字符串
        /// </summary>
        public string SourcePathString
        {
            get => string.Join(Environment.NewLine, SourcePaths);
            // 若需要，也可以添加 set 方法
        }

        /// <summary>
        /// 根據 TargetFolder 的設定，生成對應的日期時間格式的資料夾名稱
        /// 支援輸入日期格式字串 (例如: yyMMdd, yyyyMM)
        /// </summary>
        /// <returns>生成的資料夾名稱</returns>
        public void SetSubFolder(string targetFolder)
        {
            if (string.IsNullOrEmpty(targetFolder))
            {
                TargetFolder = string.Empty;
                return;
            }

            var currentTime = DateTime.Now;

            // 嘗試解析 targetFolder 的日期格式
            var result = Regex.Replace(targetFolder, @"y{2,4}|M{1,2}|d{1,2}|h{1,2}|m{1,2}|s{1,2}", match =>
            {
                return match.Value switch
                {
                    "yyyy" => currentTime.ToString("yyyy"), // 四位年份
                    "yy" => currentTime.ToString("yy"),     // 兩位年份
                    "M" => currentTime.ToString("%M"),      // 不帶前導零的月份
                    "MM" => currentTime.ToString("MM"),     // 帶前導零的月份
                    "d" => currentTime.ToString("%d"),      // 不帶前導零的日期
                    "dd" => currentTime.ToString("dd"),     // 帶前導零的日期
                    "h" => currentTime.ToString("%H"),      // 不帶前導零的24小時制
                    "hh" => currentTime.ToString("HH"),     // 帶前導零的24小時制
                    "m" => currentTime.ToString("%m"),      // 不帶前導零的分鐘
                    "mm" => currentTime.ToString("mm"),     // 帶前導零的分鐘
                    "s" => currentTime.ToString("%s"),      // 不帶前導零的秒
                    "ss" => currentTime.ToString("ss"),     // 帶前導零的秒
                    _ => match.Value                         // 預設保留原值
                };
            });

            TargetFolder = result;
        }


        /// <summary>
        /// 計算日期是該年的第幾週
        /// </summary>
        /// <param name="date">日期</param>
        /// <returns>該年的第幾週</returns>
        private static int GetWeekOfYear(DateTime date)
        {
            var culture = System.Globalization.CultureInfo.CurrentCulture;
            return culture.Calendar.GetWeekOfYear(date, System.Globalization.CalendarWeekRule.FirstDay, DayOfWeek.Monday);
        }

        /// <summary>
        /// 返回備份任務的詳細信息
        /// </summary>
        public override string ToString()
        {
            return $"SourcePaths: {string.Join(", ", SourcePaths)}\n" +                   
                   $"SourceFolders: {string.Join(", ", SourceFolders)}\n" +
                   $"SourcePath: {SourcePath}\n" +
                   $"SourceUser: {SourceUser}\n" +
                   $"TargetPath: {TargetPath}\n" +
                   $"TargetFolder: {TargetFolder}\n" +
                   $"TargetUser: {TargetUser}\n" +
                   $"HostIp: {HostIp}\n" +
                   $"FormDate: {FormDate}\n" +
                   $"ToDate: {ToDate}\n" +
                   $"IncludeFilter: {string.Join(", ", IncludeFilter)}\n" +
                   $"ExcludeFilter: {string.Join(", ", ExcludeFilter)}\n" +
                   $"Overwrite: {Overwrite}\n" +
                   $"BufSize: {BufSize}\n" +
                   $"Acl: {Acl}\n" +
                   $"Tool: {Tool}";
        }
    }
}