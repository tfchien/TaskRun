using System;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;

namespace TaskRun
{
    public class ConfigMng
    { 
        public ConfigMng()
        {
            // 確保日誌目錄存在
            LogMng.Initialize(BackupProgram.LogDirectory);     
        }

        /// <summary>
        /// 從 Excel 文件中載入或初始化備份配置。
        /// </summary>
        public BackupConfig Initialize(string configPath)
        {
            try
            {
                if (File.Exists(configPath))
                {
                    var config = LoadConfigFromExcel(configPath);
                    if (config?.BackupTasks?.Any() == true)
                    {
                        Log($"成功載入配置檔案：{configPath}");
                        return config;
                    }
                    throw new Exception("配置檔案中沒有任何備份任務");
                }
            }
            catch (Exception ex)
            {
                Log($"讀取配置檔案時發生錯誤：{ex.Message}");
                throw; // 將異常重新拋出，讓主程序處理
            }

            //Log($"Excel 檔案內容無效：{configPath}，將使用預設配置。");
            //var defaultConfig = CreateDefaultConfig();
            //SaveConfigToExcel(defaultConfig, configPath);
            //return defaultConfig;
            return null;
        }

        /// <summary>
        /// 從 Excel 檔案載入配置設定。
        /// </summary> 
        public static BackupConfig LoadConfigFromExcel(string configPath)
        {
            var config = new BackupConfig();
            config.BackupTasks = new List<BackupTask>();

            using (var package = new ExcelPackage(new FileInfo(configPath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; // 假設設定在第一個 sheet

                // 取得欄位名稱（第一列）
                var properties = typeof(BackupTask).GetProperties();
                Dictionary<string, int> columnIndexes = new Dictionary<string, int>();

                for (int col = worksheet.Dimension.Start.Column; col <= worksheet.Dimension.End.Column; col++)
                {
                    string columnName = worksheet.Cells[1, col].Value?.ToString();
                    if (!string.IsNullOrEmpty(columnName))
                    {
                        columnIndexes[columnName] = col;
                    }
                }

                // 從第二列開始讀取資料
                for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                {
                    var task = new BackupTask();
                    foreach (var property in properties)
                    {
                        if (columnIndexes.ContainsKey(property.Name))
                        {
                            var cellValue = worksheet.Cells[row, columnIndexes[property.Name]].Value;

                            if (property.Name == nameof(task.SourcePaths) && cellValue != null)
                            {
                                // 使用 SetSourcePaths 方法處理多行來源目錄
                                task.SetSourcePaths(cellValue.ToString());
                            }
                            else if (property.Name == nameof(task.TargetFolder) && cellValue != null)
                            {                              
                                // 使用 SetSubFolder 方法來生成日期時間格式的資料夾名稱
                                task.SetSubFolder(cellValue.ToString());
                                task.TargetFullPath = Path.Combine(task.TargetPath, task.TargetFolder);
                            }
                            else if (property.PropertyType == typeof(List<string>))
                            {
                                if (cellValue != null)
                                {
                                    var list = cellValue.ToString().Split(';').Select(s => s.Trim()).ToList();
                                    property.SetValue(task, list);
                                }
                            }
                            else if (property.PropertyType == typeof(bool))
                            {
                                if (cellValue != null)
                                {
                                    bool boolValue;
                                    if (bool.TryParse(cellValue.ToString(), out boolValue))
                                    {
                                        property.SetValue(task, boolValue);
                                    }
                                    else if (cellValue.ToString().Equals("1"))
                                    {
                                        property.SetValue(task, true);
                                    }
                                    else if (cellValue.ToString().Equals("0"))
                                    {
                                        property.SetValue(task, false);
                                    }
                                }
                            }
                            else if (property.PropertyType == typeof(int?))
                            {
                                if (cellValue != null)
                                {
                                    int intValue;
                                    if (int.TryParse(cellValue.ToString(), out intValue))
                                    {
                                        property.SetValue(task, intValue);
                                    }
                                }
                            }
                            else
                            {
                                if (cellValue != null)
                                {
                                    property.SetValue(task, Convert.ChangeType(cellValue, property.PropertyType));
                                }
                            }
                        }
                    }
                    config.BackupTasks.Add(task);
                }
            } 
           
            return config;
        }

        /// <summary>
        /// 將提供的配置設定保存到 Excel 檔案中。
        /// </summary>
        static void SaveConfigToExcel(BackupConfig config, string configPath)
        {
            using (var package = new ExcelPackage(new FileInfo(configPath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("BackupConfig");

                // 寫入欄位名稱（第一列）
                var properties = typeof(BackupTask).GetProperties();
                for (int i = 0; i < properties.Length; i++)
                {
                    worksheet.Cells[1, i + 1].Value = properties[i].Name;
                }

                // 寫入資料
                for (int i = 0; i < config.BackupTasks.Count; i++)
                {
                    var task = config.BackupTasks[i];
                    for (int j = 0; j < properties.Length; j++)
                    {
                        var property = properties[j];
                        var value = property.GetValue(task);

                        if (property.Name == nameof(task.SourcePaths))
                        {
                            // 使用 SourcePathString 屬性來將來源路徑寫入
                            worksheet.Cells[i + 2, j + 1].Value = task.SourcePathString; // 將多行來源路徑寫入
                        }                    
                        else if (value is List<string> list)
                        {
                            worksheet.Cells[i + 2, j + 1].Value = string.Join(";", list);
                        }
                        else
                        {
                            worksheet.Cells[i + 2, j + 1].Value = value;
                        }                     
                    }
                }

                package.Save();
                Log($"成功保存配置檔案: {configPath}");
            }
        }  
        
        /// <summary>
        /// 創建預設的備份配置。
        /// </summary>
        private BackupConfig CreateDefaultConfig()
        {
            return new BackupConfig
            {
                BackupTasks = new List<BackupTask>
                {
                    new BackupTask
                    {
                        SourcePaths = new List<string> { @"C:\Source" },
                        SourcePath = "",
                        TargetPath = @"C:\Target",
                        TargetFolder = "",
                        TargetFullPath = "",
                        HostIp = "",
                        IncludeFilter = new List<string> { "*.txt" },
                        ExcludeFilter = new List<string> { "web.config", "*.config", "*.tmp" },
                        FormDate = "",
                        ToDate = "",
                        Overwrite = false,
                        BufSize = 64,
                        Acl = false,
                        Tool = "filecopy"
                    }
                }
            };
        }

        /// <summary>
        /// 讀取 INI 配置文件中指定節的指定鍵值。
        /// </summary>
        /// <param name="iniFilePath"></param>
        /// <param name="section"></param>
        /// <param name="key"></param>
        /// <returns></returns>
        public static string ReadIniFile(string iniFilePath, string section, string key)
        {
            var iniData = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            foreach (var line in File.ReadAllLines(iniFilePath))
            {
                var trimmedLine = line.Trim();

                if (string.IsNullOrEmpty(trimmedLine) || trimmedLine.StartsWith(";")) // 忽略空行和註解
                    continue;

                if (trimmedLine.StartsWith("[") && trimmedLine.EndsWith("]")) // 讀取節
                {
                    var sectionName = trimmedLine.Trim('[', ']');
                    iniData["CurrentSection"] = sectionName;
                    continue;
                }

                var keyValue = trimmedLine.Split(new[] { '=' }, 2);
                if (keyValue.Length == 2 && iniData.TryGetValue("CurrentSection", out var currentSection) && currentSection == section)
                {
                    iniData[keyValue[0].Trim()] = keyValue[1].Trim();
                }
            }

            return iniData.TryGetValue(key, out var value) ? value : null;
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