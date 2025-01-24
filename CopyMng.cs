using System;
using System.Diagnostics;
using System.IO;
using System.Security.AccessControl; // DirectoryInfo
using System.Net; // 引入 System.Net 命名空間
using System.Net.Sockets; 
using System.Linq;
using System.Text.RegularExpressions;
using System.Collections.Concurrent; 
using System.Security.Cryptography;
using System.Text; 

namespace TaskRun
{
    public class CopyMng
    {
        private readonly List<string> localIPs;

        public CopyMng()
        {           
            // 確保日誌目錄存在
            LogMng.Initialize(BackupProgram.LogDirectory);         
            this.localIPs = GetLocalIPAddresses();
        }
 

        /// <summary>
        /// 執行單一備份任務。
        /// </summary>
        public void ExecuteBackup(BackupTask task)
        {
            // 驗證 Host IP
            if (!string.IsNullOrWhiteSpace(task.HostIp) && !localIPs.Contains(task.HostIp, StringComparer.OrdinalIgnoreCase))
            {
                Log($"任務 HostIp '{task.HostIp}' 與本機 IP 不符 ({string.Join(", ", localIPs)})，跳過...");
                return;
            }

            // 連接到目標網路路徑
            if (!ConnectNetworkPath(task.TargetPath, task.TargetUser))
            {
                Log($"無法連接到網路路徑 {task.TargetPath}");
                throw new DirectoryNotFoundException($"無法連接到網路路徑{task.TargetPath}"); 
            }

            if (task.Tool.ToLower() == "delete")
            {
                try
                {
                    ExecuteBackupTask(task);
                }
                finally
                {
                    DisconnectNetworkPath(task.TargetPath);
                }
                return;
            }
                

            try
            {
                foreach (var sourcePath in task.SourcePaths)
                {
                    task.SourcePath = sourcePath;

                    // 連接到來源網路路徑  
                    if (!ConnectNetworkPath(task.SourcePath, task.SourceUser))
                    {
                        Log($"無法連接到網路路徑 {task.SourcePath}，跳過...");
                        continue;
                    }
                   

                    try
                    {
                        // 執行備份邏輯
                        ExecuteBackupForSource(task);
                    }
                    finally
                    {
                        DisconnectNetworkPath(task.SourcePath);                        
                    }
                }
            }
            finally
            {
                DisconnectNetworkPath(task.TargetPath);
            }
        }

        private void ExecuteBackupForSource(BackupTask task)
        { 
            if (!Directory.Exists(task.TargetPath))
            {
                Log($"目標路徑不存在：{task.TargetPath}，備份任務中止。",true);
                throw new DirectoryNotFoundException($"目標路徑不存在：{task.TargetPath}");
            }

            if (!HasWritePermission(task.TargetPath))
            {
                Log($"目標路徑 '{task.TargetPath}' 沒有寫入權限，備份任務中止。",true);
                throw new UnauthorizedAccessException($"目標路徑 '{task.TargetPath}' 沒有寫入權限");
            }

            if (task.Tool.ToLower() == "delete")
            {
                Log($"刪除目標：{task.TargetPath}，執行ExecuteBackupTask");
                ExecuteBackupTask(task);
                return;
            }


            if (!Directory.Exists(task.SourcePath))
            {
                Log($"來源路徑不存在：{task.SourcePath}，備份任務中止。",true);
                throw new DirectoryNotFoundException($"來源路徑不存在：{task.SourcePath}");
            }
            

            // 如果有設定 TargetFolder，將其合併到 TargetPath
            if (!string.IsNullOrWhiteSpace(task.TargetFolder))
            {
                
                if (!Directory.Exists(task.TargetFullPath))
                {
                    Directory.CreateDirectory(task.TargetFullPath);
                    Log($"目標資料夾 {task.TargetFullPath} 已建立。");
                }
                Log($"目標資料夾: {task.TargetPath} 更新為：{task.TargetFullPath}");
                
            }else
            {
                task.TargetFullPath = task.TargetPath;
            }

            // 如果 SourceFolders 為空，直接執行備份
            if (task.SourceFolders == null || task.SourceFolders.Count == 0)
            {
                ExecuteBackupTask(task);
                return;
            }
            string sourcePath = task.SourcePath;
            // 遍歷每個子資料夾
            foreach (var subFolder in task.SourceFolders)
            {
                var fullSourcePath = Path.Combine(sourcePath, subFolder);

                if (!Directory.Exists(fullSourcePath))
                {
                    Log($"來源資料夾不存在：{fullSourcePath}，跳過...");
                    continue;
                }

                task.SourcePath = fullSourcePath;
                ExecuteBackupTask(task);
            }
            task.SourcePath = sourcePath;
        }

        /// <summary>
        /// 執行單一備份邏輯。
        /// </summary>
        static void ExecuteBackupTask(BackupTask task)
        {
            if ("delete".Equals(task.Tool, StringComparison.CurrentCultureIgnoreCase))   
                Log($"處理刪除目標：{task.TargetFullPath}，使用工具：{task.Tool}",true);
            else
                Log($"處理來源：{task.SourcePath} 到目標：{task.TargetFullPath}，使用工具：{task.Tool}", true);

            try
            {
                switch (task.Tool.ToLower())
                {
                    case "filecopy":
                        UseFileCopy(task);
                        break;
                    case "xcopy":
                        UseXCopy(task);
                        break;
                    case "robocopy":
                        UseRoboCopy(task);
                        break;
                    case "fastcopy":
                        UseFastCopy(task);
                        break;
                    case "delete":
                        UseDelete(task);
                        break;
                    default:
                        Log("未知的工具選項，將使用 File.Copy 作為預設。");
                        UseFileCopy(task);
                        break;
                }
            }
            catch (Exception ex)
            {
                Log($"備份失敗：來源 {task.SourcePath} 使用 {task.Tool}，原因：{ex.Message}",true);
                throw; // 異常重新拋出
            }
        }
 
        static void UseDelete(BackupTask task)
        {
            try
            {
                if ((task.ExcludeFilter == null || task.ExcludeFilter.Count ==0 )
                    && (task.IncludeFilter == null || task.IncludeFilter.Count ==0 )
                    && string.IsNullOrWhiteSpace(task.FormDate) && string.IsNullOrWhiteSpace(task.ToDate))                                                              
                {
                    //直接刪除資料夾 task.TargetFullPath
                    try
                    {
                        Directory.Delete(task.TargetFullPath, true);
                        Log($"刪除資料夾：{task.TargetFullPath}");
                    }
                    catch (Exception ex)
                    {
                        Log($"刪除資料夾失敗：{task.TargetFullPath}，原因：{ex.Message}",true);                        
                    }
                    return;

                }
                // Combine Include and Exclude filters if they exist
                var files = GetFilesRecursive(task.TargetFullPath, "*"); // Get all files initially
                
                // Include filter
                if (task.IncludeFilter != null && task.IncludeFilter.Any())
                {
                    files = files.Where(file => task.IncludeFilter.Any(include => MatchesPattern(file, include)));
                    Log($"Include篩選: {string.Join(", ", task.IncludeFilter)}");
                }

                // Exclude filter
                if (task.ExcludeFilter != null && task.ExcludeFilter.Any())
                {
                    files = files.Where(file => !task.ExcludeFilter.Any(exclude => MatchesPattern(file, exclude)));
                    Log($"Exclude篩選: {string.Join(", ", task.ExcludeFilter)}");
                }
                
                // Date filter
                if (!string.IsNullOrWhiteSpace(task.FormDate) || !string.IsNullOrWhiteSpace(task.ToDate) )
                {                    
                    DateTime? fromDate = ParseRelativeDate(task.FormDate)?.Date;
                    DateTime? toDate = ParseRelativeDate(task.ToDate)?.Date; 
                    files = files.Where(f =>
                        (!fromDate.HasValue || File.GetLastWriteTime(f).Date >= fromDate.Value) &&
                        (!toDate.HasValue || File.GetLastWriteTime(f).Date <= toDate.Value));
                    
                    Log($"日期篩選: {fromDate?.ToString("yyyy-MM-dd")} - {toDate?.ToString("yyyy-MM-dd")}");

                }

                // Delete files
                foreach (var file in files)
                {
                    try
                    {
                        File.Delete(file);
                        Log($"刪除檔案：{file}");
                    }
                    catch (Exception ex)
                    {
                        Log($"刪除檔案失敗：{file}，原因：{ex.Message}",true);
                    }
                }
            }
            catch (Exception ex)
            {
                Log($"刪除檔案失敗：原因：{ex.Message}");
                throw; // 將異常重新拋出，讓主程序處理
            }
        }

        static void UseFileCopy(BackupTask task)
        {
            
            try
            {
                // Combine Include and Exclude filters if they exist
                var files = GetFilesRecursive(task.SourcePath, "*"); // Get all files initially
                
                // Include filter
                if (task.IncludeFilter != null && task.IncludeFilter.Any())
                {
                    files = files.Where(file => task.IncludeFilter.Any(include => MatchesPattern(file, include)));
                    Log($"Include篩選: {string.Join(", ", task.IncludeFilter)}");
                }

                // Exclude filter
                if (task.ExcludeFilter != null && task.ExcludeFilter.Any())
                {
                    files = files.Where(file => !task.ExcludeFilter.Any(exclude => MatchesPattern(file, exclude)));
                    Log($"Exclude篩選: {string.Join(", ", task.ExcludeFilter)}");
                }
                
                // Date filter
                if (!string.IsNullOrWhiteSpace(task.FormDate) || !string.IsNullOrWhiteSpace(task.ToDate) )
                {                    
                    DateTime? fromDate = ParseRelativeDate(task.FormDate)?.Date;
                    DateTime? toDate = ParseRelativeDate(task.ToDate)?.Date; 
                    files = files.Where(f =>
                        (!fromDate.HasValue || File.GetLastWriteTime(f).Date >= fromDate.Value) &&
                        (!toDate.HasValue || File.GetLastWriteTime(f).Date <= toDate.Value));
                    
                    Log($"日期篩選: {fromDate?.ToString("yyyy-MM-dd")} - {toDate?.ToString("yyyy-MM-dd")}");

                }


                // Copy files
                const long LargeFileSizeThreshold = 1024L * 1024 * 1024; // 1 GB

                foreach (var file in files)
                {
                    string targetFile = Path.Combine(task.TargetFullPath, Path.GetFileName(file));

                    if (File.Exists(targetFile))
                    {
                        if(!task.Overwrite)
                        {
                            var sourceInfo = new FileInfo(file);
                            var targetInfo = new FileInfo(targetFile);
                            if (sourceInfo.Length == targetInfo.Length && sourceInfo.LastWriteTime == targetInfo.LastWriteTime)
                            {
                                Log($"跳過相同檔案：{targetFile}");
                                continue;
                            }
                        }else
                        {
                            //覆寫前先刪除 
                            try
                            {
                                File.Delete(targetFile);
                                Log($"覆寫檔案：{targetFile}");
                            }
                            catch(Exception ex)
                            {
                                Log($"刪除檔案失敗：{targetFile}，原因：{ex.Message}",true);
                                continue; // 刪除失敗則跳過
                            }
                        }
                    }

          
                    var sourceFileInfo = new FileInfo(file);

                    if (sourceFileInfo.Length > LargeFileSizeThreshold)
                    {
                        Log($"檔案過大，使用分塊方式複製：{file}");
                        CopyLargeFile(file, targetFile);
                    }
                    else
                    {
                        Log($"複製：{file} -> {targetFile}");
                        // 複製檔案
                        File.Copy(file, targetFile, task.Overwrite);
                        if (task.Acl)
                        { 
                            // 使用 FileInfo 來獲取和設置 ACL
                            FileInfo targetFileInfo = new FileInfo(targetFile);

                            // 取得來源檔案的 ACL
                            FileSecurity fs = sourceFileInfo.GetAccessControl();

                            // 將來源檔案的 ACL 套用到目標檔案
                            targetFileInfo.SetAccessControl(fs);
                            Log($"已複製 ACL：{file} -> {targetFile}",true);
                        }
                        
                        
                    }
                }
            }
            catch (Exception ex)
            {
                Log($"複製檔案失敗：原因：{ex.Message}");
                throw; // 將異常重新拋出，讓主程序處理
            }
        
        }

        static void CopyLargeFile(string sourcePath, string targetPath, int BufSize =  64 )
        {
            
            int BufferSize = 1024 * 1024 * BufSize; // 64 MB
            byte[] buffer = new byte[BufferSize];

            using (FileStream source = new FileStream(sourcePath, FileMode.Open, FileAccess.Read))
            using (FileStream target = new FileStream(targetPath, FileMode.Create, FileAccess.Write))
            {
                int bytesRead;
                while ((bytesRead = source.Read(buffer, 0, buffer.Length)) > 0)
                {
                    target.Write(buffer, 0, bytesRead);
                }
            }

            Log($"完成分塊複製：{sourcePath} -> {targetPath}");
        }
    
        static void UseXCopy(BackupTask task)
        {
            
            // 構建基本參數
            string arguments = $" /E /H /Y ";
            
            // 構建日期參數
            DateTime? fromDate = ParseRelativeDate(task.FormDate);                 
            if (fromDate?.CompareTo(DateTime.MinValue) > 0)
            {
                arguments += $" /D:{fromDate:yyyy-MM-dd}";
            }

            if(task.Acl)
            {
                arguments += " /O ";
            }

            // 加入檔案名稱條件與副檔名篩選
            String includeFilter = (task.IncludeFilter != null && task.IncludeFilter.Any())? task.IncludeFilter[0]: "";
            string sourcePathWithFilter = Path.Combine(task.SourcePath, includeFilter);

            // 加入來源與目標路徑
            arguments = $"\"{sourcePathWithFilter}\" \"{task.TargetFullPath}\" " + arguments;

            // 如果 Overwrite 為 true，先清除目標檔案
            if (task.Overwrite)
            {
                Log("覆蓋模式啟用，清除目標路徑檔案...");
                foreach (var file in Directory.GetFiles(task.TargetFullPath))
                {
                    try
                    {
                        File.Delete(file);
                        Log($"已刪除：{file}");
                    }
                    catch (Exception ex)
                    {
                        Log($"刪除檔案失敗：{file}，原因：{ex.Message}");
                        throw; // 將異常重新拋出，讓主程序處理
                    }
                }
            }

            // 呼叫 XCOPY
            RunExternalTool("xcopy", arguments);
        }

        static void UseRoboCopy(BackupTask task)
        {
            // 構建基本參數
            string arguments = "";

            // 構建日期參數 (robocopy 使用 /MINAGE 和 /MAXAGE)
            DateTime? fromDate = ParseRelativeDate(task.FormDate);
            DateTime? toDate = ParseRelativeDate(task.ToDate);

            if (fromDate.HasValue)
            {
                TimeSpan minAge = DateTime.Now - fromDate.Value;
                arguments += $" /MINAGE:{minAge.Days}";
            }
            if (toDate.HasValue)
            {
                TimeSpan maxAge = DateTime.Now - toDate.Value;
                arguments += $" /MAXAGE:{maxAge.Days}";
            }

            // ACL 複製
            if (task.Acl)
            {
                arguments += " /COPY:SOU "; // 複製安全性、擁有者和稽核資訊
            }
            else
            {
                arguments += " /COPY:DAT "; // 只複製資料、屬性和時間戳記 (預設)
            }

            // 包含子目錄、隱藏檔案和空目錄
            arguments += " /E ";


            // 加入檔案名稱條件與副檔名篩選
            string includeFilter = (task.IncludeFilter != null && task.IncludeFilter.Any()) ? task.IncludeFilter[0] : "*";

            // 加入來源與目標路徑
            arguments = $"\"{task.SourcePath}\" \"{task.TargetFullPath}\" \"{includeFilter}\" " + arguments;

            // 覆蓋模式處理 (robocopy 使用 /PURGE 或 /MIR)
            if (task.Overwrite)
            {
                Log("覆蓋模式啟用，使用 /PURGE 清除目標路徑多餘檔案...");
                arguments += " /PURGE "; // 刪除目標中來源沒有的檔案
            }

            //其他參數
            arguments += " /R:0 /W:0 /NP"; //不重試、不等待、不顯示進度

            try
            {
                // 呼叫 robocopy
                RunExternalTool("robocopy", arguments);
            }
            catch (Exception ex)
            {
                Log($"執行 robocopy 失敗：原因：{ex.Message}");
                throw; // 將異常重新拋出，讓主程序處理
            }
        }

        static void UseFastCopy(BackupTask task)
        {
            // FastCopy 可執行檔名稱（假設在 PATH 中，或需指定完整路徑）
            string fastCopyExecutable = "fastcopy";

            // 初始參數設置
            string arguments = "/cmd=diff /force_close /no_confirm_stop /speed=full";
        
            // 如果需要覆蓋，加入覆蓋相關選項
            if (task.Overwrite)
            {
                arguments += " /recreate";
            }

            if(task.Acl)
            {
                arguments += " /acl_copy";
            }

            if(task.BufSize.HasValue)
            { 
                arguments += $" /bufsize={task.BufSize}M";
            }

            // 設定log檔(依BackupProgram.LogDirectory路徑)
            string logFileName = $"fastcopy_{DateTime.Now:yyyyMMdd}.log";
            arguments += $" /logfile=\"{Path.Combine(BackupProgram.LogDirectory, logFileName)}\"";

            // 日期篩選（如果提供）
            DateTime? fromDate = ParseRelativeDate(task.FormDate);                   
            if (fromDate?.CompareTo(DateTime.MinValue) > 0)
            {
                arguments += $" /from_date=\"{fromDate:yyyyMMdd}\"";
            }
            DateTime? toDate = ParseRelativeDate(task.ToDate);
            if (toDate?.CompareTo(DateTime.MinValue) > 0)
            {
                arguments += $" /to_date=\"{toDate:yyyyMMdd}\"";
            }

            // 加入檔案過濾條件
            if (task.IncludeFilter != null && task.IncludeFilter.Any())
            {
                arguments += $" /include=\"{string.Join(";", task.IncludeFilter)}\"";
            }

            if (task.ExcludeFilter != null && task.ExcludeFilter.Any())
            {
                arguments += $" /exclude=\"{string.Join(";", task.ExcludeFilter)}\""; 
            }

            // 添加來源路徑與目標路徑
            string sourcePathWithFilter = Path.Combine(task.SourcePath, "*");
            arguments += $" \"{sourcePathWithFilter}\" /to=\"{task.TargetFullPath}\"";

            // 執行 FastCopy
            RunExternalTool(fastCopyExecutable, arguments);
        }


        static void RunExternalTool(string toolName, string arguments)
        {
            try
            {
                Log($"執行 {toolName} 工具，參數：{arguments}");
                var process = new Process
                {
                    StartInfo = new ProcessStartInfo
                    {
                        FileName = toolName,
                        Arguments = arguments,
                        RedirectStandardOutput = true,
                        RedirectStandardError = true,
                        UseShellExecute = false,
                        CreateNoWindow = true
                    }
                };

                process.Start();
                string output = process.StandardOutput.ReadToEnd();
                string error = process.StandardError.ReadToEnd();
                process.WaitForExit();

                Log("輸出：");
                Log(output);
                if (!string.IsNullOrEmpty(error))
                {
                    Log("錯誤：");
                    Log(error);
                }
            }
            catch (Exception ex)
            {
                Log($"執行 {toolName} 時發生錯誤：{ex.Message}");
                throw; // 將異常重新拋出，讓主程序處理
            }
        }
        
        
        static bool ConnectNetworkPath(string path, string? user)
        {
            // 判斷是否為本機路徑
            if (Path.IsPathRooted(path) && !path.StartsWith(@"\\")) // 檢查是否為磁碟根目錄且非 UNC 路徑
            {
                Log($"路徑 '{path}' 為本機路徑，無需建立網路連線。");
                return true; // 本機路徑視為已連線
            }

            if(Directory.Exists(path))
            {
                Log($"路徑 '{path}' 已存在，無需建立網路連線。");
                return true; // 已存在的路徑視為已連線
            }

            string password = EncryptMng.GetDecryptedPassword(user);

            if (string.IsNullOrWhiteSpace(user) || string.IsNullOrWhiteSpace(password))
            {
                Log($"網路路徑 '{path}' 未提供使用者名稱或密碼。");
                // 判斷是否需要連線，如果 path 是 UNC 路徑，則判斷是否可存取
                if (path.StartsWith(@"\\"))
                {
                    if (CanAccessDirectory(path))
                    {
                        Log($"可存取網路路徑 '{path}'，無需提供使用者名稱或密碼。");
                        return true;
                    }
                    else
                    {
                        Log($"無法存取網路路徑 '{path}'，請檢查路徑或提供使用者名稱和密碼。");
                        return false;
                    }
                }
                else
                {
                    Log($"非網路路徑，也未提供使用者名稱或密碼，跳過連線檢查。");
                    return true;
                }

            }

            string command = $"/C net use \"{path}\" /user:\"{user}\" \"{password}\" /persistent:no";
            Log($"連線中：{path} /user:\"{user}\"");

            var processInfo = new ProcessStartInfo("cmd.exe", command)
            {
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };

            try
            {
                using (var process = Process.Start(processInfo))
                {
                    if (process == null)
                    {
                        Log($"無法啟動 net use 命令。");
                        return false;
                    }
                    string output = process.StandardOutput.ReadToEnd();
                    string error = process.StandardError.ReadToEnd();
                    process.WaitForExit();
                    if (process.ExitCode != 0)
                    {
                        Log($"無法連線到網路位置：{path}");
                        Log($"輸出：{output}");
                        Log($"錯誤：{error}");
                        Log(process.StandardError.ReadToEnd());
                        return false;
                    }
                }
            }
            catch (System.ComponentModel.Win32Exception ex)
            {
                Log($"執行 net use 命令時發生錯誤：{ex.Message}。請確認您的系統環境。");
                return false;
            }
            catch (Exception ex)
            {
                Log($"連線網路路徑時發生未預期的錯誤：{ex.Message}");
                return false;
            }

            Log($"成功連線到網路位置：{path}",true);
            return true;
        }
        static void DisconnectNetworkPath(string path)
        {
             // 判斷是否為本機路徑
            if (Path.IsPathRooted(path) && !path.StartsWith(@"\\")) // 檢查是否為磁碟根目錄且非 UNC 路徑
            {
                Log($"路徑 '{path}' 為本機路徑，不需斷開連線。");
                return; // 本機路徑不需斷開
            }
            string command = $"/C net use \"{path}\" /delete /y";
            var processInfo = new ProcessStartInfo("cmd.exe", command)
            {
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };

            using (var process = Process.Start(processInfo))
            {
                process.WaitForExit();
                if (process.ExitCode != 0)
                {
                    Log($"無法斷開網路位置：{path}");
                    Log(process.StandardError.ReadToEnd());
                }
                else
                {
                    Log($"已斷開網路位置：{path}");
                }
            }
        } 

        // 檢查目錄的寫入權限
        static bool HasWritePermission(string path)
        {
            try
            {
                var testFilePath = Path.Combine(path, "test.log");

                // 嘗試寫入檔案以檢查權限
                using (FileStream fs = File.Create(testFilePath)) { }
                
                // 刪除測試檔案
                File.Delete(testFilePath);

                // 如果成功則返回 true
                return true;
            }
            catch (UnauthorizedAccessException)
            {
                return false; // 無法存取
            }
            catch (Exception ex)
            {
                Log($"檢查寫入權限時出現錯誤：{ex.Message}");
                return false; // 其他錯誤
            }
        } 
        
        // 使用平行執行方式篩選檔案
        static IEnumerable<string> GetFilesRecursive(string path, string? searchPattern)
        {
            var files = new ConcurrentBag<string>(); // 用於安全地收集結果檔案
            try
            {
                // 使用 Stack 來管理目錄，避免直接遞迴造成的過深呼叫
                var directories = new Stack<string>();
                directories.Push(path);

                while (directories.Count > 0)
                {
                    string currentDir = directories.Pop();

                    try
                    {
                        // 將當前目錄的檔案加入結果集
                        foreach (var file in Directory.GetFiles(currentDir, searchPattern))
                        {
                            files.Add(file);
                        }

                        // 將子目錄加入堆疊以便後續處理
                        foreach (var subDir in Directory.GetDirectories(currentDir))
                        {
                            directories.Push(subDir);
                        }
                    }
                    catch (UnauthorizedAccessException ex)
                    {
                        Log($"無法存取目錄：{currentDir}，錯誤：{ex.Message}");
                    }
                    catch (Exception ex)
                    {
                        Log($"處理目錄 {currentDir} 時出現錯誤：{ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                Log($"處理目錄結構時出現錯誤：{ex.Message}");
            }

            return files;
        }


        static IEnumerable<string> GetFilesWithExclusions(string sourcePath, string searchPattern, IEnumerable<string> excludedPatterns)
        {
            var files = Directory.GetFiles(sourcePath, searchPattern, SearchOption.AllDirectories);
            return files.Where(f => !excludedPatterns.Any(pattern => MatchesPattern(f, pattern)));
        }

        // 判斷檔案名稱是否符合通配符規則
        static bool MatchesPattern(string filePath, string pattern)
        {
            string fileName = Path.GetFileName(filePath);
            if (pattern.Contains('*'))
            {
                string regexPattern = "^" + Regex.Escape(pattern).Replace("\\*", ".*") + "$";
                return Regex.IsMatch(fileName, regexPattern, RegexOptions.IgnoreCase);
            }
            return string.Equals(fileName, pattern, StringComparison.OrdinalIgnoreCase);
        } 
        
        // 取得本機 IP 位址的函式
        static List<string> GetLocalIPAddresses()
        {
            List<string> ipAddresses = new List<string>();
            try
            {
                var host = Dns.GetHostEntry(Dns.GetHostName());
                foreach (var ip in host.AddressList)
                {
                    if (ip.AddressFamily == AddressFamily.InterNetwork)
                    {
                        ipAddresses.Add(ip.ToString());
                    }
                } 
            }
            catch (SocketException ex)
            {
                Log($"網路錯誤：{ex.Message}。");
            }
            catch (Exception ex)
            {
                Log($"取得本機 IP 位址時發生錯誤：{ex.Message}。");         
            }
            if (ipAddresses.Count == 0)
            {
                Log("找不到任何 IPv4 位址。");
                ipAddresses.Add("127.0.0.1");
            }
            return ipAddresses;
        } 
    
        static bool CanAccessDirectory(string path)
        {
            try
            {
                // 使用 DirectoryInfo
                DirectoryInfo dirInfo = new DirectoryInfo(path);

                // 取得目錄的存取控制列表
                DirectorySecurity dirSecurity = dirInfo.GetAccessControl();

                // 如果沒有例外發生，表示可以存取
                return true;
            }
            catch (UnauthorizedAccessException ex)
            {
                Log($"沒有足夠的權限存取目錄：{path}，原因：{ex.Message}");
                return false;
            }
            catch (Exception ex)
            {
                Log($"檢查目錄存取權限時發生錯誤：{ex.Message}");
                return false;
            }
        } 
        
        static DateTime? ParseRelativeDate(string? dateString)
        {
            if (string.IsNullOrWhiteSpace(dateString))
                return null;

            // 用"+|-数字W|D|h|m|s"指定日期。（如-10D＝10天前）
            // Y|M|W|D|h|m|s分别代表：年、月周、日、小时、分、秒。
            if (dateString.Length > 1 && (dateString[0] == '-' || dateString[0] == '+'))
            {
                int value;
                char unit = dateString[^1];
                if (int.TryParse(dateString[1..^1], out value))
                {
                    return AdjustDate(value, unit, dateString[0] == '-'); // 判斷正負
                }
            }

            // 檢查是否為標準日期格式
            string[] formats = { "yyyyMMdd", "yyyy/MM/dd", "yyyy-MM-dd" };
            if (DateTime.TryParseExact(dateString, formats, null, System.Globalization.DateTimeStyles.None, out DateTime exactDate))
            {
                return exactDate; // 返回解析出的標準日期
            }

            // 嘗試作為常規日期字串解析
            if (DateTime.TryParse(dateString, out DateTime parsedDate))
            {
                return parsedDate;
            }

            return null; // 無法處理的格式
        }

        /// <summary>
        /// 根據值和單位調整日期
        /// </summary>
        static DateTime AdjustDate(int value, char unit, bool isPast)
        {
            switch (unit)
            {
                case 'Y':
                    return isPast ? DateTime.Now.AddYears(-value) : DateTime.Now.AddYears(value);
                case 'M':
                    return isPast ? DateTime.Now.AddMonths(-value) : DateTime.Now.AddMonths(value);
                case 'W':
                    return isPast ? DateTime.Now.AddDays(-value * 7) : DateTime.Now.AddDays(value * 7);
                case 'D':
                    return isPast ? DateTime.Now.AddDays(-value) : DateTime.Now.AddDays(value);
                case 'h':
                    return isPast ? DateTime.Now.AddHours(-value) : DateTime.Now.AddHours(value);
                case 'm':
                    return isPast ? DateTime.Now.AddMinutes(-value) : DateTime.Now.AddMinutes(value);
                case 's':
                    return isPast ? DateTime.Now.AddSeconds(-value) : DateTime.Now.AddSeconds(value);
                default:
                    throw new ArgumentException("無效的時間單位。");
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