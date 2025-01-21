using System;
using System.IO;
using System.Security.Cryptography;
using OfficeOpenXml;
using Microsoft.Win32;

namespace TaskRun
{
    public class EncryptMng
    {
        static string registryRoot = @"HKEY_CURRENT_USER\SOFTWARE\NonSheng";
        static string registryPath = @"HKEY_CURRENT_USER\SOFTWARE\NonSheng\TaskRun"; 
        static string registryUsers = @"HKEY_CURRENT_USER\SOFTWARE\NonSheng\TaskRun\users"; 
        static string keyName = "EncryptionKey";
        static string ivName = "InitializationVector";
        public EncryptMng()
        {
            // 確保日誌目錄存在
           LogMng.Initialize(BackupProgram.LogDirectory);
        }

        /// <summary>
        /// 加密指定 Excel 檔案中的密碼。
        /// </summary>
        public void EncryptExcelPasswords(string excelFilePath)
        {
            try{
                // 確保註冊表鍵存在
                EnsureRegistryKeyExists(registryRoot);
                EnsureRegistryKeyExists(registryPath);
                EnsureRegistryKeyExists(registryUsers);               
                var users = LoadUsersFromExcel(excelFilePath);            
                var keys = LoadOrCreateKeys(excelFilePath);
                foreach (var user in users)
                {
                    bool isUerExist = string.IsNullOrEmpty( (string)Registry.GetValue(registryUsers, user.User, null));
                    if ((isUerExist || !user.Encrypted) && !string.IsNullOrEmpty(user.Password))
                    {
                        string encryptedPassword = EncryptString(user.Password, keys.Key, keys.IV);
                        //user.encryptedPassword = encryptedPassword;
                        user.Encrypted = true;
                        Registry.SetValue(registryUsers, user.User , encryptedPassword, RegistryValueKind.String);
                        Log($"用戶 {user.User} 的密碼已加密。");
                    } else {
                        Log($"用戶 {user.User} 的密碼已存在加密。");
                    }
                }
                // 寫回更新的用戶資料到 Excel
                SaveUsersToExcel(excelFilePath, users);         
                Log($"已加密 Excel 檔案中的密碼：{excelFilePath}");
            } catch (Exception ex)
            {
                Log($"加密 Excel 檔案時發生錯誤：{ex.Message} :{ex.StackTrace}");
                throw;
            }
          
            
        }

        
        private (byte[] Key, byte[] IV) LoadOrCreateKeys(string excelFilePath)
        {
            byte[] key;
            byte[] iv;

             // 嘗試從 Excel 檔案中載入金鑰
            var existingKeys = LoadKeysFromExcel(excelFilePath); // 方法：自定義來從 Excel 中讀取金鑰和 IV

            if (existingKeys.Key != null && existingKeys.IV != null)
            {
                // 如果 Excel 中有金鑰，則使用這些金鑰
                key = existingKeys.Key;
                iv = existingKeys.IV;

                // 更新註冊表中的金鑰
                Registry.SetValue(registryPath, keyName, key, RegistryValueKind.Binary);
                Registry.SetValue(registryPath, ivName, iv, RegistryValueKind.Binary);

                Log($"從 Excel 檔案載入金鑰，並更新註冊表");
            }
            else  
            {
                // 如果註冊表中沒有金鑰，則生成新的金鑰
                var keys =  GenerateNewKeys();
                key = keys.Key;
                iv = keys.IV;    
                SaveKeysToExcel(excelFilePath, keys.Key, keys.IV);
                Log($"生成新的加密金鑰與 IV，儲存於註冊表與 user.xlsx");
            }
 
            return (key, iv);
        }
        public List<BackupUser> LoadUsersFromExcel(string excelFilePath)
        {
            var users = new List<BackupUser>();
            using (ExcelPackage package = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets["users"];

                // 確保包含正確的欄位
                int userColumn = FindColumnIndex(worksheet, "user");
                int passwordColumn = FindColumnIndex(worksheet, "password");
                int encryptedColumn = FindColumnIndex(worksheet, "encrypted");

                if (userColumn == -1 || passwordColumn == -1 || encryptedColumn == -1)
                {
                    throw new Exception("Excel 檔案缺少必要的欄位 (user, password, encrypted)。");
                }

                for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                {
                    var user = worksheet.Cells[row, userColumn].Value?.ToString();
                    var password = worksheet.Cells[row, passwordColumn].Value?.ToString();
                    var isEncrypted = bool.TryParse(worksheet.Cells[row, encryptedColumn].Value?.ToString(), out bool encrypted) ? encrypted : false;

                    if (user != null)
                    {
                        users.Add(new BackupUser(user, password, isEncrypted));
                    }
                }
            }
            return users;
        }

        public void SaveUsersToExcel(string excelFilePath, List<BackupUser> users)
        {
            using (ExcelPackage package = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets["users"];
                 // 確保包含正確的欄位
                int userColumn = FindColumnIndex(worksheet, "user");
                int passwordColumn = FindColumnIndex(worksheet, "password");
                int encryptedColumn = FindColumnIndex(worksheet, "encrypted");

                for (int row = 0; row < users.Count; row++)
                {
                    worksheet.Cells[row + 2, userColumn].Value = users[row].User;
                    worksheet.Cells[row + 2, passwordColumn].Value = users[row].Password;
                    worksheet.Cells[row + 2, encryptedColumn].Value = users[row].Encrypted.ToString();
                }
                package.Save();
            }
        }


        private (byte[] Key, byte[] IV) LoadKeysFromExcel(string excelFilePath)
        {
            // 初始為空的金鑰和 IV
            byte[] key = null;
            byte[] iv = null;

            using (ExcelPackage package = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets["Keys"]; // 假設用戶在 Keys 工作表中存儲金鑰
                
                if (worksheet != null)
                {
                     // 確保包含正確的欄位
                    int keyColumn = FindColumnIndex(worksheet, "Key");
                    int ivColumn = FindColumnIndex(worksheet, "IV");
                    // 讀取金鑰和 IV 的 Base64 字串並轉換成 byte[]
                    var keyBase64 = worksheet.Cells[2, keyColumn].Value?.ToString(); // 假設金鑰在第2行第1列
                    var ivBase64 = worksheet.Cells[2, ivColumn].Value?.ToString(); // 假設 IV 在第2行第2列

                    if (!string.IsNullOrEmpty(keyBase64))
                        key = Convert.FromBase64String(keyBase64);
                    else 
                        throw new Exception("找不到金鑰");

                    if (!string.IsNullOrEmpty(ivBase64))
                        iv = Convert.FromBase64String(ivBase64);
                    else
                        throw new Exception("找不到 IV");
                }
            }

            return (key, iv);
        }

        private void SaveKeysToExcel(string excelFilePath, byte[] key, byte[] iv)
        {
            using (ExcelPackage package = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                // 檢查是否已存在名為 "Keys" 的工作表，若無則創建一個新的
                ExcelWorksheet worksheet = package.Workbook.Worksheets["Keys"];
                if (worksheet == null)
                {
                    worksheet = package.Workbook.Worksheets.Add("Keys");
                    // 設定標題
                    worksheet.Cells[1, 1].Value = "Key";
                    worksheet.Cells[1, 2].Value = "IV";
                }

                // 將金鑰和 IV 寫入
                worksheet.Cells[2, 1].Value = Convert.ToBase64String(key);
                worksheet.Cells[2, 2].Value = Convert.ToBase64String(iv);

                // 保存更改
                package.Save();
            }
        }


              /// <summary>
        /// 找到指定列名稱在 Excel 工作表中的索引。
        /// </summary>
        private static int FindColumnIndex(ExcelWorksheet worksheet, string columnName)
        {
            // 將目標欄位名稱轉換為小寫
            string targetColumnName = columnName.ToLower();

            for (int col = 1; col <= worksheet.Dimension.Columns; col++)
            {
                // 將當前欄位名稱轉換為小寫並進行比較
                if (worksheet.Cells[1, col].Value?.ToString().Trim().ToLower() == targetColumnName)
                    return col; // 返回找到的索引
            }
            return -1; // 如果找不到，返回 -1
        }
 

        /// <summary>
        /// 使用 AES 加密提供的純文字。
        /// </summary>
        public static string EncryptString(string plainText, byte[] key, byte[] iv)
        {
            try
            {
                using (Aes aesAlg = Aes.Create())
                {
                    aesAlg.Key = key;
                    aesAlg.IV = iv;

                    ICryptoTransform encryptor = aesAlg.CreateEncryptor(aesAlg.Key, aesAlg.IV);

                    using (var msEncrypt = new MemoryStream())
                    using (var csEncrypt = new CryptoStream(msEncrypt, encryptor, CryptoStreamMode.Write))
                    using (var swEncrypt = new StreamWriter(csEncrypt))
                    {
                        swEncrypt.Write(plainText);
                        swEncrypt.Close();
                        Log($"生成加密文本。");
                        return Convert.ToBase64String(msEncrypt.ToArray());
                    }
                }
            }
            catch (Exception ex)
            {
                Log($"加密文本(EncryptString)時發生錯誤：{ex.Message} :{ex.StackTrace}");
                throw;
            }
            
        }
        public static string DecryptString(string cipherText, byte[] key, byte[] iv)
        {
            try
            {
                using (Aes aesAlg = Aes.Create())
                {
                    aesAlg.Key = key;
                    aesAlg.IV = iv;

                    ICryptoTransform decryptor = aesAlg.CreateDecryptor(aesAlg.Key, aesAlg.IV);

                    using (var msDecrypt = new MemoryStream(Convert.FromBase64String(cipherText)))
                    using (var csDecrypt = new CryptoStream(msDecrypt, decryptor, CryptoStreamMode.Read))
                    using (var srDecrypt = new StreamReader(csDecrypt))
                    {
                        return srDecrypt.ReadToEnd();
                    }
                }
            }
            catch (Exception ex)
            {
                Log($"解密文本(DecryptString)時發生錯誤：{ex.Message} :{ex.StackTrace}");
                throw;
            }           
        }

        /// <summary>
        /// 生成新金鑰和 IV，並返回它們
        /// </summary>
        private (byte[] Key, byte[] IV) GenerateNewKeys()
        {
            using (Aes aesAlg = Aes.Create())
            {
                byte[] key = aesAlg.Key;
                byte[] iv = aesAlg.IV;

                // 儲存新的金鑰和 IV 到註冊表
                Registry.SetValue(registryPath, keyName, key, RegistryValueKind.Binary);
                Registry.SetValue(registryPath, ivName, iv, RegistryValueKind.Binary);

                Log($"生成新的加密金鑰與 IV，並儲存於註冊表: {keyName}, {ivName}");
                return (key, iv);
            }
        }

        public static string GetDecryptedPassword(string username)
        {            
            string encryptedPassword = (string)Registry.GetValue(registryUsers, username, null);
            if (string.IsNullOrEmpty(encryptedPassword))
            {
                Log($"找不到用戶 {username} 的加密密碼 ({registryUsers}\\{username})");
                throw new Exception($"找不到用戶 {username} 的加密密碼");
            }
            byte[] key = (byte[])Registry.GetValue(registryPath, keyName, null);
            byte[] iv = (byte[])Registry.GetValue(registryPath, ivName, null);
            if(key == null || iv == null)
            {
                throw new Exception("找不到加密金鑰或 IV");
            }
            // 解密密碼
            return DecryptString(encryptedPassword, key, iv);
        }


        private bool EnsureRegistryKeyExists(string registryPath)
        {
            // 開始從根鍵 HKEY_CURRENT_USER 開始
            using (RegistryKey CurrentUser = Registry.CurrentUser)
            {
                // 使用 RegistryKey.OpenSubKey 方法來檢查指定路徑的註冊表鍵是否存在
                using (RegistryKey subKey = CurrentUser.OpenSubKey(registryPath))
                {
                    if (subKey == null)
                    {
                        // 如果註冊表鍵不存在，則創建它
                        using (RegistryKey newSubKey = CurrentUser.CreateSubKey(registryPath))
                        {
                            if (newSubKey == null)
                            {
                                Log($"無法確保註冊表鍵 {registryPath} 的存在。");
                                throw new Exception($"無法確保註冊表鍵 {registryPath} 的存在。");
                            }
                        }
                        return false;
                    }
                }
            }
            return true;
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