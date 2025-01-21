# TaskMng

TaskMng 是一個用 C# 編寫的高級任務管理應用程式，專為多台主機的自動化備份工作而設計。此專案提供了靈活的配置選項、多種備份模式和增強的安全性功能。

## 主要特點

- **Excel 配置管理**：通過 Excel 文件輕鬆設定和管理多台主機的備份任務。
- **多種備份模式**：支援 File Copy、Robocopy、XCopy 和 FastCopy 模式。
- **靈活的參數設定**：直接在配置文件中設定複雜的參數，無需額外的批次檔。
- **安全的憑證管理**：使用加密方式將連線帳號安全存儲在 Windows 註冊表中。
- **智能執行**：根據配置檢查本機 IP，只執行符合條件的任務。
- **Windows 排程友好**：設計適合直接在 Windows 任務排程器中執行。

## 專案結構

TaskMng.sln 
taskrun/ 
- BackupProgram.cs 
- BackupTask.cs 
- BackupUser.cs 
- ConfigMng.cs 
- CopyMng.cs 
- EncryptMng.cs 
- LogMng.cs 
- taskrun.csproj

## 開始使用

### 先決條件

- .NET 8.0 SDK 或更高版本
- Visual Studio 或其他 C# IDE

### 安裝

1. 克隆儲存庫：
    ```sh
    git clone https://github.com/tfchien/TaskMng.git
    cd TaskMng/taskrun
    ```

2. 還原相依項：
    ```sh
    dotnet restore
    ```

### 建置專案

在 taskrun 目錄中執行以下命令：
```sh
dotnet build
```


非常感謝您提供這些重要資訊。我會根據您的描述重新編寫 README 文件，以更好地反映專案的核心功能和優勢。以下是修改後的 README 內容：

markdown
Copy
# TaskMng

TaskMng 是一個用 C# 編寫的高級任務管理應用程式，專為多台主機的自動化備份工作而設計。此專案提供了靈活的配置選項、多種備份模式和增強的安全性功能。

## 主要特點

- **Excel 配置管理**：通過 Excel 文件輕鬆設定和管理多台主機的備份任務。
- **多種備份模式**：支援 File Copy、Robocopy、XCopy 和 FastCopy 模式。
- **靈活的參數設定**：直接在配置文件中設定複雜的參數，無需額外的批次檔。
- **安全的憑證管理**：使用加密方式將連線帳號安全存儲在 Windows 註冊表中。
- **智能執行**：根據配置檢查本機 IP，只執行符合條件的任務。
- **Windows 排程友好**：設計適合直接在 Windows 任務排程器中執行。

## 專案結構

TaskMng.sln 
taskrun/ 
- BackupProgram.cs 
- BackupTask.cs 
- BackupUser.cs 
- ConfigMng.cs 
- CopyMng.cs 
- EncryptMng.cs 
- LogMng.cs 
- taskrun.csproj

## 開始使用

### 先決條件

- .NET 8.0 SDK 或更高版本
- Visual Studio 或其他 C# IDE

### 安裝

1. 克隆儲存庫：
    ```sh
    git clone https://github.com/tfchien/TaskMng.git
    cd TaskMng/taskrun
    ```

2. 還原相依項：
    ```sh
    dotnet restore
    ```

### 建置專案

在 taskrun 目錄中執行以下命令：
```sh
dotnet build
```

執行專案
使用以下命令運行專案：

```sh
dotnet run
```

##使用說明
在 Excel 配置文件中設定備份任務和參數。
使用 EncryptMng 類別將連線憑證加密存儲在 Windows 註冊表中。
將應用程式添加到 Windows 任務排程器中。
TaskRun 將根據配置自動執行符合條件的備份任務。

##配置管理

 1.配置管理由 ConfigMng 類別處理。您可以在 Excel 文件中定義以下內容：
 
 2.備份來源和目標
 
 3.備份模式（File Copy、Robocopy、XCopy、FastCopy）
 
 4.特定的備份參數
 
 5.執行條件（如 IP 限制 

 6. 設定檔格式
    - file: backup_config.xlsx
    - worksheet: BackupConfig
    - column header:
      
    SourcePaths	SourceFolders	SourceUser	TargetPath	TargetFolder	TargetUser	HostIp	FormDate	ToDate	IncludeFilter	ExcludeFilter	Overwrite	BufSize	Acl	Tool

     ![image](https://github.com/user-attachments/assets/09b1f741-9a38-44b6-8f20-ce88974a5b73)

    - file: users.xlsx
    - worksheet: users
    - column header:
      
    user	password	encrypted
    
    ![image](https://github.com/user-attachments/assets/9f6334b0-4d0a-47e0-95cd-b2f94a46ceb1)



##日誌記錄
 - 日誌功能由 LogMng 類別管理。
 - 同時產生console log及文字記錄
 - 文字記錄檔可在ini檔設置


##貢獻

  歡迎貢獻！請 fork 此儲存庫並創建一個包含您更改的 pull request。

##授權

  本專案採用 MIT 授權 - 詳情請參閱 LICENSE 文件。
