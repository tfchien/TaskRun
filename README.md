# TaskMng
TaskMng 是一個用 C# 編寫的高級任務管理應用程式，專為多台主機的自動化備份工作而設計。此專案提供了靈活的配置選項、多種備份模式和增強的安全性功能。

## 主要特點

- **Excel 配置管理**：通過 Excel 文件輕鬆設定和管理多台主機的備份任務。
- **多種備份模式**：支援 File Copy、Robocopy、XCopy、FastCopy 和 Delete 模式。
- **靈活的參數設定**：直接在配置文件中設定複雜的參數，無需額外的批次檔。
- **安全的憑證管理**：使用加密方式將連線帳號安全存儲在 Windows 註冊表中。
- **智能執行**：根據配置檢查本機 IP，只執行符合條件的任務。
- **Windows 排程友好**：設計適合直接在 Windows 任務排程器中執行。

## 專案結構

TaskRun.sln 
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
    git clone https://github.com/tfchien/TaskRun.git
    cd TaskRun/taskrun
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

建立執行檔
使用以下命令單一執行檔：

```sh
dotnet publish -c release -r win-x64 -p:PublishSingleFile=true -p:PublishTrimmed=true
```
### 複製執行到到執行目錄

假設你執行資料夾在d:\backup\taskrun，

```sh
xcopy taskrun\bin\release\net8.0\win-x64\publish\taskrun.exe d:\backup\taskrun /E /H /Y
```

## 使用說明

在 Excel 配置文件中設定備份任務和參數。
使用 EncryptMng 類別將連線憑證加密存儲在 Windows 註冊表中。
將應用程式添加到 Windows 任務排程器中。
TaskRun 將根據配置自動執行符合條件的備份任務。

## 配置管理

 1. 配置管理由 ConfigMng 類別處理。您可以在 Excel 文件中定義以下內容：
 
 2. 備份來源和目標
 
 3. 備份模式（File Copy、Robocopy、XCopy、FastCopy、Delete）
 
 4. 特定的備份參數
 
 5. 執行條件（如 IP 限制 

 6. 設定檔格式
    - file: backup_config.xlsx
    - worksheet: BackupConfig
    - column header:
      
    SourcePaths	SourceFolders	SourceUser	TargetPath	TargetFolder	TargetUser	HostIp	FormDate	ToDate	IncludeFilter	ExcludeFilter	Overwrite	BufSize	Acl	Tool

    ![image](https://github.com/user-attachments/assets/4e1058d3-4fe4-4d71-8272-21f105a6d7ec)


    - file: users.xlsx
    - worksheet: users
    - column header:
      
    user	password	encrypted

    ![image](https://github.com/user-attachments/assets/27d86c53-9588-445f-b17a-d78edc2fbcbe)

## 日誌記錄
 - 日誌功能由 LogMng 類別管理。
 - 同時產生console log及文字記錄
 - 文字記錄檔可在ini檔設置

  ini檔內容如下:
  
  [Paths]    
    BaseDirectory=  
    ConfigFile=backup_config.xlsx  
    UserFile=backup_user.xlsx

  [Windows]    
    isLogConfig=false    
    force_close=false


##貢獻

  歡迎貢獻！請 fork 此儲存庫並創建一個包含您更改的 pull request。

##授權

  本專案採用 MIT 授權 - 詳情請參閱 LICENSE 文件。
