---
name: gas-sheet-deploy-powershell
description: Windows PowerShell version of the GAS deployment workflow. Use when deploying Google Apps Script projects to Google Sheets on Windows. Covers Node.js, clasp setup, and batch deployment using native PowerShell commands (Invoke-RestMethod).
---

# Google Sheet + Apps Script 部署流程 (Windows PowerShell 版)

這份文件提供專為 Windows 使用者設計的 PowerShell 部署腳本，用於將 Google Apps Script 專案部署到 Google Sheets。

## clasp 首次安裝流程 (Windows)

### 1. 確認 Node.js 版本 (需 v22+)

在 PowerShell 執行：
```powershell
node -v
```
若未安裝或版本過低，請前往 [nodejs.org](https://nodejs.org/) 下載最新的 LTS 版本安裝。

### 2. 安裝 clasp

```powershell
npm install -g @google/clasp
clasp --version   # 確認安裝成功
```

### 3. 啟用 Google Apps Script API

前往 [https://script.google.com/home/usersettings](https://script.google.com/home/usersettings)，將「Google Apps Script API」切換為 **ON**。

### 4. 登入 Google 帳號

```powershell
clasp login
```
瀏覽器會自動開啟，完成授權後，登入資訊會存在 `$HOME\.clasprc.json`。

### 5. 確認登入狀態 (PowerShell 版)

```powershell
$claspRc = Get-Content "$HOME\.clasprc.json" | ConvertFrom-Json
Write-Host "clasp 登入狀態: $($claspRc.tokens.default.access_token -ne $null)"
```

---

## 前置需求

執行腳本前，請確保：
1. 已透過 `clasp login` 登入。
2. **權限注意**：若需要匯入 CSV 到試算表，建議使用自訂 OAuth Creds 登入（需含 `spreadsheets` 範圍）。

---

## 單一資料夾部署 (PowerShell)

請將以下程式碼存為 `deploy_one.ps1` 或直接貼上執行：

```powershell
# 設定
$FOLDER = "C:\你的路徑\Session0X_主題"
$NAME = Split-Path $FOLDER -Leaf

# 取得 Access Token
$claspRc = Get-Content "$HOME\.clasprc.json" | ConvertFrom-Json
$TOKEN = $claspRc.tokens.default.access_token

$headers = @{
    "Authorization" = "Bearer $TOKEN"
    "Content-Type"  = "application/json"
}

# 1. 建立試算表
$body = @{
    name     = $NAME
    mimeType = "application/vnd.google-apps.spreadsheet"
} | ConvertTo-Json
$sheet = Invoke-RestMethod -Uri "https://www.googleapis.com/drive/v3/files" -Method Post -Headers $headers -Body $body
$SHEET_ID = $sheet.id
Write-Host "試算表建立成功: $SHEET_ID"

# 2. 搜尋同名雲端資料夾並移動 (選用)
$query = "name = '$NAME' and mimeType = 'application/vnd.google-apps.folder' and trashed = false"
$searchUri = "https://www.googleapis.com/drive/v3/files?q=$([Uri]::EscapeDataString($query))&pageSize=1"
try {
    $searchResult = Invoke-RestMethod -Uri $searchUri -Method Get -Headers $headers
    if ($searchResult.files.Count -gt 0) {
        $DRIVE_FOLDER_ID = $searchResult.files[0].id
        Invoke-RestMethod -Uri "https://www.googleapis.com/drive/v3/files/$SHEET_ID?addParents=$DRIVE_FOLDER_ID" -Method Patch -Headers $headers
        Write-Host "已移至 Drive 資料夾 ✅"
    }
} catch {
    Write-Warning "查無對應資料夾或移動失敗"
}

# 3. 建立 Apps Script 並綁定
$scriptBody = @{
    title    = "${NAME}_script"
    parentId = $SHEET_ID
} | ConvertTo-Json
$script = Invoke-RestMethod -Uri "https://script.googleapis.com/v1/projects" -Method Post -Headers $headers -Body $scriptBody
$SCRIPT_ID = $script.scriptId
Write-Host "Script 建立成功: $SCRIPT_ID"

# 4. 準備本地檔案
if (-not (Test-Path "$FOLDER\appsscript.json")) {
    $manifest = @{
        timeZone         = "Asia/Taipei"
        dependencies     = @{}
        exceptionLogging = "STACKDRIVER"
        runtimeVersion   = "V8"
    } | ConvertTo-Json
    $manifest | Set-Content "$FOLDER\appsscript.json"
}

$claspJson = @{
    scriptId = $SCRIPT_ID
    rootDir  = "."
} | ConvertTo-Json
$claspJson | Set-Content "$FOLDER\.clasp.json"

# 5. 推送程式碼
Set-Location $FOLDER
clasp push --force

# 6. 匯入 CSV (若有)
if (Test-Path "$FOLDER\sample_data.csv") {
    # 簡易 CSV 讀取 (不包含複雜的引號處理)
    $lines = Get-Content "$FOLDER\sample_data.csv"
    $values = @()
    foreach ($line in $lines) {
        $values += ,($line.Split(","))
    }
    $payload = @{ values = $values } | ConvertTo-Json -Depth 10
    
    try {
        Invoke-RestMethod -Uri "https://sheets.googleapis.com/v4/spreadsheets/$SHEET_ID/values/Sheet1!A1?valueInputOption=USER_ENTERED" `
                          -Method Put -Headers $headers -Body $payload
        Write-Host "CSV 資料匯入 ✅"
    } catch {
        Write-Error "CSV 匯入失敗，請確認是否具備 spreadsheets 權限"
    }
}

Write-Host "--------------------------------"
Write-Host "✅ 部署完成！"
Write-Host "試算表: https://docs.google.com/spreadsheets/d/$SHEET_ID"
Write-Host "Script: https://script.google.com/d/$SCRIPT_ID/edit"
```

---

## 批次部署 (所有未部署的 Session)

```powershell
$BASE = "C:\你的路徑\google-app-script-course"
$claspRc = Get-Content "$HOME\.clasprc.json" | ConvertFrom-Json
$TOKEN = $claspRc.tokens.default.access_token
$headers = @{ "Authorization" = "Bearer $TOKEN"; "Content-Type" = "application/json" }

# 找出所有含 code.gs 的資料夾
$gsFiles = Get-ChildItem -Path $BASE -Filter "code.gs" -Recurse

foreach ($gs in $gsFiles) {
    $folder = $gs.DirectoryName
    $name = $gs.Directory.Name
    
    # 檢查是否已部署
    if (Test-Path "$folder\.clasp.json") {
        Write-Host "⏭ 跳過已部署: $name" -ForegroundColor Gray
        continue
    }

    Write-Host "▶ 正在部署: $name ..." -ForegroundColor Cyan

    # 1. 建立試算表
    $sheet = Invoke-RestMethod -Uri "https://www.googleapis.com/drive/v3/files" -Method Post -Headers $headers -Body (@{name=$name; mimeType="application/vnd.google-apps.spreadsheet"} | ConvertTo-Json)
    $sheetId = $sheet.id

    # 2. 建立 Script
    $script = Invoke-RestMethod -Uri "https://script.googleapis.com/v1/projects" -Method Post -Headers $headers -Body (@{title="${name}_script"; parentId=$sheetId} | ConvertTo-Json)
    $scriptId = $script.scriptId

    # 3. 本地設定
    if (-not (Test-Path "$folder\appsscript.json")) {
        '{"timeZone":"Asia/Taipei","dependencies":{},"exceptionLogging":"STACKDRIVER","runtimeVersion":"V8"}' | Set-Content "$folder\appsscript.json"
    }
    (@{scriptId=$scriptId; rootDir="."} | ConvertTo-Json) | Set-Content "$folder\.clasp.json"

    # 4. Push
    Push-Location $folder
    clasp push --force
    Pop-Location

    Write-Host "  ✅ 完成: $sheetId" -ForegroundColor Green
}
```

---

## 驗證部署狀態 (PowerShell)

```powershell
$BASE = "."
$total = (Get-ChildItem -Path $BASE -Filter "code.gs" -Recurse).Count
$deployed = (Get-ChildItem -Path $BASE -Filter ".clasp.json" -Recurse).Count
Write-Host "部署進度：$deployed / $total"
```

---

## 常見問題 (Windows 特有)

1. **數位簽署錯誤**：若執行 `.ps1` 報錯「此系統上已停用指令碼執行」，請以管理員身分開啟 PowerShell 並執行：
   ```powershell
   Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
   ```
2. **路徑包含空白**：路徑若有空白，請務必使用引號包起來，例如 `"$FOLDER\code.gs"`。
3. **編碼問題**：若 CSV 匯入出現亂碼，請確認檔案為 `UTF-8` 編碼。
