# 現在のディレクトリ
$currentPath = Get-Location

# アドインファイル名
$addinFileName = "SearchExtensionAddin.xlam"

# 元のアドインファイルのパス
$sourceAddinPath = "$currentPath\SearchExtensionAddin.xlam"

# ユーザ名を取得
$userProfile = [Environment]::GetEnvironmentVariable("USERPROFILE")

# コピー先のパス
$destinationPath = "$userProfile\AppData\Roaming\Microsoft\AddIns\"

# ファイルをコピー
Copy-Item -Path $sourceAddinPath -Destination $destinationPath -Force

if (-not (Test-Path $sourceAddinPath)) {
  New-Item -Type File "$currentPath\copy_ng"
}

# ショートカットを作成
$WshShell = New-Object -ComObject WScript.Shell
$ShortcutPath = ".\AddIns.lnk"
$Shortcut = $WshShell.CreateShortcut($ShortcutPath)
$Shortcut.TargetPath = $destinationPath
$Shortcut.Save()
# COM オブジェクトの解放
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($WshShell) | Out-Null
