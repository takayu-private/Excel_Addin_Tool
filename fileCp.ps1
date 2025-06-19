# 現在のディレクトリ
$currentPath = Get-Location

# アドインファイル名
$SearchAddinFile = "SearchExtensionAddin.xlam"
$FunctionAddinFile = "FunctionExtensionAddin.xlam"
$addinFileList = $SearchAddinFile, $FunctionAddinFile

# ユーザ名を取得
$userProfile = [Environment]::GetEnvironmentVariable("USERPROFILE")

# コピー先のパス
$destinationPath = "$userProfile\AppData\Roaming\Microsoft\AddIns\"

foreach ($filePath in $addinFileList) {

  # アドインファイルのパス
  $addinPath = "$currentPath\$filePath"

  # ファイルをコピー
  Copy-Item -Path $addinPath -Destination $destinationPath -Force

  if (-not (Test-Path $addinPath)) {
    New-Item -Type File "$currentPath\copy_ng"
  }
}

# ショートカットを作成
$WshShell = New-Object -ComObject WScript.Shell
$ShortcutPath = "$currentPath\AddIns.lnk"
$Shortcut = $WshShell.CreateShortcut($ShortcutPath)
$Shortcut.TargetPath = $destinationPath
$Shortcut.Save()

# COM オブジェクトの解放
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($WshShell) | Out-Null
