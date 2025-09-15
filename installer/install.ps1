# 管理者権限チェック
if (-not ([Security.Principal.WindowsPrincipal] `
    [Security.Principal.WindowsIdentity]::GetCurrent()
    ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Error "このスクリプトは管理者として実行してください。"
    exit
}

# パス設定
$programFilesPath = Join-Path $env:ProgramFiles "OnenoteAddin"
$programDataPath  = Join-Path $env:ProgramData  "OnenoteAddin"

# ディレクトリ作成
New-Item -Path $programFilesPath -ItemType Directory -Force | Out-Null
New-Item -Path $programDataPath  -ItemType Directory -Force | Out-Null

# ファイルのコピー
Copy-Item -Path "${PSScriptRoot}\ProgramFiles\*" -Destination $programFilesPath -Force
Copy-Item -Path "${PSScriptRoot}\ProgramData\*" -Destination $programDataPath -Force

# Regasm実行
$dllFullPath = Join-Path $programFilesPath "OnenoteAddin.dll"
$version = "v4.0.30319" # !!FIXME!!
$frameworkPath = "${env:windir}\Microsoft.NET\Framework64\${version}\regasm.exe"

if (-Not (Test-Path $frameworkPath)) {
    Write-Error "regasm.exe が見つかりません: $frameworkPath"
    exit
}

& $frameworkPath $dllFullPath /codebase

if ($LASTEXITCODE -eq 0) {
    Write-Host "regasm 登録が完了しました。"
} else {
    Write-Error "regasm 登録に失敗しました。"
}
