# �Ǘ��Ҍ����`�F�b�N
if (-not ([Security.Principal.WindowsPrincipal] `
    [Security.Principal.WindowsIdentity]::GetCurrent()
    ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Error "���̃X�N���v�g�͊Ǘ��҂Ƃ��Ď��s���Ă��������B"
    exit
}

# �p�X�ݒ�
$programFilesPath = Join-Path $env:ProgramFiles "OnenoteAddin"
$programDataPath  = Join-Path $env:ProgramData  "OnenoteAddin"

# Regasm���s
$dllFullPath = Join-Path $programFilesPath "OnenoteAddin.dll"
$version = "v4.0.30319" # !!FIXME!!
$frameworkPath = "${env:windir}\Microsoft.NET\Framework64\${version}\regasm.exe"

if (-Not (Test-Path $frameworkPath)) {
    Write-Error "regasm.exe ��������܂���: $frameworkPath"
    exit
}

& $frameworkPath $dllFullPath /unregister

if ($LASTEXITCODE -eq 0) {
    Write-Host "regasm �o�^�������������܂����B"
} else {
    Write-Error "regasm �o�^�����Ɏ��s���܂����B"
}

# �f�B���N�g���폜
Remove-Item -Path $programFilesPath -Recurse -Confirm
Remove-Item -Path $programDataPath -Recurse -Confirm

