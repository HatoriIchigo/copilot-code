<#
.SYNOPSIS
    copilot-code をインストールし、環境変数を登録する。

.DESCRIPTION
    copilot.ps1 / copilot.xlsx を InstallDir（既定 ~/.copilot-code）へコピーし、
    起動用の copilot-code.cmd を作る。さらにユーザー環境変数を登録する。
      COPILOT_CODE_HOME : InstallDir
      Path              : InstallDir を追加（どのディレクトリからでも copilot-code で起動できる）
    何度実行しても同じ結果になる。管理者権限は不要。

.EXAMPLE
    powershell -File ./setup.ps1
#>
param(
    [string]$InstallDir = (Join-Path $env:USERPROFILE '.copilot-code'),
    [switch]$Uninstall
)

$ErrorActionPreference = 'Stop'
$InstallDir = [IO.Path]::GetFullPath($InstallDir)

function Get-UserPath {
    $p = [Environment]::GetEnvironmentVariable('Path', 'User')
    if ($null -eq $p) { '' } else { $p }
}
function Split-PathList([string]$p) {
    @($p -split ';' | Where-Object { $_ -ne '' })
}
function Test-SameDir([string]$a, [string]$b) {
    $a.TrimEnd('\') -ieq $b.TrimEnd('\')
}

if ($Uninstall) {
    $kept = Split-PathList (Get-UserPath) | Where-Object { -not (Test-SameDir $_ $InstallDir) }
    [Environment]::SetEnvironmentVariable('Path', ($kept -join ';'), 'User')
    [Environment]::SetEnvironmentVariable('COPILOT_CODE_HOME', $null, 'User')
    if (Test-Path -LiteralPath $InstallDir) { Remove-Item -LiteralPath $InstallDir -Recurse -Force }
    Write-Host "削除しました: $InstallDir（作業用コピーの %LOCALAPPDATA%\copilot-code は残しています）"
    return
}

# --- 1. ファイルをコピー ---
foreach ($f in 'copilot.ps1', 'copilot.xlsx') {
    if (-not (Test-Path -LiteralPath (Join-Path $PSScriptRoot $f))) {
        throw "$f が見つかりません（setup.ps1 と同じフォルダに置いてください）"
    }
}
New-Item -ItemType Directory -Force -Path $InstallDir | Out-Null

if ((Resolve-Path -LiteralPath $PSScriptRoot).Path.TrimEnd('\') -ine $InstallDir.TrimEnd('\')) {
    Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'copilot.ps1') -Destination $InstallDir -Force

    # マスタのブックは、プロンプトを編集して使うことがある。内容が違えば退避してから上書きする
    $src = Join-Path $PSScriptRoot 'copilot.xlsx'
    $dst = Join-Path $InstallDir 'copilot.xlsx'
    if ((Test-Path -LiteralPath $dst) -and
        ((Get-FileHash -LiteralPath $src).Hash -ne (Get-FileHash -LiteralPath $dst).Hash)) {
        $bak = Join-Path $InstallDir 'copilot.xlsx.bak'
        Copy-Item -LiteralPath $dst -Destination $bak -Force
        Write-Host "既存のブックを退避しました: $bak"
    }
    Copy-Item -LiteralPath $src -Destination $dst -Force
}

# --- 2. 起動用コマンド ---
$cmd = "@echo off`r`npowershell -NoProfile -ExecutionPolicy Bypass -File `"%~dp0copilot.ps1`" %*`r`n"
[IO.File]::WriteAllText((Join-Path $InstallDir 'copilot-code.cmd'), $cmd, [Text.Encoding]::ASCII)

# --- 3. 環境変数（ユーザー） ---
[Environment]::SetEnvironmentVariable('COPILOT_CODE_HOME', $InstallDir, 'User')
$env:COPILOT_CODE_HOME = $InstallDir

$userPaths = Split-PathList (Get-UserPath)
if (-not ($userPaths | Where-Object { Test-SameDir $_ $InstallDir })) {
    [Environment]::SetEnvironmentVariable('Path', ((@($userPaths) + $InstallDir) -join ';'), 'User')
}
if (-not (Split-PathList $env:Path | Where-Object { Test-SameDir $_ $InstallDir })) {
    $env:Path = "$env:Path;$InstallDir"
}

Write-Host "インストールしました: $InstallDir"
Write-Host '  COPILOT_CODE_HOME を登録し、Path に追加しました。'
Write-Host '  新しく開いたターミナルから、作業したいディレクトリで copilot-code と入力してください。'
