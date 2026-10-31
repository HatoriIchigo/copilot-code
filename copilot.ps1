param(
    [int]$Port = 9222,
    [switch]$Test,
    [int]$TimeoutSec = 900,
    [int]$MaxSteps = 10,
    [int]$MaxOutputChars = 50000,   # ツール1回の出力の上限。Copilot へはこの範囲を全文送る
    [int]$CellPreviewChars = 1000,  # Observation: セルに書く上限。超えた分は Copilot へ送るだけでシートには載せない
    [switch]$ReadOnly,              # 指定すると Write / Update を無効にする
    [int]$ReplyTimeoutSec = 300,
    [int]$ShellTimeoutSec = 300,    # Shell ツール1回の実行時間の上限
    [switch]$BringToFront,          # 指定すると、送信のたびに Excel を前面に出してフォーカスを移す
    [int]$PreambleEvery = 20        # ツールの説明つきの長いプロンプトを送り直す間隔 (ユーザ入力の回数)
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes
Add-Type @'
using System;
using System.Text;
using System.Runtime.InteropServices;
public class CopilotWin {
    delegate bool EnumProc(IntPtr h, IntPtr l);
    [DllImport("user32.dll")] static extern bool EnumChildWindows(IntPtr p, EnumProc cb, IntPtr l);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetClassName(IntPtr h, StringBuilder s, int n);
    [DllImport("user32.dll")] static extern bool SetForegroundWindow(IntPtr h);
    [DllImport("user32.dll")] static extern IntPtr SetFocus(IntPtr h);
    [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
    [DllImport("user32.dll")] static extern bool AttachThreadInput(uint a, uint b, bool f);
    [DllImport("kernel32.dll")] static extern uint GetCurrentThreadId();

    public static IntPtr FindChild(IntPtr parent, string className) {
        IntPtr found = IntPtr.Zero;
        EnumChildWindows(parent, (h, l) => {
            var s = new StringBuilder(256);
            GetClassName(h, s, 256);
            if (s.ToString() == className) { found = h; return false; }
            return true;
        }, IntPtr.Zero);
        return found;
    }

    public static void Focus(IntPtr main, IntPtr target) {
        uint pid;
        uint tid = GetWindowThreadProcessId(main, out pid);
        uint me = GetCurrentThreadId();
        SetForegroundWindow(main);
        AttachThreadInput(me, tid, true);
        SetFocus(target);
        AttachThreadInput(me, tid, false);
    }
}
'@

function Write-Log($msg) { Write-Host ("[{0:HH:mm:ss}] {1}" -f (Get-Date), $msg) }

# --- Chrome DevTools Protocol (WebView2 の Copilot ペインを操作する) ---

function Connect-Cdp {
    $t = Invoke-RestMethod "http://127.0.0.1:$Port/json" | Where-Object { $_.type -eq 'page' -and $_.title -like '*Copilot*' } | Select-Object -First 1
    if (-not $t) { return $null }
    $ws = New-Object Net.WebSockets.ClientWebSocket
    $ws.ConnectAsync([Uri]$t.webSocketDebuggerUrl, [Threading.CancellationToken]::None).Wait()
    $script:cdpId = 0
    $ws
}

function Invoke-Cdp($ws, [string]$method, $params = @{}) {
    $script:cdpId++
    $id = $script:cdpId
    $json = @{ id = $id; method = $method; params = $params } | ConvertTo-Json -Depth 10 -Compress
    $bytes = [Text.Encoding]::UTF8.GetBytes($json)
    $cts = New-Object Threading.CancellationTokenSource 30000
    $ws.SendAsync([ArraySegment[byte]]$bytes, 'Text', $true, $cts.Token).Wait()
    while ($true) {
        $ms = New-Object IO.MemoryStream
        $buf = New-Object byte[] 65536
        do {
            $r = $ws.ReceiveAsync([ArraySegment[byte]]$buf, $cts.Token).Result
            $ms.Write($buf, 0, $r.Count)
        } until ($r.EndOfMessage)
        $msg = [Text.Encoding]::UTF8.GetString($ms.ToArray()) | ConvertFrom-Json
        if ($msg.id -eq $id) { return $msg }
    }
}

function Invoke-CdpEval($ws, [string]$expression) {
    $res = Invoke-Cdp $ws 'Runtime.evaluate' @{ expression = $expression; returnByValue = $true }
    if ($res.result.exceptionDetails) { throw ($res.result.exceptionDetails | ConvertTo-Json -Depth 5) }
    $res.result.result.value
}

$jsInput = @'
document.querySelector('[aria-label="チャット入力"]')
'@
$jsReplyCount = @'
document.querySelectorAll('.fai-CopilotMessage__content').length
'@
$jsUserCount = @'
document.querySelectorAll('.fai-UserMessage__message').length
'@
# 長いプロンプト (ツールの説明つき) を最後に送ってから数えたユーザ入力の回数。起動直後は必ず長いほうを送る
$script:turnsSinceFull = $PreambleEvery
$jsGenerating = @'
!!Array.from(document.querySelectorAll('button')).find(b => b.getAttribute('aria-label') === 'Stop')
'@
$jsLastUserMessage = @'
(() => { const m = document.querySelectorAll('.fai-UserMessage__message'); return m.length ? m[m.length - 1].innerText : '' })()
'@
# 返信の全文(text)、コードブロックを除いた本文(prose)、コードブロックの中身(codes)を返す
$jsLastReply = @'
(() => {
  const m = document.querySelectorAll('.fai-CopilotMessage__content');
  if (!m.length) return null;
  const el = m[m.length - 1];
  const pres = Array.from(el.querySelectorAll('pre'));
  const codes = pres.map(p => (p.querySelector('code') || p).innerText.replace(/\r?\n$/, ''));
  const text = el.innerText;
  const old = pres.map(p => p.style.display);
  pres.forEach(p => { p.style.display = 'none'; });
  const prose = el.innerText;
  pres.forEach((p, i) => { p.style.display = old[i]; });
  return JSON.stringify({ text: text, prose: prose, codes: codes });
})()
'@

# --- Excel UI ---

function Get-ExcelWindow($title) {
    $p = Get-Process EXCEL -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowHandle -ne 0 -and $_.MainWindowTitle -like "$title*" } | Select-Object -First 1
    if ($p) { [Windows.Automation.AutomationElement]::FromHandle($p.MainWindowHandle) }
}

function Find-ByProperty($root, $property, $value) {
    $root.FindFirst('Descendants', (New-Object Windows.Automation.PropertyCondition($property, $value)))
}

function Test-CopilotPaneOpen($win) {
    # MsoWorkPane はステータスバーにも使われているので、名前でも絞る
    $cond = New-Object Windows.Automation.AndCondition(
        (New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::ClassNameProperty, 'MsoWorkPane')),
        (New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::NameProperty, 'Copilot')))
    $pane = $win.FindFirst('Descendants', $cond)
    [bool]($pane -and -not $pane.Current.IsOffscreen)
}

function Wait-Until([scriptblock]$cond, [int]$sec, [int]$intervalMs = 300) {
    $deadline = (Get-Date).AddSeconds($sec)
    while ((Get-Date) -lt $deadline) {
        $v = & $cond
        if ($v) { return $v }
        Start-Sleep -Milliseconds $intervalMs
    }
    $null
}

# Copilot ペインを開き(閉じていれば)、WebView2 にフォーカスを移す
function Show-CopilotPane($win) {
    if (-not (Test-CopilotPaneOpen $win)) {
        $button = Wait-Until { Find-ByProperty $win ([Windows.Automation.AutomationElement]::NameProperty) 'Copilot Chat' } 30
        if (-not $button) { throw 'Copilot Chat button not found.' }
        $button.GetCurrentPattern([Windows.Automation.InvokePattern]::Pattern).Invoke()
        if (-not (Wait-Until { Test-CopilotPaneOpen $win } 15)) { throw 'Copilot pane did not open.' }
    }
    $main = [IntPtr]$win.Current.NativeWindowHandle
    $webView = Wait-Until { $h = [CopilotWin]::FindChild($main, 'Chrome_WidgetWin_0'); if ($h -ne [IntPtr]::Zero) { $h } } 15
    # 既定では Excel を前面に出さない (他の作業を邪魔しない)。入力は CDP で直接届ける
    if ($webView -and $BringToFront) { [CopilotWin]::Focus($main, $webView) }
}

# C2 の文字列を Copilot に送り、回答テキストを返す
# $short があるときは、チャットが空 (新しいチャット) か、前回の長いプロンプトから $PreambleEvery 回たったときだけ
# $text (長いほう) を送り、それ以外は $short を送る。$short が無ければ常に $text
function Ask-Copilot($win, [string]$text, $short = $null) {
    Show-CopilotPane $win
    # ペインが開いて入力欄が描画されるまで待つ
    $ws = Wait-Until {
        try {
            $w = Connect-Cdp
            if ($w) {
                if (Invoke-CdpEval $w "!!($($jsInput.Trim()))") { $w } else { $w.Dispose() }
            }
        } catch { $script:cdpError = $_.Exception.Message; $null }
    } 45 1000
    if (-not $ws) { throw "Copilot pane (CDP) not reachable. $script:cdpError" }
    try {
        $before = [int](Invoke-CdpEval $ws $jsReplyCount)
        if ($null -ne $short) {
            $users = [int](Invoke-CdpEval $ws $jsUserCount)
            if ($users -eq 0 -or $script:turnsSinceFull -ge $PreambleEvery) {
                $script:turnsSinceFull = 1
                Write-Log 'Prompt: full (tool instructions included)'
            } else {
                $text = [string]$short
                $script:turnsSinceFull++
                Write-Log "Prompt: short ($script:turnsSinceFull / $PreambleEvery)"
            }
        }
        # ウィンドウが非アクティブでも、ページを「フォーカスあり」として扱わせる
        Invoke-Cdp $ws 'Emulation.setFocusEmulationEnabled' @{ enabled = $true } | Out-Null
        Invoke-CdpEval $ws "($($jsInput.Trim())).focus()" | Out-Null
        # 入力欄は Enter で送信されるため、改行は Shift+Enter で入れる
        $lines = $text -split "\r?\n"
        for ($i = 0; $i -lt $lines.Count; $i++) {
            if ($i -gt 0) {
                $key = @{ key = 'Enter'; code = 'Enter'; windowsVirtualKeyCode = 13; modifiers = 8 }
                Invoke-Cdp $ws 'Input.dispatchKeyEvent' ($key + @{ type = 'keyDown'; text = "`r" }) | Out-Null
                Invoke-Cdp $ws 'Input.dispatchKeyEvent' ($key + @{ type = 'keyUp' }) | Out-Null
            }
            if ($lines[$i] -ne '') { Invoke-Cdp $ws 'Input.insertText' @{ text = $lines[$i] } | Out-Null }
        }
        Start-Sleep -Milliseconds 500
        $sent = Invoke-CdpEval $ws "(() => { const b = document.querySelector('button[aria-label=""Send""]'); if (b) { b.click(); return true } return false })()"
        if (-not $sent) { throw 'Send button not found.' }
        # 返信が増え、かつ生成が止まるまで待つ
        $done = Wait-Until {
            ([int](Invoke-CdpEval $ws $jsReplyCount)) -gt $before -and -not (Invoke-CdpEval $ws $jsGenerating)
        } $ReplyTimeoutSec 1000
        if (-not $done) { throw "Timed out waiting for Copilot reply (${ReplyTimeoutSec}s)." }
        Start-Sleep -Milliseconds 500
        # 送った全文が（最終行まで）届いたかを確認する
        $sentText = Invoke-CdpEval $ws $jsLastUserMessage
        $tail = ($lines | Where-Object { $_.Trim() -ne '' } | Select-Object -Last 1) -replace '\s', ''
        if (($sentText -replace '\s', '') -notlike "*$tail*") {
            $msg = "Prompt was not fully delivered. sent=[$sentText]"
            if ($script:Test) { throw $msg } else { Write-Log "WARN: $msg" }
        }
        $parts = (Invoke-CdpEval $ws $jsLastReply) | ConvertFrom-Json
        @{ Text = ([string]$parts.text).Trim(); Prose = ([string]$parts.prose).Trim(); Codes = @($parts.codes) }
    } finally {
        $ws.Dispose()
    }
}

# テンプレートのプレースホルダを置換する。pairs は @(パターン, 値) の配列で、前から順に置換する
function Expand-Template([string]$template, $pairs) {
    $ic = [Text.RegularExpressions.RegexOptions]::IgnoreCase
    $text = $template
    foreach ($p in $pairs) {
        $value = $p[1]
        $text = [regex]::Replace($text, $p[0], [Text.RegularExpressions.MatchEvaluator]{ param($m) $value }, $ic)
    }
    $text
}

# プロンプトシートのテンプレートに、設定シートとチャットシートの値を埋め込む
function Build-Prompt($wb, [string]$userText, [switch]$Short) {
    # 2回目以降用 (B4) が空なら $null を返し、常に初回用 (B2) を送る
    $cell = if ($Short) { 'B4' } else { 'B2' }
    $template = [string]$wb.Worksheets.Item('プロンプト').Range($cell).Value2
    if ($Short -and $template.Trim() -eq '') { return $null }
    $directory = [string]$wb.Worksheets.Item('設定').Range('C2').Value2
    # ユーザ入力は最後に埋め込み、その中の文字列が再置換されないようにする
    Expand-Template $template @(
        @('__SUMMARY__', ''),   # 要約は未対応なので空で埋める
        @('%DIRECTORY%', $directory),
        @('__PROMPT__', $userText))
}

# ツール実行結果を Copilot に返すプロンプト（プロンプトシートの B3）
function Build-Followup($wb, [string]$action, [string]$observation) {
    $template = [string]$wb.Worksheets.Item('プロンプト').Range('B3').Value2
    if ($template.Trim() -eq '') { $template = "Action: __ACTION__`n`n__OBSERVATION__" }
    $directory = [string]$wb.Worksheets.Item('設定').Range('C2').Value2
    Expand-Template $template @(
        @('%DIRECTORY%', $directory),
        @('__ACTION__', $action),
        @('__OBSERVATION__', $observation))
}

function Limit-Text([string]$text, [int]$max) {
    if ($text.Length -le $max) { return $text }
    $text.Substring(0, $max) + "`n...(以降省略: $($text.Length - $max) 文字)"
}

# 先頭が = + - @ だと数式扱いになるので、文字列として書き込めるようにする
function Protect-Formula([string]$text) {
    if ($text -match '^[=+\-@]') { "'" + $text } else { $text }
}

# 応答の末尾にある「Action: ツール名(引数)」を取り出す。無ければ $null
function Get-Action([string]$text) {
    $lines = @($text -split "\r?\n" | Where-Object { $_.Trim() -ne '' } | Select-Object -Last 3)
    [array]::Reverse($lines)
    foreach ($l in $lines) {
        # 行全体を囲む Markdown の強調 (` や *) だけ外す。Shell のコマンド内の * や ` は残す
        $s = $l.Trim().Trim('`', '*').Trim()
        if ($s -match '^Action:\s*(\w+)\((.*)\)\s*$') {
            $a = $Matches[2].Trim()
            if ($Matches[1] -ne 'Shell') { $a = $a.Trim('"', "'") }
            return @{ Tool = $Matches[1]; Arg = $a }
        }
    }
    $null
}

# UTF-8 のテキストとして読む。UTF-8 でない・バイナリのときは Error を返す
function Read-Utf8Strict([string]$full) {
    $bytes = [IO.File]::ReadAllBytes($full)
    $bom = $bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF
    $skip = if ($bom) { 3 } else { 0 }
    try { $text = (New-Object Text.UTF8Encoding($false, $true)).GetString($bytes, $skip, $bytes.Length - $skip) }
    catch { return @{ Error = 'ERROR: UTF-8 以外のファイルは編集できません' } }
    if ($text.Contains([string][char]0)) { return @{ Error = 'ERROR: バイナリファイルは編集できません' } }
    @{ Text = $text; Bom = $bom }
}

# 書き換える前のファイルを、ディレクトリの外のバックアップ置き場に退避する。退避先のパスを返す
function Backup-File([string]$full, [string]$rootFull) {
    $rel = $full.Substring($rootFull.Length).TrimStart('\')
    $dest = Join-Path (Join-Path $BackupRoot (Get-Date -Format 'yyyyMMdd-HHmmss-fff')) $rel
    New-Item -ItemType Directory -Force (Split-Path $dest) | Out-Null
    Copy-Item -LiteralPath $full -Destination $dest
    $dest
}

# 書き込みを拒否する場合はその理由を返す。許可なら $null
function Get-WriteDenied([string]$full, [string]$rootFull) {
    if ($ReadOnly) { return 'ERROR: 書き込みツールは無効です (-ReadOnly 指定)' }
    if (($full.Substring($rootFull.Length) -split '\\') -contains '.git') { return 'ERROR: .git 配下は変更できません' }
    $null
}

# PowerShell のコマンドを実行する。作業ディレクトリは root。ディレクトリの外へは閉じ込められない
function Invoke-Shell([string]$command, [string]$root) {
    if ($ReadOnly) { return 'ERROR: Shell は無効です (-ReadOnly 指定)' }
    if ([string]::IsNullOrWhiteSpace($command)) { return 'ERROR: 実行するコマンドがありません' }
    $script = '[Console]::OutputEncoding=[Text.Encoding]::UTF8; $ProgressPreference="SilentlyContinue"; ' + $command
    # -EncodedCommand だと標準エラーが CLIXML になるので、一時 .ps1 (UTF-8 BOM) に書いて -File で実行する
    $tmp = Join-Path ([IO.Path]::GetTempPath()) ("copilot-shell-" + [guid]::NewGuid().ToString('N') + '.ps1')
    [IO.File]::WriteAllText($tmp, $script, (New-Object Text.UTF8Encoding($true)))
    $psi = New-Object Diagnostics.ProcessStartInfo
    $psi.FileName = 'powershell.exe'
    $psi.Arguments = "-NoProfile -NonInteractive -ExecutionPolicy Bypass -File `"$tmp`""
    $psi.WorkingDirectory = $root
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true
    $psi.RedirectStandardInput = $true
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.StandardOutputEncoding = [Text.Encoding]::UTF8
    $psi.StandardErrorEncoding = [Text.Encoding]::UTF8
    $p = [Diagnostics.Process]::Start($psi)
    $p.StandardInput.Close()
    $outTask = $p.StandardOutput.ReadToEndAsync()
    $errTask = $p.StandardError.ReadToEndAsync()
    $timedOut = -not $p.WaitForExit($ShellTimeoutSec * 1000)
    if ($timedOut) {
        # 自分が起動したプロセスだけを、子プロセスごと止める
        & taskkill.exe /T /F /PID $p.Id 2>&1 | Out-Null
        $p.WaitForExit(5000) | Out-Null
    }
    $out = ''; $err = ''
    if ($outTask.Wait(5000)) { $out = $outTask.Result }
    if ($errTask.Wait(5000)) { $err = $errTask.Result }
    Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
    $text = "[stdout]`n$($out.TrimEnd())"
    if ($err.Trim() -ne '') { $text += "`n[stderr]`n$($err.TrimEnd())" }
    if ($timedOut) { return "ERROR: ${ShellTimeoutSec} 秒でタイムアウトしたため中止しました`n$text" }
    $text += "`n[exit code] $($p.ExitCode)"
    if ($text.Length -gt $MaxOutputChars) { $text = $text.Substring(0, $MaxOutputChars) + "`n...(出力が長いため $MaxOutputChars 文字で切りました)" }
    $text
}

# ツールを実行する。参照・変更できるのは root（設定シートのディレクトリ）の下だけ。
# payload は Write / Update が書き込む内容（Copilot の返信のコードブロック）
function Invoke-Tool([string]$tool, [string]$arg, [string]$root, $payload = $null) {
    $rootFull = [IO.Path]::GetFullPath($root).TrimEnd('\')
    if ($tool -eq 'Shell') {
        # 引数はコマンドそのもの。引数が空で、コードブロックがあるときはその内容を実行する
        $cmd = if ($arg.Trim() -eq '' -and $null -ne $payload) { [string]$payload } else { $arg.Trim() }
        return Invoke-Shell $cmd $rootFull
    }
    $pathArg = $arg
    $start = 1
    $count = [long][int]::MaxValue
    $endArg = 0
    $grepPattern = ''
    if ($tool -eq 'Read' -and $arg -match '^(.*?)\s*,\s*(\d+)(?:\s*,\s*(\d+))?\s*$') {
        # Read は「パス, 開始行, 行数」の形でも呼べる
        $pathArg = $Matches[1].Trim().Trim('"', "'")
        $start = [Math]::Max(1, [int]$Matches[2])
        if ($Matches[3]) { $count = [long]$Matches[3] }
    } elseif ($tool -eq 'Update') {
        if ($arg -notmatch '^(.*?)\s*,\s*(\d+)\s*,\s*(\d+)\s*$') { return 'ERROR: Update(パス, 開始行, 終了行) の形で指定してください' }
        $pathArg = $Matches[1].Trim().Trim('"', "'")
        $start = [int]$Matches[2]
        $endArg = [int]$Matches[3]
    } elseif ($tool -eq 'Grep') {
        # Grep は「パターン」または「パターン, パス」。最後のカンマ以降が実在するパスのときだけパスとみなす
        $pathArg = '.'
        $grepPattern = $arg.Trim().Trim('"', "'")
        if ($arg -match '^(.*),\s*([^,]+?)\s*$') {
            $pat = $Matches[1].Trim().Trim('"', "'")
            $cand = $Matches[2].Trim().Trim('"', "'")
            $candFull = try { [IO.Path]::GetFullPath($(if ([IO.Path]::IsPathRooted($cand)) { $cand } else { Join-Path $rootFull $cand })) } catch { $null }
            if ($candFull -and (Test-Path -LiteralPath $candFull)) { $grepPattern = $pat; $pathArg = $cand }
        }
    }

    $target = if ($pathArg -eq '' -or $pathArg -eq '.') { $rootFull }
              elseif ([IO.Path]::IsPathRooted($pathArg)) { $pathArg }
              else { Join-Path $rootFull $pathArg }
    try { $full = [IO.Path]::GetFullPath($target).TrimEnd('\') } catch { return "ERROR: パスが不正です: $pathArg" }
    if ($full -ne $rootFull -and -not $full.StartsWith($rootFull + '\', [StringComparison]::OrdinalIgnoreCase)) {
        return "ERROR: ディレクトリの外は参照できません: $pathArg"
    }

    switch ($tool) {
        'List' {
            if (-not (Test-Path -LiteralPath $full -PathType Container)) { return "ERROR: ディレクトリではありません: $pathArg" }
            $items = @(Get-ChildItem -LiteralPath $full -Force | Sort-Object @{ Expression = { -not $_.PSIsContainer } }, Name)
            $names = @($items | Select-Object -First 200 | ForEach-Object { if ($_.PSIsContainer) { $_.Name + '/' } else { $_.Name } })
            $out = $names -join "`n"
            if ($items.Count -gt 200) { $out += "`n...(他 $($items.Count - 200) 件)" }
            if ($out -eq '') { return '(空のディレクトリ)' }
            return $out
        }
        'Read' {
            if (-not (Test-Path -LiteralPath $full -PathType Leaf)) { return "ERROR: ファイルではありません: $pathArg" }
            if ((Get-Item -LiteralPath $full -Force).Length -gt 5MB) { return 'ERROR: ファイルが大きすぎます (5MB 超)' }
            $text = [IO.File]::ReadAllText($full, [Text.Encoding]::UTF8)
            if ($text.Contains([string][char]0)) { return 'ERROR: バイナリファイルは読めません' }

            # 行番号を付けて返す。出力が上限を超えるときは途中で切り、続きの開始行を知らせる
            $all = @($text -split "\r?\n")
            if ($all.Count -gt 1 -and $all[$all.Count - 1] -eq '') { $all = $all[0..($all.Count - 2)] }
            $total = $all.Count
            if ($start -gt $total) { return "ERROR: 開始行が範囲外です (全 $total 行)" }
            $end = [int][Math]::Min([long]$total, $start + $count - 1)
            $out = New-Object 'System.Collections.Generic.List[string]'
            $len = 0
            $last = $start - 1
            for ($i = $start; $i -le $end; $i++) {
                $line = "${i}: " + $all[$i - 1]
                if ($line.Length -gt $MaxOutputChars) { $line = $line.Substring(0, $MaxOutputChars) + '...(行が長いため省略)' }
                if ($out.Count -gt 0 -and $len + $line.Length + 1 -gt $MaxOutputChars) { break }
                $out.Add($line)
                $len += $line.Length + 1
                $last = $i
            }
            $body = $out -join "`n"
            if ($last -lt $end) {
                return $body + "`n...(出力上限のため途中まで。全 $total 行のうち $start〜$last 行を表示。続きは Read($pathArg, $($last + 1)) で読めます)"
            }
            if ($start -gt 1 -or $end -lt $total) { return $body + "`n(全 $total 行のうち $start〜$end 行を表示)" }
            return $body
        }
        'Grep' {
            if (-not (Test-Path -LiteralPath $full)) { return "ERROR: 見つかりません: $pathArg" }
            if ($grepPattern -eq '') { return 'ERROR: 検索パターンが空です' }
            try { $rx = New-Object Text.RegularExpressions.Regex($grepPattern, [Text.RegularExpressions.RegexOptions]::IgnoreCase) }
            catch { return "ERROR: 正規表現が不正です: $grepPattern" }
            $files = if (Test-Path -LiteralPath $full -PathType Leaf) { ,(Get-Item -LiteralPath $full -Force) }
                     else { Get-ChildItem -LiteralPath $full -Recurse -File -Force -ErrorAction SilentlyContinue | Where-Object { $_.FullName -notmatch '\\\.git\\' } }
            $out = New-Object 'System.Collections.Generic.List[string]'
            $len = 0
            $hits = 0
            $truncated = $false
            :files foreach ($f in $files) {
                if ($f.Length -gt 5MB) { continue }
                try { $t = [IO.File]::ReadAllText($f.FullName, [Text.Encoding]::UTF8) } catch { continue }
                if ($t.Contains([string][char]0)) { continue }
                $rel = $f.FullName.Substring($rootFull.Length).TrimStart('\')
                $ls = $t -split "\r?\n"
                for ($i = 0; $i -lt $ls.Count; $i++) {
                    if (-not $rx.IsMatch($ls[$i])) { continue }
                    $text = $ls[$i]
                    if ($text.Length -gt 300) { $text = $text.Substring(0, 300) + '...' }
                    $line = "${rel}:$($i + 1): $text"
                    if ($hits -ge 200 -or ($out.Count -gt 0 -and $len + $line.Length + 1 -gt $MaxOutputChars)) { $truncated = $true; break files }
                    $out.Add($line)
                    $len += $line.Length + 1
                    $hits++
                }
            }
            if ($hits -eq 0) { return '(一致なし)' }
            $body = $out -join "`n"
            if ($truncated) { return $body + "`n...(件数または出力の上限のため、ここまで $hits 件。パターンやパスを絞ってください)" }
            return $body + "`n($hits 件)"
        }
        'Write' {
            $denied = Get-WriteDenied $full $rootFull
            if ($denied) { return $denied }
            if ($full -eq $rootFull -or (Test-Path -LiteralPath $full -PathType Container)) { return "ERROR: ファイルを指定してください: $pathArg" }
            if ($null -eq $payload) { return 'ERROR: 書き込む内容がありません。Action の次に、内容をコードブロック (```) で書いてください' }
            $exists = Test-Path -LiteralPath $full -PathType Leaf
            $bom = $false
            $nl = "`n"
            $bak = ''
            if ($exists) {
                $r = Read-Utf8Strict $full
                if ($r.Error) { return $r.Error }
                $bom = $r.Bom
                if ($r.Text.Contains("`r`n")) { $nl = "`r`n" }
                $bak = Backup-File $full $rootFull
            }
            $newLines = if ($payload -eq '') { @() } else { @($payload -split "\r?\n") }
            $body = $newLines -join $nl
            if ($newLines.Count -gt 0) { $body += $nl }
            New-Item -ItemType Directory -Force (Split-Path $full) | Out-Null
            [IO.File]::WriteAllText($full, $body, (New-Object Text.UTF8Encoding($bom)))
            $what = if ($exists) { "上書きしました。バックアップ: $bak" } else { '新規作成しました' }
            return "OK: $pathArg を$what ($($newLines.Count) 行, $($body.Length) 文字)"
        }
        'Update' {
            $denied = Get-WriteDenied $full $rootFull
            if ($denied) { return $denied }
            if (-not (Test-Path -LiteralPath $full -PathType Leaf)) { return "ERROR: ファイルではありません: $pathArg" }
            if ($null -eq $payload) { return 'ERROR: 書き込む内容がありません。Action の次に、内容をコードブロック (```) で書いてください。行を削除するときは空のコードブロックにしてください' }
            $r = Read-Utf8Strict $full
            if ($r.Error) { return $r.Error }
            $nl = if ($r.Text.Contains("`r`n")) { "`r`n" } else { "`n" }
            $trail = $r.Text.EndsWith("`n")
            $list = New-Object 'System.Collections.Generic.List[string]'
            if ($r.Text -ne '') {
                $list.AddRange([string[]]($r.Text -split "\r?\n"))
                if ($trail) { $list.RemoveAt($list.Count - 1) }
            }
            $total = $list.Count
            $s = $start
            $e = $endArg
            if ($s -lt 1 -or $s -gt $total + 1) { return "ERROR: 開始行が範囲外です (全 $total 行。1〜$($total + 1) を指定)" }
            if ($e -lt $s - 1 -or $e -gt $total) { return "ERROR: 終了行が範囲外です (全 $total 行。$($s - 1)〜$total を指定)" }
            $newLines = if ($payload -eq '') { @() } else { @($payload -split "\r?\n") }
            $bak = Backup-File $full $rootFull
            $list.RemoveRange($s - 1, $e - $s + 1)
            if ($newLines.Count -gt 0) { $list.InsertRange($s - 1, [string[]]$newLines) }
            $body = $list -join $nl
            if ($trail -and $list.Count -gt 0) { $body += $nl }
            [IO.File]::WriteAllText($full, $body, (New-Object Text.UTF8Encoding($r.Bom)))

            $what = if ($e -lt $s) { "$s 行目の前に $($newLines.Count) 行を挿入しました" } else { "$s〜$e 行を $($newLines.Count) 行で置き換えました" }
            $from = [Math]::Max(1, $s - 2)
            $to = [Math]::Min($list.Count, $s + $newLines.Count + 1)
            $ctx = if ($to -ge $from) { ($from..$to | ForEach-Object { "${_}: " + $list[$_ - 1] }) -join "`n" } else { '(空)' }
            return "OK: $pathArg の$what (全 $total 行 → $($list.Count) 行)。バックアップ: $bak`n更新後の該当部分:`n$ctx"
        }
        default { return "ERROR: 未対応のツールです: $tool (使えるのは List, Read, Grep, Write, Update, Shell)" }
    }
}

# 1回のユーザ入力を処理する。Action が返る間はツールを実行して結果を返し続ける。
# 戻り値は次の入力待ち USER: 行の番号
function Invoke-Turn($wb, $sheet, [string]$title, [int]$inputRow, [string]$userText) {
    $directory = [string]$wb.Worksheets.Item('設定').Range('C2').Value2
    $row = $inputRow
    try {
        $prompt = Build-Prompt $wb $userText
        $short = Build-Prompt $wb $userText -Short
        foreach ($p in $prompt, $short) {
            if ($Test -and $p -and $p -match '__(PROMPT|SUMMARY)__|%DIRECTORY%') { throw "Unreplaced placeholder in prompt: $p" }
        }
        for ($step = 1; ; $step++) {
            $reply = @(Ask-Copilot (Get-ExcelWindow $title) $prompt $short)[-1]
            $short = $null   # 短いプロンプトの候補は、ユーザ入力の最初の送信だけ。ツール結果の返信には使わない
            $answer = $reply.Text
            Write-Log "Reply: $answer"
            $row++
            Set-Cell $sheet $row 2 'Result:'
            Set-Cell $sheet $row 3 (Protect-Formula (Limit-Text $answer 30000))
            # Action は、コードブロックを除いた本文の末尾から探す。Write / Update の内容は最後のコードブロック
            $action = Get-Action $reply.Prose
            if (-not $action) { break }
            $payload = if ($reply.Codes.Count -gt 0) { [string]$reply.Codes[$reply.Codes.Count - 1] } else { $null }
            if ($step -ge $MaxSteps) { Write-Log "Reached MaxSteps ($MaxSteps)."; break }

            $call = "$($action.Tool)($($action.Arg))"
            $obs = [string](Invoke-Tool $action.Tool $action.Arg $directory $payload)
            Write-Log "Action: $call ($($obs.Length) chars)"
            # シートには先頭のプレビューだけ載せ、全文は Copilot へ送る
            $preview = $obs
            if ($obs.Length -gt $CellPreviewChars) {
                $preview = $obs.Substring(0, $CellPreviewChars) + "`n...(全 $($obs.Length) 文字。全文は Copilot に送信済み)"
            }
            $row++
            Set-Cell $sheet $row 2 'Action:'
            # Write / Update は、書き込む内容も一緒に載せる
            $callCell = if ($null -ne $payload -and $action.Tool -in 'Write', 'Update') { $call + "`n" + (Limit-Text $payload $CellPreviewChars) } else { $call }
            Set-Cell $sheet $row 3 (Protect-Formula $callCell)
            $row++
            Set-Cell $sheet $row 2 'Observation:'
            Set-Cell $sheet $row 3 (Protect-Formula $preview)
            $prompt = Build-Followup $wb $call $obs
        }
    } catch {
        if ($Test) { throw }
        $msg = $_.Exception.Message
        Write-Log "Copilot request failed: $msg"
        $row++
        Set-Cell $sheet $row 2 'Error:'
        Set-Cell $sheet $row 3 $msg
    }
    $row++
    Set-Cell $sheet $row 2 'USER:'
    $row
}

# セルに書く。編集中などで Excel に拒否された場合は少し待って再試行する
function Set-Cell($sheet, [int]$row, [int]$col, [string]$value) {
    for ($i = 0; $i -lt 60; $i++) {
        try { $sheet.Cells($row, $col).Value2 = $value; return }
        catch [System.Runtime.InteropServices.COMException] { Start-Sleep -Milliseconds 500 }
    }
    throw "Could not write cell ($row, $col)."
}

# --- main ---

$launchDir = (Get-Location).ProviderPath
$master = Join-Path $PSScriptRoot 'copilot.xlsx'
if (-not (Test-Path $master)) {
    Write-Error "File not found: $master"
    exit 1
}

# マスタは触らず、コピーを開く
$workDir = Join-Path $env:TEMP ('copilot-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory $workDir | Out-Null
$copy = Join-Path $workDir 'copilot.xlsx'
Copy-Item $master $copy
$title = [IO.Path]::GetFileName($copy)

# テストでは、中身が決まったディレクトリを対象にする（List / Read の動作確認用）
if ($Test) {
    $launchDir = Join-Path $workDir 'data'
    New-Item -ItemType Directory $launchDir | Out-Null
    $notes = (1..150 | ForEach-Object { "行 ${_}: ダミーのテキストです。" }) + '最後の行: 合言葉は MOMIJI-4821 です。'
    Set-Content -Path (Join-Path $launchDir 'notes.txt') -Value $notes -Encoding UTF8
}
$defaultMaxOutputChars = $MaxOutputChars

# Write / Update が書き換える前のファイルの退避先（ディレクトリの外）
$BackupRoot = if ($Test) { Join-Path $workDir 'backup' } else { Join-Path $env:LOCALAPPDATA 'copilot-code\backup' }

# WebView2 の DevTools ポートと a11y を有効にして、新しい Excel インスタンスで開く
$env:WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS = "--remote-debugging-port=$Port --force-renderer-accessibility --disable-backgrounding-occluded-windows --disable-renderer-backgrounding --disable-background-timer-throttling"
Start-Process excel.exe -ArgumentList '/x', "`"$copy`""
Write-Log "Opened copy: $copy"

$exitCode = 0
$wb = $null
$app = $null
try {
    $win = Wait-Until { Get-ExcelWindow $title } 60 500
    if (-not $win) { throw 'Excel window not found.' }

    $wb = Wait-Until { try { [Runtime.InteropServices.Marshal]::BindToMoniker($copy) } catch { $null } } 60 500
    if (-not $wb) { throw 'Could not attach to the workbook.' }
    $app = $wb.Application
    $sheet = $wb.Worksheets.Item('チャット')
    $wb.Worksheets.Item('設定').Range('C2').Value2 = $launchDir

    # チャット欄は B列が USER: / Result:、C列が本文。入力待ちの USER: 行を探す（なければ作る）
    $lastRow = $sheet.Cells($sheet.Rows.Count, 2).End(-4162).Row   # xlUp
    $lastLabel = [string]$sheet.Cells($lastRow, 2).Value2
    if ($lastLabel -eq 'USER:') {
        $inputRow = $lastRow
    } else {
        $inputRow = if ($lastLabel -eq '') { 2 } else { $lastRow + 1 }
        Set-Cell $sheet $inputRow 2 'USER:'
    }
    Write-Log "Attached. Directory: $launchDir. Watching C$inputRow."

    # ペインは起動時に開いておく。開く操作で Excel が一瞬前面に出るので、
    # Excel が自分で前面に出る起動直後のうちに済ませ、送信時に他の作業を邪魔しないようにする
    try {
        $w = Wait-Until { Get-ExcelWindow $title } 30
        if ($w) { Show-CopilotPane $w; Write-Log 'Copilot pane ready.' }
    } catch { Write-Log "WARN: could not open the Copilot pane at startup: $($_.Exception.Message)" }

    $last = [string]$sheet.Cells($inputRow, 3).Value2
    $deadline = (Get-Date).AddSeconds($TimeoutSec)
    $injectedRow = 0
    $testIdx = 0
    $tests = @(
        @{ q = '日本の首都はどこですか？一言で答えてください。'; expect = '東京' },
        # expect: 最終回答に含まれるはずの文字列 / actions: Action: 行の最小数 / obs: Observation: に含まれるはずの文字列
        # cap: その回の出力上限 / check: 終了後にファイルの状態を検証する
        @{ q = 'List ツールを使って、このディレクトリにあるファイルの一覧を調べ、ファイル名をすべて答えてください。'; expect = 'notes.txt'; obs = 'notes.txt'; actions = 1 },
        # 出力上限を小さくして、途中で切れた続きを Copilot が行番号指定で読めるか
        @{ q = 'Read ツールで notes.txt を最後まで読み、最後の行に書かれている合言葉を答えてください。省略と表示されたら続きを読むこと。'; expect = 'MOMIJI-4821'; actions = 2; cap = 1500 },
        @{ q = 'Grep ツールで notes.txt から「MOMIJI」を含む行を探し、その行番号を答えてください。'; expect = '151'; obs = 'MOMIJI-4821'; actions = 1 },
        @{
            q = 'Write ツールで hello.txt を新規作成し、内容は「こんにちは、世界」の1行だけにしてください。'; actions = 1
            check = {
                $p = Join-Path $launchDir 'hello.txt'
                if (-not (Test-Path -LiteralPath $p)) { throw 'hello.txt が作成されていません' }
                $c = [IO.File]::ReadAllText($p, [Text.Encoding]::UTF8)
                if ($c -notlike '*こんにちは、世界*') { throw "hello.txt の内容が違います: $c" }
            }
        },
        @{
            q = 'Update ツールで notes.txt の 2 行目だけを「行 2: 書き換えました」に置き換えてください。'; actions = 1
            check = {
                $p = Join-Path $launchDir 'notes.txt'
                $b = [IO.File]::ReadAllBytes($p)
                if (-not ($b[0] -eq 0xEF -and $b[1] -eq 0xBB -and $b[2] -eq 0xBF)) { throw 'notes.txt の BOM が失われました' }
                $l = [IO.File]::ReadAllLines($p, [Text.Encoding]::UTF8)
                if ($l.Count -ne 151) { throw "notes.txt の行数が変わりました: $($l.Count)" }
                if ($l[1] -ne '行 2: 書き換えました') { throw "2 行目が違います: $($l[1])" }
                if ($l[0] -ne '行 1: ダミーのテキストです。' -or $l[2] -ne '行 3: ダミーのテキストです。') { throw '前後の行が変わっています' }
                if (-not (Get-ChildItem $BackupRoot -Recurse -Filter notes.txt -ErrorAction SilentlyContinue)) { throw 'バックアップがありません' }
            }
        },
        # Shell: 日本語の出力と作業ディレクトリ (ls 相当) を確認する
        @{ q = 'Shell ツールで PowerShell の「Write-Output ("あ" * 3 + (6 * 7)); Get-ChildItem -Name」を実行し、出力された内容をそのまま答えてください。'; expect = 'あああ42'; obs = 'notes.txt'; actions = 1 }
    )

    while ($true) {
        if ($Test -and (Get-Date) -gt $deadline) { throw 'Test timed out.' }

        try {
            if ($Test -and $injectedRow -ne $inputRow) {
                $MaxOutputChars = if ($tests[$testIdx].cap) { $tests[$testIdx].cap } else { $defaultMaxOutputChars }
                Set-Cell $sheet $inputRow 3 $tests[$testIdx].q
                $injectedRow = $inputRow
                Write-Log "TEST: wrote C$inputRow."
            }

            $cur = [string]$sheet.Cells($inputRow, 3).Value2
            if ($cur -ne $last -and $cur.Trim() -ne '') {
                $last = $cur
                Write-Log "C$inputRow changed: $cur"
                $startRow = $inputRow
                $inputRow = @(Invoke-Turn $wb $sheet $title $inputRow $cur)[-1]
                $last = ''
                Write-Log "Watching C$inputRow."

                if ($Test) {
                    $t = $tests[$testIdx]
                    $resultRow = $inputRow - 1
                    $res = [string]$sheet.Cells($resultRow, 3).Value2
                    if ($t.expect -and $res -notlike "*$($t.expect)*") { throw "TEST FAIL (round $($testIdx + 1): C$resultRow = $res)" }
                    if ($t.check) {
                        try { & $t.check } catch { throw "TEST FAIL (round $($testIdx + 1)): $($_.Exception.Message)" }
                    }
                    if ($t.actions) {
                        $actCount = @($startRow..$resultRow | Where-Object { [string]$sheet.Cells($_, 2).Value2 -eq 'Action:' }).Count
                        if ($actCount -lt $t.actions) { throw "TEST FAIL (round $($testIdx + 1): Action 行が $actCount 件 (期待 $($t.actions) 件以上))" }
                    }
                    if ($t.obs) {
                        $obsText = ($startRow..$resultRow | Where-Object { [string]$sheet.Cells($_, 2).Value2 -eq 'Observation:' } |
                            ForEach-Object { [string]$sheet.Cells($_, 3).Value2 }) -join "`n"
                        if ($obsText -notlike "*$($t.obs)*") { throw "TEST FAIL (round $($testIdx + 1): Observation = $obsText)" }
                    }
                    Write-Log "TEST round $($testIdx + 1) PASS (C$resultRow = $res)"
                    $testIdx++
                    if ($testIdx -ge $tests.Count) {
                        $labels = (2..$inputRow | ForEach-Object { [string]$sheet.Cells($_, 2).Value2 }) -join ','
                        $parts = $tests | ForEach-Object { if ($_.actions) { 'USER:,Result:(,Action:,Observation:,Result:)+' } else { 'USER:,Result:' } }
                        $want = '^' + ($parts -join ',') + ',USER:$'
                        if ($labels -notmatch $want) { throw "TEST FAIL (B column = $labels, expected pattern $want)" }
                        Write-Log "TEST PASS (B column = $labels)"
                        break
                    }
                }
            } elseif ($cur -ne $last) {
                $last = $cur
            }
        } catch [System.Runtime.InteropServices.COMException] {
            # ブックが生きていれば、セル編集中などで呼び出しが拒否されただけなので再試行
            $alive = $false
            try { $null = $wb.Name; $alive = $true } catch { }
            if ($alive) { Start-Sleep -Milliseconds 500; continue }
            Write-Log 'Workbook closed. Exiting.'
            $wb = $null
            break
        }
        Start-Sleep -Milliseconds 500
    }
} catch {
    Write-Log "ERROR: $($_.Exception.Message)"
    $exitCode = 1
} finally {
    if ($Test -and $app) {
        try { $app.DisplayAlerts = $false; $app.Quit() } catch { }
        Start-Sleep 2
    }
    if ($Test -or -not $wb) {
        Remove-Item $workDir -Recurse -Force -ErrorAction SilentlyContinue
    }
}
exit $exitCode
