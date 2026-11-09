<#
.SYNOPSIS
  추가기능(addins/<id>/)을 Windows용 Excel에서 컴파일해 <id>.xlam과 <id>.zip으로 만든다.
.DESCRIPTION
  1) addins/<id>/*.bas(.cls) + addins/_common/*.bas 를 CP949+CRLF로 임시 변환
  2) 보이지 않는 Excel(Visible=False)에서 새 통합 문서에 가져오기, ThisWorkbook_*.txt가 있으면 ThisWorkbook에 넣기
  3) VBE 'VBAProject 컴파일'(CommandBars Id 578)이 끝나 비활성이 되는지로 컴파일 성공 판정
  4) out\<id>.xlam(xlFileFormat 55)과 out\<id>.zip(docs/03 패키지 규격) 만들기
  키보드 입력·SendKeys·화면 조작은 하지 않는다. 이 스크립트가 띄운 Excel만 종료한다.
.PARAMETER Addin     addins\ 아래 폴더 이름(= 추가기능 id). `_common`을 주면 공용 라이브러리만 컴파일해 clockout-lib.xlam/.zip으로 만든다
.PARAMETER RunTests  Test_ 로 시작하는 Public Sub를 시험용 사본에서 실행(창이 뜨지 않게 MsgBox는 기록으로 바꿈)
.PARAMETER OutDir    출력 폴더(기본: 저장소 루트의 out)
.EXAMPLE
  powershell -ExecutionPolicy Bypass -File tools\build_addin.ps1 -Addin image-fit
#>
param(
    [Parameter(Mandatory = $true)][string]$Addin,
    [switch]$RunTests,
    [string]$OutDir
)
$ErrorActionPreference = 'Stop'
try { [Console]::OutputEncoding = [Text.Encoding]::UTF8 } catch {}

function Fail([string]$msg, [int]$code = 1) {
    Write-Host ""
    Write-Host "[중단] $msg" -ForegroundColor Red
    $script:exitCode = $code
    throw 'BUILD_ABORT'
}
function Info([string]$msg) { Write-Host $msg }

$script:exitCode = 0
$repo = Split-Path -Parent $PSScriptRoot
$addinDir = Join-Path $repo "addins\$Addin"
$commonDir = Join-Path $repo 'addins\_common'
$clientDir = Join-Path $repo 'addins\_client'
$isLib = ($Addin -eq '_common')
$outId = if ($isLib) { 'clockout-lib' } else { $Addin }
if (-not $isLib -and $Addin -notmatch '^[a-z0-9-]+$') { Write-Host "[중단] 추가기능 id는 영문 소문자·숫자·하이픈만 쓸 수 있습니다: $Addin" -ForegroundColor Red; exit 1 }
if (-not (Test-Path $addinDir)) { Write-Host "[중단] 폴더가 없습니다: $addinDir" -ForegroundColor Red; exit 1 }
if (-not $OutDir) { $OutDir = Join-Path $repo 'out' }

# --- CP949 인코더(Windows PowerShell 5.1과 PowerShell 7 모두) ---
try { [Text.Encoding]::RegisterProvider([Text.CodePagesEncodingProvider]::Instance) } catch {}
$cp949 = [Text.Encoding]::GetEncoding(949, (New-Object Text.EncoderExceptionFallback), (New-Object Text.DecoderExceptionFallback))

# --- 소스 모으기 ---
# 공용 lib를 따로 설치해 쓰는 추가기능(선언의 dependencies에 clockout-lib이 있고 bundled가 아님)은 lib 소스 대신
# addins\_client\CoLibClient.bas(얇은 연결부)를 넣는다. 그렇지 않으면 예전처럼 addins\_common 소스를 함께 컴파일한다.
$srcExt = '.bas', '.cls', '.frm'
$declFile = Get-ChildItem $addinDir -File -Filter '*.clockout.json' | Select-Object -First 1
$useClient = $false
if ($declFile -and -not $isLib) {
    try {
        $declObj = [IO.File]::ReadAllText($declFile.FullName, [Text.Encoding]::UTF8) | ConvertFrom-Json
        foreach ($dep in @($declObj.dependencies)) { if ($dep.id -eq 'clockout-lib' -and -not $dep.bundled) { $useClient = $true } }
    } catch { Write-Host "[중단] 선언 파일을 읽을 수 없습니다: $($declFile.Name) ($($_.Exception.Message))" -ForegroundColor Red; exit 1 }
}
# 선언에 commands가 있으면 리본 버튼을 만든다(칼퇴 탭). 클릭 처리 모듈 CoRibbon.bas를 함께 넣고, 저장 뒤 .xlam에 리본 XML을 끼운다.
$cmds = @()
if ($declObj -and $declObj.commands) { $cmds = @($declObj.commands) }
$hasCommands = ($cmds.Count -gt 0 -and -not $isLib)
if ($hasCommands) {
    foreach ($c in $cmds) {
        if ($c.id -notmatch '^[A-Za-z][A-Za-z0-9_]*$') { Write-Host "[중단] commands의 id는 영문으로 시작하는 영문·숫자·밑줄이어야 합니다: $($c.id)" -ForegroundColor Red; exit 1 }
        if (-not $c.label) { Write-Host "[중단] commands '$($c.id)'에 label이 없습니다." -ForegroundColor Red; exit 1 }
        if ($c.macro -notmatch '^[A-Za-z][A-Za-z0-9_]*$') { Write-Host "[중단] commands '$($c.id)'의 macro 이름이 올바르지 않습니다: $($c.macro)" -ForegroundColor Red; exit 1 }
    }
    if (($cmds | Group-Object id | Where-Object { $_.Count -gt 1 })) { Write-Host "[중단] commands의 id가 겹칩니다." -ForegroundColor Red; exit 1 }
}
$srcFiles = @()
$srcFiles += Get-ChildItem $addinDir -File | Where-Object { $_.Extension -in $srcExt }
if ($hasCommands) { $srcFiles += Get-ChildItem (Join-Path $clientDir 'ribbon') -File -Filter 'CoRibbon.bas' }
if (-not $isLib) {
    if ($useClient) { if (Test-Path $clientDir) { $srcFiles += Get-ChildItem $clientDir -File | Where-Object { $_.Extension -in $srcExt } } }
    elseif (Test-Path $commonDir) { $srcFiles += Get-ChildItem $commonDir -File | Where-Object { $_.Extension -in $srcExt } }
}
$dups = $srcFiles | Group-Object Name | Where-Object { $_.Count -gt 1 }
if ($dups) { Write-Host "[중단] 같은 이름의 모듈 파일이 여러 폴더에 있습니다: $(($dups | ForEach-Object Name) -join ', ')" -ForegroundColor Red; exit 1 }
if (-not $srcFiles) { Write-Host "[중단] 가져올 .bas 파일이 없습니다." -ForegroundColor Red; exit 1 }
if ($useClient) { Write-Host "공용 lib는 따로 설치해 쓰는 방식입니다(CoLibClient 연결부를 넣습니다)." }
$twFile = Get-ChildItem $addinDir -File -Filter 'ThisWorkbook_*.txt' | Select-Object -First 1
$readme = Join-Path $addinDir 'README.md'

$tmp = Join-Path ([IO.Path]::GetTempPath()) ("clockout_build_" + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $tmp | Out-Null
$xl = $null; $xlPid = 0; $watch = $null
try {
    # --- CP949+CRLF 임시 변환 ---
    $encDir = Join-Path $tmp 'cp949'; New-Item -ItemType Directory -Path $encDir | Out-Null
    $modules = @()
    foreach ($f in $srcFiles) {
        $text = [IO.File]::ReadAllText($f.FullName, [Text.Encoding]::UTF8)
        if ($f.Name -eq 'CoRibbon.bas') { $text = $text.Replace('CoRibbon_Click', 'CoRibbon_Click_' + ($outId -replace '[^A-Za-z0-9]', '_')) }
        $text = ($text -replace "`r`n", "`n") -replace "`n", "`r`n"
        try { $bytes = $cp949.GetBytes($text) }
        catch { Fail "$($f.Name): CP949로 쓸 수 없는 글자가 있습니다. (이모지·특수문자 등) 해당 글자를 고쳐 주세요." }
        $outPath = Join-Path $encDir $f.Name
        [IO.File]::WriteAllBytes($outPath, $bytes)
        $frx = [IO.Path]::ChangeExtension($f.FullName, '.frx')
        if ($f.Extension -eq '.frm' -and (Test-Path $frx)) { Copy-Item $frx (Join-Path $encDir ([IO.Path]::GetFileName($frx))) }
        $nm = [IO.Path]::GetFileNameWithoutExtension($f.Name)
        if ($text -match '(?m)^Attribute VB_Name = "([^"]+)"') { $nm = $Matches[1] }
        $modules += [pscustomobject]@{ Name = $nm; Path = $outPath; Src = $f.FullName; Text = $text }
    }
    Info "소스 $($modules.Count)개를 CP949로 변환했습니다."
    if ($hasCommands) {
        foreach ($c in $cmds) {
            $found = $false
            foreach ($m in $modules) { if ($m.Text -match ('(?mi)^\s*Public Sub ' + [regex]::Escape($c.macro) + '\s*\(')) { $found = $true; break } }
            if (-not $found) { Fail "commands '$($c.id)'의 매크로 Public Sub $($c.macro)를 소스에서 찾지 못했습니다." }
        }
    }

    # --- Excel 시작(보이지 않게) ---
    try { $xl = New-Object -ComObject Excel.Application }
    catch { Fail "Windows용 Excel(데스크톱판)을 실행할 수 없습니다. Excel이 설치된 PC에서만 컴파일할 수 있습니다." 2 }
    $xl.Visible = $false; $xl.DisplayAlerts = $false
    try {
        Add-Type -Namespace ClockOutBuild -Name Win -MemberDefinition '[System.Runtime.InteropServices.DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint pid);' -ErrorAction SilentlyContinue
        $p = [uint32]0
        [void][ClockOutBuild.Win]::GetWindowThreadProcessId([IntPtr]$xl.Hwnd, [ref]$p)
        $xlPid = [int]$p
    } catch {}
    # 멈췄을 때를 대비한 안전장치: 이 스크립트가 띄운 Excel만 4분 뒤 강제 종료
    if ($xlPid -gt 0) { $watch = Start-Job -ArgumentList $xlPid -ScriptBlock { param($p) Start-Sleep 240; Stop-Process -Id $p -Force -ErrorAction SilentlyContinue } }

    function New-Book($xlApp) {
        $b = $xlApp.Workbooks.Add()
        $ok = $false
        try { $v = $b.VBProject; if ($v -ne $null) { $null = $v.VBComponents.Count; $ok = $true } } catch {}
        if (-not $ok) {
            Write-Host ""
            Write-Host "[중단] Excel이 VBA 프로젝트 접근을 막고 있습니다." -ForegroundColor Red
            Write-Host "  Excel > 파일 > 옵션 > 보안 센터 > 보안 센터 설정 > 매크로 설정에서" -ForegroundColor Yellow
            Write-Host "  'VBA 프로젝트 개체 모델에 대한 액세스 신뢰'를 체크하고 확인을 누른 뒤 다시 실행하세요." -ForegroundColor Yellow
            Write-Host "  (이 설정은 Windows 계정별로 저장됩니다. 이 스크립트는 레지스트리를 바꾸지 않습니다.)" -ForegroundColor Yellow
            $script:exitCode = 3
            throw 'BUILD_ABORT'
        }
        return $b
    }

    # .frm은 가져오기 대신 폼을 새로 만들어 코드를 넣는다(.frx 없는 코드 전용 폼은 Import가 실패한다).
    function Import-Source($project, [string]$path) {
        if ([IO.Path]::GetExtension($path) -ne '.frm') { $null = $project.VBComponents.Import($path); return }
        $ftext = $cp949.GetString([IO.File]::ReadAllBytes($path))
        $fname = [IO.Path]::GetFileNameWithoutExtension($path)
        if ($ftext -match '(?m)^Attribute VB_Name = "([^"]+)"') { $fname = $Matches[1] }
        $fcap = $fname
        if ($ftext -match '(?m)^\s*Caption\s*=\s*"([^"]*)"') { $fcap = $Matches[1] }
        $at = $ftext.LastIndexOf('Attribute VB_Exposed')
        if ($at -lt 0) { throw "폼 파일에 Attribute 줄이 없습니다: $path" }
        $eol = $ftext.IndexOf("`n", $at)
        $fcode = $ftext.Substring($eol + 1)
        $comp = $project.VBComponents.Add(3)
        $comp.Name = $fname
        try { $comp.Properties.Item('Caption').Value = $fcap } catch {}
        $cm = $comp.CodeModule
        if ($cm.CountOfLines -gt 0 -and $cm.Lines(1, $cm.CountOfLines) -match 'Option Explicit') { $fcode = $fcode -replace '(?m)^Option Explicit\r?\n', '' }
        $cm.AddFromString($fcode)
    }

    # --- 1) 빌드 통합 문서: 가져오기 + 컴파일 ---
    $wb = New-Book $xl
    $vp = $wb.VBProject
    try { $vp.Name = ('ClockOut_' + ($outId -replace '[^A-Za-z0-9]', '_')) } catch {}
    foreach ($m in $modules) {
        try { Import-Source $vp $m.Path }
        catch { Fail "가져오기 실패: $($m.Src)`n  $($_.Exception.Message)" }
    }
    if ($twFile) {
        $twCode = [IO.File]::ReadAllText($twFile.FullName, [Text.Encoding]::UTF8)
        $vp.VBComponents.Item($wb.CodeName).CodeModule.AddFromString($twCode)
    }
    Info "Excel에 가져왔습니다. 컴파일합니다..."
    $btn = $xl.VBE.CommandBars.FindControl(1, 578)
    if ($btn -eq $null) { Fail "VBE의 'VBAProject 컴파일' 명령을 찾지 못했습니다. (Excel 버전 문제일 수 있음)" }
    $xl.VBE.ActiveVBProject = $vp
    $btn.Execute()
    if ($btn.Enabled) {
        $where = ''
        try {
            $cp = $xl.VBE.ActiveCodePane
            if ($cp -ne $null) {
                $sl = 0; $sc = 0; $el = 0; $ec = 0
                $cp.GetSelection([ref]$sl, [ref]$sc, [ref]$el, [ref]$ec)
                $modName = $cp.CodeModule.Parent.Name
                $line = $cp.CodeModule.Lines($sl, 1).Trim()
                $where = "`n  오류 위치(추정): 모듈 '$modName' ${sl}번째 줄: $line"
            }
        } catch {}
        Fail "컴파일 오류가 있습니다.$where`n  (Excel에서 같은 파일을 가져와 '디버그 > VBAProject 컴파일'로 위치를 확인할 수 있습니다.)" 4
    }
    Info "컴파일 성공."

    # --- 2) 시험(선택): 별도 통합 문서, MsgBox는 기록으로 바꾼 사본 ---
    if ($RunTests) {
        $tests = @()
        foreach ($m in $modules) {
            foreach ($mt in [regex]::Matches($m.Text, '(?m)^\s*Public Sub (Test_\w+)\s*\(\s*\)')) {
                $tests += [pscustomobject]@{ Mod = $m.Name; Proc = $mt.Groups[1].Value }
            }
        }
        if (-not $tests) { Info "시험 함수(Test_)가 없어 시험을 건너뜁니다." }
        else {
            $log = Join-Path $tmp 'testlog.txt'
            $tDir = Join-Path $tmp 'testcopy'; New-Item -ItemType Directory -Path $tDir | Out-Null
            $tpaths = @()
            foreach ($m in $modules) {
                $t = $m.Text
                if ($t -match '(?m)Public Sub Test_') {
                    $t = [regex]::Replace($t, '(?m)^([ 	]*)MsgBox ([^
]*), vb\w+(?:[ \t]*,[ \t]*[^\r\n]+)?[ \t]*(?=\r?$)', '${1}ZLog ${2}')
                    $t += "`r`nPrivate Sub ZLog(ByVal t As String)`r`n    Dim n As Integer: n = FreeFile`r`n    Open `"$log`" For Append As #n`r`n    Print #n, t`r`n    Close #n`r`nEnd Sub`r`n"
                    if ($t -match '(?m)^\s*MsgBox ') { Fail "시험 함수가 있는 모듈 '$($m.Name)'에 단순 MsgBox 형식이 아닌 호출이 있어 창이 뜰 수 있습니다. -RunTests 없이 실행하세요." }
                }
                $tp = Join-Path $tDir ([IO.Path]::GetFileName($m.Path))
                $frxSrc = [IO.Path]::ChangeExtension($m.Path, '.frx'); if (Test-Path $frxSrc) { Copy-Item $frxSrc (Join-Path $tDir ([IO.Path]::GetFileName($frxSrc))) }
                try { [IO.File]::WriteAllBytes($tp, $cp949.GetBytes($t)) } catch { Fail "시험용 사본 변환 실패: $($m.Name) (임시 폴더 경로에 CP949로 쓸 수 없는 글자가 있을 수 있음)" }
                $tpaths += $tp
            }
            $tb = New-Book $xl
            foreach ($tp in $tpaths) { Import-Source $tb.VBProject $tp }
            $failed = 0
            foreach ($t in $tests) {
                Remove-Item $log -ErrorAction SilentlyContinue
                try { $null = $xl.Run("$($t.Mod).$($t.Proc)") }
                catch { Write-Host "  [실패] $($t.Mod).$($t.Proc): $($_.Exception.Message)" -ForegroundColor Red; $failed++; continue }
                $res = ''
                if (Test-Path $log) { $res = [IO.File]::ReadAllText($log, [Text.Encoding]::GetEncoding(949)) }
                if ($res -match '모두 통과') { Info "  [통과] $($t.Mod).$($t.Proc)" }
                else { Write-Host "  [실패] $($t.Mod).$($t.Proc): $($res.Trim())" -ForegroundColor Red; $failed++ }
            }
            $tb.Close($false)
            if ($failed -gt 0) { Fail "시험 $failed 개가 실패했습니다. (xlam/zip은 만들지 않았습니다)" 5 }
        }
    }

    # --- 3) 저장 + zip ---
    New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
    $xlam = Join-Path $OutDir "$outId.xlam"
    $zip = Join-Path $OutDir "$outId.zip"
    Remove-Item $xlam, $zip -ErrorAction SilentlyContinue
    $wb.SaveAs($xlam, 55)
    $wb.Close($false)
    Info "저장: $xlam"

    if ($hasCommands) {
        Add-Type -AssemblyName System.IO.Compression
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        $esc = { param($t) [Security.SecurityElement]::Escape([string]$t) }
        # 추가기능마다 그룹 하나(id가 추가기능별로 달라 겹치지 않음). main:true 명령은 큰 단추, 나머지는 작은 단추(3줄씩 쌓임), 작은 단추가 6개를 넘으면 나머지는 "더 보기" 드롭다운으로.
        $aid = ($outId -replace '[^A-Za-z0-9]', '_')
        $gname = if ($declObj.group) { [string]$declObj.group } elseif ($declObj.name) { [string]$declObj.name } else { $outId }
        $aicon = if ($declObj.icon) { [string]$declObj.icon } else { 'MacroPlay' }
        $okIcon = { param($list)
            foreach ($n in ([string]$list -split ',')) { $n = $n.Trim(); if (-not $n) { continue }
                # 유효한 이름도 보이지 않는 Excel에서는 그림 변환이 E_UNEXPECTED(0x8000FFFF)로 실패한다.
                # 없는 이름만 E_INVALIDARG(0x80070057)를 던지므로 그것만 무효로 본다.
                try { $null = $xl.CommandBars.GetImageMso($n, 16, 16); return $n }
                catch { if ($_.Exception.HResult -ne -2147024809) { return $n } } }
            return $null }
        $badIcons = @()
        $mk = { param($c, $large, $ind)
            $want = if ($c.icon) { [string]$c.icon } else { $aicon }
            $icon = & $okIcon $want
            if (-not $icon) { $script:badIcons += $want; $icon = & $okIcon $aicon; if (-not $icon) { $icon = 'MacroPlay' } }
            $tip = if ($c.tip) { " screentip=`"$(& $esc $c.label)`" supertip=`"$(& $esc $c.tip)`"" } else { '' }
            $sz = if ($large) { ' size="large"' } else { '' }
            "$ind<button idQ=`"ck:btn_${aid}_$($c.id)`" label=`"$(& $esc $c.label)`" imageMso=`"$(& $esc $icon)`"$sz tag=`"$(& $esc $c.macro)`" onAction=`"CoRibbon_Click_$aid`"$tip />`n" }
        $larges = @($cmds | Where-Object { $_.main -eq $true })
        $smalls = @($cmds | Where-Object { $_.main -ne $true })
        if ($larges.Count -eq 0 -and $smalls.Count -le 3) { $larges = $smalls; $smalls = @() }
        $inner = New-Object Text.StringBuilder
        foreach ($c in $larges) { [void]$inner.Append((& $mk $c $true '      ')) }
        $shown = if ($smalls.Count -gt 6) { $smalls[0..4] } else { $smalls }
        foreach ($c in $shown) { [void]$inner.Append((& $mk $c $false '      ')) }
        if ($smalls.Count -gt 6) {
            $mi = & $okIcon 'MacroPlay'
            [void]$inner.Append("      <menu idQ=`"ck:mnu_$aid`" label=`"더 보기`" imageMso=`"$mi`">`n")
            foreach ($c in $smalls[5..($smalls.Count - 1)]) { [void]$inner.Append((& $mk $c $false '        ')) }
            [void]$inner.Append("      </menu>`n")
        }
        if ($badIcons.Count) { Write-Host "[알림] 쓸 수 없는 아이콘 이름이라 기본 아이콘으로 바꿨습니다: $(($badIcons | Select-Object -Unique) -join ', ')" -ForegroundColor Yellow }
        $ui = "<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?>`n" +
              "<customUI xmlns=`"http://schemas.microsoft.com/office/2009/07/customui`" xmlns:ck=`"urn:clockout:ribbon`">`n" +
              "  <ribbon>`n    <tabs>`n" +
              "    <tab idQ=`"ck:main`" label=`"칼퇴`" insertAfterMso=`"TabView`">`n" +
              "     <group idQ=`"ck:grp_$aid`" label=`"$(& $esc $gname)`">`n" + $inner.ToString() + "     </group>`n    </tab>`n    </tabs>`n  </ribbon>`n</customUI>`n"
        $utf8 = New-Object Text.UTF8Encoding($false)
        $za = [IO.Compression.ZipFile]::Open($xlam, 'Update')
        try {
            $old1 = $za.GetEntry('customUI/customUI14.xml'); if ($old1) { $old1.Delete() }
            $en = $za.CreateEntry('customUI/customUI14.xml')
            $sw = New-Object IO.StreamWriter($en.Open(), $utf8); $sw.Write($ui); $sw.Dispose()
            $re = $za.GetEntry('_rels/.rels')
            $sr = New-Object IO.StreamReader($re.Open(), $utf8); $rels = $sr.ReadToEnd(); $sr.Dispose()
            if ($rels -notmatch 'customUI14') {
                $rels = $rels.Replace('</Relationships>', '<Relationship Id="rIdCkUi" Type="http://schemas.microsoft.com/office/2007/relationships/ui/extensibility" Target="customUI/customUI14.xml"/></Relationships>')
                $re.Delete()
                $re2 = $za.CreateEntry('_rels/.rels')
                $sw = New-Object IO.StreamWriter($re2.Open(), $utf8); $sw.Write($rels); $sw.Dispose()
            }
        } finally { $za.Dispose() }
        Info "리본 버튼 $($cmds.Count)개를 '칼퇴' 탭에 넣었습니다."
    }

    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    function Add-Entry($arc, [string]$file, [string]$name) {
        [void][IO.Compression.ZipFileExtensions]::CreateEntryFromFile($arc, $file, $name, [IO.Compression.CompressionLevel]::Optimal)
    }
    $za = [IO.Compression.ZipFile]::Open($zip, 'Create')
    try {
        Add-Entry $za $xlam "$outId.xlam"
        if ($declFile) { Add-Entry $za $declFile.FullName $declFile.Name }
        if (Test-Path $readme) { Add-Entry $za $readme 'README.md' }
        foreach ($f in ($srcFiles | Sort-Object Name)) {
            Add-Entry $za $f.FullName ("src/" + $f.Name)
            $frxZ = [IO.Path]::ChangeExtension($f.FullName, '.frx')
            if ($f.Extension -eq '.frm' -and (Test-Path $frxZ)) { Add-Entry $za $frxZ ("src/" + [IO.Path]::GetFileName($frxZ)) }
        }
        if ($twFile) { Add-Entry $za $twFile.FullName ("src/" + $twFile.Name) }
    } finally { $za.Dispose() }
    Info "저장: $zip"
    $za = [IO.Compression.ZipFile]::OpenRead($zip)
    try { Info "zip 내용:"; foreach ($e in $za.Entries) { Info "  $($e.FullName)" } } finally { $za.Dispose() }
    Info ""
    Info "완료."
}
catch {
    if ($_.Exception.Message -ne 'BUILD_ABORT') {
        Write-Host ""
        Write-Host "[오류] $($_.Exception.Message)" -ForegroundColor Red
        if ($script:exitCode -eq 0) { $script:exitCode = 1 }
    }
}
finally {
    if ($xl -ne $null) {
        try { foreach ($b in @($xl.Workbooks)) { try { $b.Saved = $true; $b.Close($false) } catch {} } } catch {}
        try { $xl.Quit() } catch {}
        try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) } catch {}
    }
    if ($watch) { Stop-Job $watch -ErrorAction SilentlyContinue; Remove-Job $watch -Force -ErrorAction SilentlyContinue }
    if ($xlPid -gt 0) {
        # 이 스크립트가 띄운 Excel 프로세스만 정리(사용자 Excel은 건드리지 않음)
        for ($i = 0; $i -lt 20; $i++) { if (-not (Get-Process -Id $xlPid -ErrorAction SilentlyContinue)) { break }; Start-Sleep -Milliseconds 250 }
        Stop-Process -Id $xlPid -Force -ErrorAction SilentlyContinue
    }
    Remove-Item $tmp -Recurse -Force -ErrorAction SilentlyContinue
}
exit $script:exitCode
