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
$isLib = ($Addin -eq '_common')
$outId = if ($isLib) { 'clockout-lib' } else { $Addin }
if (-not $isLib -and $Addin -notmatch '^[a-z0-9-]+$') { Write-Host "[중단] 추가기능 id는 영문 소문자·숫자·하이픈만 쓸 수 있습니다: $Addin" -ForegroundColor Red; exit 1 }
if (-not (Test-Path $addinDir)) { Write-Host "[중단] 폴더가 없습니다: $addinDir" -ForegroundColor Red; exit 1 }
if (-not $OutDir) { $OutDir = Join-Path $repo 'out' }

# --- CP949 인코더(Windows PowerShell 5.1과 PowerShell 7 모두) ---
try { [Text.Encoding]::RegisterProvider([Text.CodePagesEncodingProvider]::Instance) } catch {}
$cp949 = [Text.Encoding]::GetEncoding(949, (New-Object Text.EncoderExceptionFallback), (New-Object Text.DecoderExceptionFallback))

# --- 소스 모으기 ---
$srcFiles = @()
$srcFiles += Get-ChildItem $addinDir -File | Where-Object { $_.Extension -in '.bas', '.cls' }
if (-not $isLib -and (Test-Path $commonDir)) { $srcFiles += Get-ChildItem $commonDir -File | Where-Object { $_.Extension -in '.bas', '.cls' } }
$dups = $srcFiles | Group-Object Name | Where-Object { $_.Count -gt 1 }
if ($dups) { Write-Host "[중단] 같은 이름의 모듈 파일이 addins\$Addin 와 addins\_common 에 모두 있습니다: $(($dups | ForEach-Object Name) -join ', ')" -ForegroundColor Red; exit 1 }
if (-not $srcFiles) { Write-Host "[중단] 가져올 .bas 파일이 없습니다." -ForegroundColor Red; exit 1 }
$twFile = Get-ChildItem $addinDir -File -Filter 'ThisWorkbook_*.txt' | Select-Object -First 1
$declFile = Get-ChildItem $addinDir -File -Filter '*.clockout.json' | Select-Object -First 1
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
        $text = ($text -replace "`r`n", "`n") -replace "`n", "`r`n"
        try { $bytes = $cp949.GetBytes($text) }
        catch { Fail "$($f.Name): CP949로 쓸 수 없는 글자가 있습니다. (이모지·특수문자 등) 해당 글자를 고쳐 주세요." }
        $outPath = Join-Path $encDir $f.Name
        [IO.File]::WriteAllBytes($outPath, $bytes)
        $nm = [IO.Path]::GetFileNameWithoutExtension($f.Name)
        if ($text -match '(?m)^Attribute VB_Name = "([^"]+)"') { $nm = $Matches[1] }
        $modules += [pscustomobject]@{ Name = $nm; Path = $outPath; Src = $f.FullName; Text = $text }
    }
    Info "소스 $($modules.Count)개를 CP949로 변환했습니다."

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

    # --- 1) 빌드 통합 문서: 가져오기 + 컴파일 ---
    $wb = New-Book $xl
    $vp = $wb.VBProject
    try { $vp.Name = ('ClockOut_' + ($outId -replace '[^A-Za-z0-9]', '_')) } catch {}
    foreach ($m in $modules) {
        try { $null = $vp.VBComponents.Import($m.Path) }
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
                try { [IO.File]::WriteAllBytes($tp, $cp949.GetBytes($t)) } catch { Fail "시험용 사본 변환 실패: $($m.Name) (임시 폴더 경로에 CP949로 쓸 수 없는 글자가 있을 수 있음)" }
                $tpaths += $tp
            }
            $tb = New-Book $xl
            foreach ($tp in $tpaths) { $null = $tb.VBProject.VBComponents.Import($tp) }
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
        foreach ($f in ($srcFiles | Sort-Object Name)) { Add-Entry $za $f.FullName ("src/" + $f.Name) }
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
