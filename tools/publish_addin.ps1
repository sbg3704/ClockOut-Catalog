<#
.SYNOPSIS
  추가기능을 빌드(zip)한 뒤 GitHub Release에 올리고 catalog.json에 SHA-256과 함께 등록한다.
.DESCRIPTION
  1) tools\build_addin.ps1 -Addin <id> 로 out\<id>.zip 만들기(-SkipBuild면 건너뛰고 있는 zip 사용)
  2) zip의 SHA-256·크기 계산, addins\<id>\<id>.clockout.json 에서 name/version/author/description/host 읽기
  3) Release(태그 <id>-<version>)를 만들고 zip 올리기(이미 있으면 파일 덮어쓰기)  -- PC에 설치·로그인된 GitHub CLI(gh) 사용
  4) catalog.json 의 addins 항목을 같은 id 기준으로 바꾸거나 추가하고 updated 날짜 갱신
  토큰은 저장소·스크립트에 넣지 않는다(gh auth login 으로 로그인한 계정 권한 사용). 커밋·푸시는 하지 않는다: 끝난 뒤 catalog.json 변경을 PR로 올린다.
.PARAMETER Addin     addins\ 아래 폴더 이름(= 추가기능 id). `_common`은 clockout-lib 로 올린다
.PARAMETER SkipBuild 이미 만든 out\<id>.zip 사용
.PARAMETER DryRun    빌드·해시·catalog.json 갱신 내용만 보여 주고 GitHub에는 아무것도 올리지 않으며 파일도 바꾸지 않는다
.EXAMPLE
  powershell -ExecutionPolicy Bypass -File tools\publish_addin.ps1 -Addin image-fit
#>
param(
    [Parameter(Mandatory = $true)][string]$Addin,
    [switch]$SkipBuild,
    [switch]$DryRun
)
$ErrorActionPreference = 'Stop'
try { [Console]::OutputEncoding = [Text.Encoding]::UTF8 } catch {}
$repoSlug = 'sbg3704/ClockOut-Catalog'
$repo = Split-Path -Parent $PSScriptRoot
$isLib = ($Addin -eq '_common')
$outId = if ($isLib) { 'clockout-lib' } else { $Addin }
$zip = Join-Path $repo "out\$outId.zip"

function Stop-With([string]$msg, [int]$code = 1) { Write-Host "[중단] $msg" -ForegroundColor Red; exit $code }

# 1) 빌드
if (-not $SkipBuild) {
    & powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'build_addin.ps1') -Addin $Addin
    if ($LASTEXITCODE -ne 0) { Stop-With "빌드가 실패해 올리지 않습니다(build_addin.ps1 종료 코드 $LASTEXITCODE)." $LASTEXITCODE }
}
if (-not (Test-Path $zip)) { Stop-With "zip이 없습니다: $zip" }

# 2) 해시와 선언
$sha = (Get-FileHash $zip -Algorithm SHA256).Hash.ToLower()
$size = (Get-Item $zip).Length
$declFile = Get-ChildItem (Join-Path $repo "addins\$Addin") -Filter '*.clockout.json' -File | Select-Object -First 1
if (-not $declFile) { Stop-With "선언 파일(*.clockout.json)이 없습니다." }
$decl = [IO.File]::ReadAllText($declFile.FullName, [Text.Encoding]::UTF8) | ConvertFrom-Json
if (-not $decl.version) { Stop-With "선언 파일에 version이 없습니다." }
$tag = "$outId-$($decl.version)"
$url = "https://github.com/$repoSlug/releases/download/$tag/$outId.zip"
$entry = [ordered]@{
    id = $outId; name = [string]$decl.name; version = [string]$decl.version
    author = [string]$decl.author; host = [string]$decl.host; description = [string]$decl.description
    url = $url; sha256 = $sha; size = $size
}
Write-Host "태그: $tag"; Write-Host "SHA-256: $sha"; Write-Host "주소: $url"

# 3) Release 업로드
if ($DryRun) { Write-Host "[DryRun] GitHub에 올리지 않고 catalog.json도 바꾸지 않았습니다." -ForegroundColor Yellow; exit 0 }
if (-not (Get-Command gh -ErrorAction SilentlyContinue)) { Stop-With "gh(GitHub CLI)가 없습니다. winget install --id GitHub.cli -e 후 gh auth login" 2 }
& gh auth status *> $null
if ($LASTEXITCODE -ne 0) { Stop-With "gh 로그인이 필요합니다: gh auth login" 3 }
& gh release view $tag --repo $repoSlug *> $null
if ($LASTEXITCODE -eq 0) {
    & gh release upload $tag $zip --clobber --repo $repoSlug
} else {
    & gh release create $tag $zip --repo $repoSlug --title "$($entry.name) $($entry.version)" --notes ([string]$decl.description)
}
if ($LASTEXITCODE -ne 0) { Stop-With "Release 업로드에 실패했습니다." 4 }

# 4) catalog.json 갱신(UTF-8, BOM 없음)
$catPath = Join-Path $repo 'catalog.json'
$cat = [IO.File]::ReadAllText($catPath, [Text.Encoding]::UTF8) | ConvertFrom-Json
$list = @($cat.addins | Where-Object { $_.id -ne $outId }) + [pscustomobject]$entry
$cat.addins = @($list | Sort-Object id)
$cat.updated = (Get-Date).ToString('yyyy-MM-dd')
$json = ($cat | ConvertTo-Json -Depth 6)
[IO.File]::WriteAllText($catPath, ($json -replace "`r`n", "`n") + "`n", (New-Object Text.UTF8Encoding($false)))
Write-Host "완료: Release 업로드와 catalog.json 갱신. catalog.json 변경을 PR로 올려 주세요." -ForegroundColor Green
