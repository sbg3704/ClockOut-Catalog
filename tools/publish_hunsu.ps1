<#
.SYNOPSIS
  [초안] 하위 프로그램(훈수꾼) 틀 zip을 GitHub Release로 올리고 catalog.json의 apps 구역에 등록한다.
.DESCRIPTION
  상태: 초안이다. 실제 Release 생성은 사용자가 PC에서 직접 한다(클라우드 세션은 공개 Release 생성 권한 검사에 막힌다).
        -DryRun으로 해시와 등록될 항목만 확인할 수 있고, 그때는 GitHub에 아무것도 올리지 않으며 파일도 바꾸지 않는다.
        이 스크립트는 아직 PowerShell로 실행해 본 적이 없다(클라우드에 PowerShell이 없음). 처음 돌릴 때 -DryRun부터 쓴다.
  1) -Zip 으로 받은 틀 zip의 SHA-256·크기를 계산한다. 틀 zip 만들기는 훈수꾼 저장소의 빌드 일이다(이 스크립트는 만들지 않는다).
  2) -Meta 의 JSON(id, name, version, icon, summary, description, preview, requires, host)에서 항목 내용을 읽는다.
  3) Release(태그 app-<id>-<판>)를 만들고 zip을 올린다(이미 있으면 파일 덮어쓰기) -- PC에 설치·로그인된 GitHub CLI(gh) 사용.
  4) catalog.json 의 apps 항목을 같은 id 기준으로 바꾸거나 추가하고 updated 날짜를 갱신한다. addins·packs·libraries는 건드리지 않는다.
  토큰은 저장소·스크립트에 넣지 않는다(gh auth login 으로 로그인한 계정 권한 사용). 커밋·푸시는 하지 않는다: 끝난 뒤 catalog.json 변경을 PR로 올린다.
.PARAMETER Zip     올릴 틀 zip 경로 (예: C:\work\hunsu-0.1.000.zip)
.PARAMETER Meta    항목 내용 JSON 경로 (예: ClockOut-HunSu 저장소의 app-meta.json)
.PARAMETER DryRun  해시·태그·등록될 항목만 보여 주고 아무것도 올리거나 바꾸지 않는다
.EXAMPLE
  powershell -ExecutionPolicy Bypass -File tools\publish_hunsu.ps1 -Zip C:\work\hunsu-0.1.000.zip -Meta C:\work\app-meta.json -DryRun
#>
param(
    [Parameter(Mandatory = $true)][string]$Zip,
    [Parameter(Mandatory = $true)][string]$Meta,
    [switch]$DryRun
)
$ErrorActionPreference = 'Stop'
try { [Console]::OutputEncoding = [Text.Encoding]::UTF8 } catch {}
$repoSlug = 'sbg3704/ClockOut-Catalog'
$repo = Split-Path -Parent $PSScriptRoot
$maxBytes = 2GB   # CatalogAppParser.MaxAppSizeBytes 와 같은 값(용량은 걱정하지 않기로 함, 안전 상한)

function Stop-With([string]$msg, [int]$code = 1) { Write-Host "[중단] $msg" -ForegroundColor Red; exit $code }

if (-not (Test-Path $Zip)) { Stop-With "zip이 없습니다: $Zip" }
if (-not (Test-Path $Meta)) { Stop-With "항목 내용 파일이 없습니다: $Meta" }
$m = [IO.File]::ReadAllText((Resolve-Path $Meta), [Text.Encoding]::UTF8) | ConvertFrom-Json
foreach ($k in 'id', 'name', 'version') { if (-not $m.$k) { Stop-With "항목 내용에 $k 가 없습니다." } }
if ($m.id -notmatch '^[A-Za-z0-9][A-Za-z0-9.-]*$' -or $m.id -like '*..*') { Stop-With "id 모양이 올바르지 않습니다(영문·숫자·하이픈·점)." }
if ($m.version -notmatch '^[A-Za-z0-9][A-Za-z0-9.-]*$' -or $m.version -like '*..*') { Stop-With "version 모양이 올바르지 않습니다." }

$sha = (Get-FileHash $Zip -Algorithm SHA256).Hash.ToLower()
$size = (Get-Item $Zip).Length
if ($size -gt $maxBytes) { Stop-With "zip이 2GB를 넘습니다($size 바이트). 앱이 이 항목을 건너뜁니다." }
$tag = "app-$($m.id)-$($m.version)"
$assetName = "$($m.id).zip"
$url = "https://github.com/$repoSlug/releases/download/$tag/$assetName"

$entry = [ordered]@{ id = [string]$m.id; name = [string]$m.name; version = [string]$m.version }
foreach ($k in 'icon', 'summary', 'description', 'preview') { if ($m.$k) { $entry[$k] = [string]$m.$k } }
$entry.url = $url; $entry.sha256 = $sha; $entry.size = $size
if ($m.requires) { $entry.requires = $m.requires }
if ($m.host) { $entry.host = [string]$m.host }
Write-Host "태그: $tag"; Write-Host "SHA-256: $sha"; Write-Host "크기: $size"; Write-Host "주소: $url"

if ($DryRun) {
    Write-Host "등록될 항목:"; Write-Host ($entry | ConvertTo-Json -Depth 5)
    Write-Host "[DryRun] GitHub에 올리지 않고 catalog.json도 바꾸지 않았습니다." -ForegroundColor Yellow; exit 0
}
if (-not (Get-Command gh -ErrorAction SilentlyContinue)) { Stop-With "gh(GitHub CLI)가 없습니다. winget install --id GitHub.cli -e 후 gh auth login" 2 }
$ErrorActionPreference = 'Continue'   # gh가 stderr에 쓰는 것을 오류로 취급하지 않게(PowerShell 5.1)
& gh auth status 2>&1 | Out-Null
if ($LASTEXITCODE -ne 0) { Stop-With "gh 로그인이 필요합니다: gh auth login" 3 }
# Release 파일 이름은 항상 <id>.zip 이어야 위 주소와 맞는다. 다른 이름이면 임시로 복사해 올린다.
$upload = $Zip
if ((Split-Path -Leaf $Zip) -ne $assetName) { $upload = Join-Path ([IO.Path]::GetTempPath()) $assetName; Copy-Item $Zip $upload -Force }
& gh release view $tag --repo $repoSlug 2>&1 | Out-Null
if ($LASTEXITCODE -eq 0) {
    & gh release upload $tag $upload --clobber --repo $repoSlug
} else {
    & gh release create $tag $upload --repo $repoSlug --title "$($m.name) $($m.version)" --notes ([string]$m.summary)
}
if ($LASTEXITCODE -ne 0) { Stop-With "Release 업로드에 실패했습니다." 4 }
$ErrorActionPreference = 'Stop'

# catalog.json 갱신(UTF-8, BOM 없음). apps 칸이 없으면 만든다.
$catPath = Join-Path $repo 'catalog.json'
$cat = [IO.File]::ReadAllText($catPath, [Text.Encoding]::UTF8) | ConvertFrom-Json
$apps = @(); if ($cat.PSObject.Properties.Name -contains 'apps' -and $cat.apps) { $apps = @($cat.apps | Where-Object { $_.id -ne $entry.id }) }
$apps = @($apps + [pscustomobject]$entry | Sort-Object id)
if ($cat.PSObject.Properties.Name -contains 'apps') { $cat.apps = $apps } else { $cat | Add-Member -NotePropertyName apps -NotePropertyValue $apps }
$cat.updated = (Get-Date).ToString('yyyy-MM-dd')
$json = ($cat | ConvertTo-Json -Depth 8)
[IO.File]::WriteAllText($catPath, ($json -replace "`r`n", "`n") + "`n", (New-Object Text.UTF8Encoding($false)))
Write-Host "완료: Release 업로드와 catalog.json 갱신. catalog.json 변경을 PR로 올려 주세요." -ForegroundColor Green
