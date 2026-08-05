$ErrorActionPreference = 'Stop'

$root = (Resolve-Path (Split-Path -Parent $MyInvocation.MyCommand.Path)).Path
Set-Location -LiteralPath $root

function Remove-IfExists([string] $Path) {
    if (Test-Path -LiteralPath $Path) {
        Remove-Item -LiteralPath $Path -Recurse -Force
    }
}

function Invoke-Checked([string] $FilePath, [string[]] $Arguments) {
    & $FilePath @Arguments
    if ($LASTEXITCODE -ne 0) {
        throw "$FilePath failed with exit code $LASTEXITCODE."
    }
}

$project = Join-Path $root 'src\EmailManage.App\EmailManage.App.csproj'
$version = (Select-Xml -LiteralPath $project -XPath '/Project/PropertyGroup/Version').Node.InnerText.Trim()
if ($version -notmatch '^\d+\.\d+\.\d+$') {
    throw "The project version '$version' is not a three-part release version."
}

$buildRoot = Join-Path $env:TEMP "MailZen-Build-$version"
$publishRoot = Join-Path $buildRoot 'publish'
$installerOutput = Join-Path $env:TEMP "MailZen-Inno-$version"
$rootAppExe = Join-Path $root "MailZen-$version.exe"
$rootInstallerExe = Join-Path $root "MailZenSetup-$version.exe"

Write-Host "MailZen version: $version" -ForegroundColor Cyan

$isccCandidates = @(
    (Get-Command ISCC.exe -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Source -First 1),
    'C:\Program Files (x86)\Inno Setup 6\ISCC.exe',
    'C:\Program Files\Inno Setup 6\ISCC.exe',
    (Join-Path $env:LOCALAPPDATA 'Programs\Inno Setup 6\ISCC.exe')
) | Where-Object { $_ -and (Test-Path -LiteralPath $_) } | Select-Object -Unique

if (-not $isccCandidates) {
    throw 'Inno Setup was not found. Install Inno Setup 6 so ISCC.exe is available, then run Build-MailZen.bat again.'
}

# Remove every old executable and generated release directory in the repository.
Get-ChildItem -LiteralPath $root -Recurse -File -Filter '*.exe' -ErrorAction SilentlyContinue |
    Where-Object { $_.FullName -notmatch '\.git\\' } |
    Remove-Item -Force
Get-ChildItem -LiteralPath $root -Directory -Filter 'publish-current-*' -ErrorAction SilentlyContinue |
    Remove-Item -Recurse -Force
Get-ChildItem -LiteralPath (Join-Path $root 'installer') -File -Filter '*.exe' -ErrorAction SilentlyContinue |
    Remove-Item -Force
Remove-IfExists $buildRoot

New-Item -ItemType Directory -Path $publishRoot -Force | Out-Null

Invoke-Checked 'dotnet' @('build', 'src\EmailManage.sln', '-c', 'Release')
Invoke-Checked 'dotnet' @('test', 'tests\EmailManage.Tests\EmailManage.Tests.csproj', '-c', 'Release')
Invoke-Checked 'dotnet' @('publish', $project, '-c', 'Release', '-r', 'win-x64', '--self-contained', 'true', '-o', $publishRoot)

$iss = Join-Path $root 'installer\MailZen.iss'
$iscc = @($isccCandidates)[0]
Remove-IfExists $installerOutput
New-Item -ItemType Directory -Path $installerOutput -Force | Out-Null
Invoke-Checked $iscc @('/Q', "/O$installerOutput", "/DMyAppVersion=$version", "/DMyAppOutputBaseFilename=MailZenSetup", "/DMyAppSourceDir=$publishRoot", $iss)

$builtInstaller = Join-Path $installerOutput 'MailZenSetup.exe'
if (-not (Test-Path -LiteralPath $builtInstaller)) {
    throw "Inno Setup completed but did not create $builtInstaller."
}
if (-not (Test-Path -LiteralPath (Join-Path $publishRoot 'MailZen.exe'))) {
    throw "Publish completed but did not create $publishRoot\MailZen.exe."
}

Copy-Item -LiteralPath (Join-Path $publishRoot 'MailZen.exe') -Destination $rootAppExe -Force
Move-Item -LiteralPath $builtInstaller -Destination $rootInstallerExe -Force

# Remove all temporary and compiler output directories so no nested executable remains.
Remove-IfExists $buildRoot
Get-ChildItem -LiteralPath $root -Recurse -Directory -ErrorAction SilentlyContinue |
    Where-Object { $_.FullName -notmatch '\.git\\' -and $_.Name -in @('bin', 'obj') } |
    Sort-Object { $_.FullName.Length } -Descending |
    Remove-Item -Recurse -Force -ErrorAction SilentlyContinue

$nestedExecutables = Get-ChildItem -LiteralPath $root -Recurse -File -Filter '*.exe' |
    Where-Object { $_.FullName -notmatch '\.git\\' -and $_.DirectoryName -ne $root }
if ($nestedExecutables) {
    $nestedExecutables | Remove-Item -Force
    throw 'Nested executable files were found and removed; refusing to push until the repository is rechecked.'
}

$rootExecutables = @(Get-ChildItem -LiteralPath $root -File -Filter '*.exe' | Select-Object -ExpandProperty Name)
$expected = @("MailZen-$version.exe", "MailZenSetup-$version.exe")
if ((@($rootExecutables | Sort-Object) -join '|') -cne (@($expected | Sort-Object) -join '|')) {
    throw "Expected exactly $($expected -join ', ') in the repository root, found: $($rootExecutables -join ', ')"
}

git add -A
if (-not (git diff --cached --quiet)) {
    git commit -m "Build MailZen $version release"
}
Invoke-Checked 'git' @('-c', 'protocol.version=0', 'push', 'origin', 'main')

$local = (git rev-parse HEAD).Trim()
$remote = ((git ls-remote origin refs/heads/main).Split("`t")[0]).Trim()
if ($local -ne $remote) {
    throw "Push verification failed. Local commit $local does not match remote $remote."
}

Write-Host "Completed MailZen $version." -ForegroundColor Green
Write-Host "Root executables: $($expected -join ', ')" -ForegroundColor Green
Write-Host "GitHub main commit: $remote" -ForegroundColor Green
