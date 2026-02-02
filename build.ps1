<#
.SYNOPSIS
    Build system for Modern CharMap (WinUI 3).

.DESCRIPTION
    Cascading workflow: deploy -> build -> install -> doctor.
    Each verb automatically runs everything below it and short-circuits when
    the system is already in the correct state.

.PARAMETER Command
    doctor   - Check whether all required tooling is present and report.
    install  - Run doctor, then install / fix anything that is missing.
    build    - Run install, then restore, format, lint, and compile.
    deploy   - Run build, then publish a self-contained app with Start Menu shortcut.
               This is the DEFAULT when no command is given.

.PARAMETER Platform
    Target platform: x64 (default), x86, or ARM64.

.PARAMETER Configuration
    Build configuration: Release (default) or Debug.

.EXAMPLE
    .\build.ps1                    # full deploy (default)
    .\build.ps1 doctor             # just check the environment
    .\build.ps1 build -Platform x64 -Configuration Debug
#>
[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [ValidateSet('doctor', 'install', 'build', 'deploy')]
    [string]$Command = 'deploy',

    [ValidateSet('x64', 'x86', 'ARM64')]
    [string]$Platform = 'x64',

    [ValidateSet('Debug', 'Release')]
    [string]$Configuration = 'Release'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ── Project metadata ────────────────────────────────────────────────────────
$script:ProjectName = 'ModernCharMap'
$script:ProjectDir  = Join-Path $PSScriptRoot 'ModernCharMap.WinUI'
$script:ProjectFile = Join-Path $script:ProjectDir 'ModernCharMap.WinUI.csproj'
$script:DeployDir   = Join-Path $env:LOCALAPPDATA $script:ProjectName
$script:ExeName     = 'ModernCharMap.WinUI.exe'

$script:Requirements = @{
    DotNetMajor          = 9
    DotNetMinSdk         = '9.0.300'
    TargetFramework      = 'net9.0-windows10.0.19041.0'
    WindowsSdkBuild      = 19041
    WinAppSdkVersion     = '1.6.241114003'            # NuGet package version
    WinAppSdkMajorMinor  = '1.6'                      # for MSIX/DDLM matching
}

# Maps Platform param to the MSIX subfolder name
$script:PlatformToMsix = @{
    'x64'   = 'win10-x64'
    'x86'   = 'win10-x86'
    'ARM64' = 'win10-arm64'
}

# ── Helpers ─────────────────────────────────────────────────────────────────
function Write-Status {
    param([string]$Label, [bool]$Ok, [string]$Detail = '')
    $icon  = if ($Ok) { '[OK]' } else { '[MISSING]' }
    $color = if ($Ok) { 'Green' } else { 'Red' }
    $msg   = "$icon $Label"
    if ($Detail) { $msg += "  ($Detail)" }
    Write-Host $msg -ForegroundColor $color
}

function Find-DotNetSdk {
    <# Returns the best matching .NET SDK version string, or $null. #>
    $sdks = & dotnet --list-sdks 2>$null
    if (-not $sdks) { return $null }
    foreach ($line in $sdks) {
        if ($line -match '^(\d+\.\d+\.\d+)\s') {
            $ver = $Matches[1]
            if (([version]$ver).Major -eq $script:Requirements.DotNetMajor -and
                [version]$ver -ge [version]$script:Requirements.DotNetMinSdk) {
                return $ver
            }
        }
    }
    return $null
}

function Find-VSWhere {
    $candidates = @(
        "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe"
        "$env:ProgramFiles\Microsoft Visual Studio\Installer\vswhere.exe"
    )
    foreach ($c in $candidates) {
        if (Test-Path $c) { return $c }
    }
    return $null
}

function Find-VSInstallations {
    <#
    Returns an array of hashtables:
        Path, DisplayName, Version, MSBuildPath, HasPriTasks
    Scans both vswhere and well-known filesystem locations.
    #>
    $results = @()

    # --- vswhere ---------------------------------------------------------------
    $vswhere = Find-VSWhere
    if ($vswhere) {
        $raw = & $vswhere -all -format json -prerelease 2>$null
        $json = if ($raw) { $raw | ConvertFrom-Json } else { @() }
        foreach ($inst in $json) {
            $msbuild = Join-Path $inst.installationPath 'MSBuild\Current\Bin\amd64\MSBuild.exe'
            if (-not (Test-Path $msbuild)) {
                $msbuild = Join-Path $inst.installationPath 'MSBuild\Current\Bin\MSBuild.exe'
            }
            $priGlob     = Join-Path $inst.installationPath 'MSBuild\Microsoft\VisualStudio\v*\AppxPackage\Microsoft.Build.Packaging.Pri.Tasks.dll'
            $priResolved  = @(Resolve-Path $priGlob -ErrorAction SilentlyContinue)
            $hasPri       = $priResolved.Count -gt 0
            if (Test-Path $msbuild) {
                $results += @{
                    Path        = $inst.installationPath
                    DisplayName = $inst.displayName
                    Version     = $inst.installationVersion
                    MSBuildPath = $msbuild
                    HasPriTasks = $hasPri
                }
            }
        }
    }

    # --- well-known directories (catches Insiders / Preview) -------------------
    $searchRoots = @("$env:ProgramFiles\Microsoft Visual Studio")
    foreach ($root in $searchRoots) {
        if (-not (Test-Path $root)) { continue }
        Get-ChildItem $root -Directory -ErrorAction SilentlyContinue | ForEach-Object {
            Get-ChildItem $_.FullName -Directory -ErrorAction SilentlyContinue | ForEach-Object {
                $instPath = $_.FullName
                if ($results | Where-Object { $_.Path -eq $instPath }) { return }

                $msbuild = Join-Path $instPath 'MSBuild\Current\Bin\amd64\MSBuild.exe'
                if (-not (Test-Path $msbuild)) {
                    $msbuild = Join-Path $instPath 'MSBuild\Current\Bin\MSBuild.exe'
                }
                $priGlob     = Join-Path $instPath 'MSBuild\Microsoft\VisualStudio\v*\AppxPackage\Microsoft.Build.Packaging.Pri.Tasks.dll'
                $priResolved  = @(Resolve-Path $priGlob -ErrorAction SilentlyContinue)
                $hasPri       = $priResolved.Count -gt 0
                if (Test-Path $msbuild) {
                    $results += @{
                        Path        = $instPath
                        DisplayName = (Split-Path $instPath -Leaf)
                        Version     = 'unknown'
                        MSBuildPath = $msbuild
                        HasPriTasks = $hasPri
                    }
                }
            }
        }
    }
    return $results
}

function Select-BestVS {
    param([array]$Installations)
    $withPri = $Installations | Where-Object { $_.HasPriTasks }
    if ($withPri) { return ($withPri | Select-Object -First 1) }
    if ($Installations.Count -gt 0) { return $Installations[0] }
    return $null
}

function Get-WinAppSdkNuGetDir {
    <# Returns the NuGet package folder for the pinned WinAppSDK version. #>
    $dir = Join-Path $env:USERPROFILE ".nuget\packages\microsoft.windowsappsdk\$($script:Requirements.WinAppSdkVersion)"
    if (Test-Path $dir) { return $dir }
    return $null
}

function Test-WinAppRuntime {
    <#
    Checks that the Windows App Runtime MSIX packages for the required
    WinAppSDK version + platform are registered for the current user.
    WinAppSDK 1.6 -> package versions start with 6000.*
    WinAppSDK 1.7 -> 7000.*,  1.8 -> 8000.*, etc.
    #>
    $mm = $script:Requirements.WinAppSdkMajorMinor          # e.g. '1.6'
    $majorPrefix = "$($mm.Split('.')[1])000"                 # e.g. '6000'
    $archSuffix  = $Platform.ToLower().Substring(0, 2)       # 'x6' for x64, 'x8' for x86, 'ar' for ARM64

    # 1. Framework package:  Microsoft.WindowsAppRuntime.1.6
    $framework = @(Get-AppxPackage -Name "Microsoft.WindowsAppRuntime.$mm" -ErrorAction SilentlyContinue |
                   Where-Object { $_.Version -like "$majorPrefix.*" })
    if ($framework.Count -eq 0) { return $false }

    # 2. DDLM package:  Microsoft.WinAppRuntime.DDLM.<version>-<arch>
    $ddlm = @(Get-AppxPackage -Name "Microsoft.WinAppRuntime.DDLM.$majorPrefix.*-$archSuffix" -ErrorAction SilentlyContinue)
    if ($ddlm.Count -eq 0) { return $false }

    # 3. Main package:  MicrosoftCorporationII.WinAppRuntime.Main.1.6
    $main = @(Get-AppxPackage -Name "MicrosoftCorporationII.WinAppRuntime.Main.$mm" -ErrorAction SilentlyContinue |
              Where-Object { $_.Version -like "$majorPrefix.*" })
    if ($main.Count -eq 0) { return $false }

    return $true
}

function Install-WinAppRuntimeMsix {
    <#
    Registers the MSIX packages from the NuGet cache for the current platform.
    #>
    $nugetDir = Get-WinAppSdkNuGetDir
    if (-not $nugetDir) {
        Write-Warning "Windows App SDK NuGet package not found at expected path. Run 'dotnet restore' first."
        return $false
    }
    $msixSub = $script:PlatformToMsix[$Platform]
    $msixDir = Join-Path $nugetDir "tools\MSIX\$msixSub"
    if (-not (Test-Path $msixDir)) {
        Write-Warning "MSIX directory not found: $msixDir"
        return $false
    }

    $msixFiles = Get-ChildItem $msixDir -Filter '*.msix' | Sort-Object Name
    foreach ($msix in $msixFiles) {
        Write-Host "  Registering $($msix.Name) ..." -ForegroundColor Yellow
        try {
            Add-AppxPackage -Path $msix.FullName -ErrorAction Stop
        }
        catch {
            # Already installed or newer version present is fine
            if ($_.Exception.Message -match 'already installed' -or
                $_.Exception.Message -match 'higher version') {
                Write-Host "    (already satisfied)" -ForegroundColor DarkGray
            }
            else {
                Write-Warning "    Failed: $($_.Exception.Message)"
            }
        }
    }
    return $true
}

function New-StartMenuShortcut {
    param(
        [string]$ExePath,
        [string]$ShortcutName
    )
    $startMenu = Join-Path $env:APPDATA 'Microsoft\Windows\Start Menu\Programs'
    $lnkPath   = Join-Path $startMenu "$ShortcutName.lnk"

    $shell    = New-Object -ComObject WScript.Shell
    $shortcut = $shell.CreateShortcut($lnkPath)
    $shortcut.TargetPath       = $ExePath
    $shortcut.WorkingDirectory = (Split-Path $ExePath -Parent)
    # Use the .ico file directly — more reliable than extracting from the EXE resource.
    $icoPath = Join-Path (Split-Path $ExePath -Parent) 'Assets\app.ico'
    if (Test-Path $icoPath) {
        $shortcut.IconLocation = "$icoPath,0"
    } else {
        $shortcut.IconLocation = "$ExePath,0"
    }
    $shortcut.Description      = 'Modern CharMap - Unicode character viewer'
    $shortcut.Save()

    # Release COM
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($shortcut) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($shell)    | Out-Null

    return $lnkPath
}

# ── DOCTOR ──────────────────────────────────────────────────────────────────
function Invoke-Doctor {
    Write-Host ''
    Write-Host '=== Doctor ===' -ForegroundColor Cyan
    Write-Host ''

    $status = [ordered]@{
        DotNetSdk      = $null
        VSInstall      = $null
        MSBuildPath    = $null
        HasPriTasks    = $false
        ProjectExists  = $false
        WinAppRuntime  = $false
        NuGetDir       = $null
        CanBuild       = $false
        AllGood        = $false
    }

    # 1. .NET SDK
    $sdk = Find-DotNetSdk
    $status.DotNetSdk = $sdk
    Write-Status ".NET $($script:Requirements.DotNetMajor) SDK (>= $($script:Requirements.DotNetMinSdk))" `
                 ($null -ne $sdk) `
                 $(if ($sdk) { "found $sdk" } else { "run: winget install Microsoft.DotNet.SDK.9" })

    # 2. Visual Studio + MSBuild
    $vsAll  = Find-VSInstallations
    $vsBest = Select-BestVS -Installations $vsAll
    $status.VSInstall   = $vsBest
    $status.MSBuildPath = if ($vsBest) { $vsBest.MSBuildPath } else { $null }
    $status.HasPriTasks = if ($vsBest) { $vsBest.HasPriTasks } else { $false }

    $vsOk = $null -ne $vsBest
    Write-Status 'Visual Studio with MSBuild' $vsOk `
                 $(if ($vsBest) { "$($vsBest.DisplayName) - $($vsBest.Path)" } else {
                     'Install Visual Studio 2022+ with .NET desktop workload'
                 })

    Write-Status 'PRI tasks (Microsoft.Build.Packaging.Pri.Tasks.dll)' $status.HasPriTasks `
                 $(if ($status.HasPriTasks) { 'present' } else {
                     'Install the Windows App SDK VS extension or Desktop workload'
                 })

    # 3. Project file
    $status.ProjectExists = Test-Path $script:ProjectFile
    Write-Status 'Project file' $status.ProjectExists $script:ProjectFile

    # 4. WinAppSDK NuGet package
    $nugetDir = Get-WinAppSdkNuGetDir
    $status.NuGetDir = $nugetDir
    Write-Status "Windows App SDK $($script:Requirements.WinAppSdkMajorMinor) NuGet package" `
                 ($null -ne $nugetDir) `
                 $(if ($nugetDir) { $nugetDir } else { "run: dotnet restore" })

    # 5. Windows App Runtime MSIX packages (DDLM, Main, Singleton)
    $runtimeOk = $false
    if ($nugetDir) {
        $runtimeOk = Test-WinAppRuntime
    }
    $status.WinAppRuntime = $runtimeOk
    Write-Status "Windows App Runtime $($script:Requirements.WinAppSdkMajorMinor) ($Platform)" `
                 $runtimeOk `
                 $(if ($runtimeOk) { 'all MSIX packages registered' } else {
                     "run: .\build.ps1 install"
                 })

    # 6. dotnet format
    $fmtOk = $null -ne $sdk
    Write-Status 'dotnet format (code formatter)' $fmtOk 'ships with .NET SDK'

    # Summary: build prerequisites vs run prerequisites
    $status.CanBuild = (
        $null -ne $sdk -and
        $vsOk -and
        $status.HasPriTasks -and
        $status.ProjectExists -and
        $null -ne $nugetDir
    )
    $status.AllGood = $status.CanBuild -and $runtimeOk

    Write-Host ''
    if ($status.AllGood) {
        Write-Host 'All checks passed.' -ForegroundColor Green
    }
    elseif ($status.CanBuild) {
        Write-Host 'Build prerequisites OK. Runtime packages missing (needed to run, not build).' -ForegroundColor Yellow
        Write-Host 'Run:  .\build.ps1 install' -ForegroundColor Yellow
    }
    else {
        Write-Host 'Some checks failed. Run:  .\build.ps1 install' -ForegroundColor Yellow
    }
    Write-Host ''

    return $status
}

# ── INSTALL ─────────────────────────────────────────────────────────────────
function Invoke-Install {
    Write-Host ''
    Write-Host '=== Install ===' -ForegroundColor Cyan
    Write-Host ''

    $status = Invoke-Doctor

    if ($status.AllGood) {
        Write-Host 'Nothing to install - environment is ready.' -ForegroundColor Green
        return $status
    }

    $changed = $false

    # 1. .NET SDK
    if (-not $status.DotNetSdk) {
        Write-Host 'Installing .NET 9 SDK via winget ...' -ForegroundColor Yellow
        $winget = Get-Command winget -ErrorAction SilentlyContinue
        if ($winget) {
            & winget install Microsoft.DotNet.SDK.9 --accept-source-agreements --accept-package-agreements
            $changed = $true
        }
        else {
            Write-Warning 'winget not found. Please install .NET 9 SDK manually from https://dot.net/download'
        }
    }

    # 2. Visual Studio / MSBuild / PRI tasks
    if (-not $status.VSInstall -or -not $status.HasPriTasks) {
        $vsInstaller = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vs_installer.exe"
        $hasInstaller = Test-Path $vsInstaller

        if ($hasInstaller -and $status.VSInstall) {
            Write-Host 'Adding required VS workloads ...' -ForegroundColor Yellow
            $installPath = $status.VSInstall.Path
            $vsModifyArgs = @('modify', '--installPath', "`"$installPath`"", '--quiet',
                              '--add', 'Microsoft.VisualStudio.Workload.ManagedDesktop',
                              '--add', 'Microsoft.VisualStudio.Component.Windows11SDK.22621')
            Start-Process $vsInstaller -ArgumentList $vsModifyArgs -Wait -NoNewWindow
            $changed = $true
        }
        elseif (-not $status.VSInstall) {
            $winget = Get-Command winget -ErrorAction SilentlyContinue
            if ($winget) {
                Write-Host 'Installing Visual Studio 2022 Community via winget ...' -ForegroundColor Yellow
                & winget install Microsoft.VisualStudio.2022.Community `
                    --override "--add Microsoft.VisualStudio.Workload.ManagedDesktop --add Microsoft.VisualStudio.Component.Windows11SDK.22621 --passive" `
                    --accept-source-agreements --accept-package-agreements
                $changed = $true
            }
            else {
                Write-Warning @"
Visual Studio not found and winget is unavailable.
Install Visual Studio 2022+ with the .NET Desktop Development workload:
https://visualstudio.microsoft.com/downloads/
"@
            }
        }
    }

    # 3. NuGet restore
    if (Test-Path $script:ProjectFile) {
        Write-Host 'Restoring NuGet packages ...' -ForegroundColor Yellow
        & dotnet restore $script:ProjectFile --verbosity quiet
        if ($LASTEXITCODE -ne 0) {
            Write-Warning 'NuGet restore had issues; build may still work via MSBuild -restore.'
        }
        $changed = $true
    }

    # 4. Windows App Runtime MSIX packages
    if (-not $status.WinAppRuntime) {
        Write-Host "Registering Windows App Runtime $($script:Requirements.WinAppSdkMajorMinor) MSIX packages ..." -ForegroundColor Yellow
        $installed = Install-WinAppRuntimeMsix
        if ($installed) { $changed = $true }
    }

    # Re-check
    if ($changed) {
        Write-Host ''
        Write-Host 'Re-checking after install ...' -ForegroundColor Cyan
        $status = Invoke-Doctor

        if (-not $status.WinAppRuntime -and $status.CanBuild) {
            Write-Host 'Note: Runtime packages may need a new terminal session to be detected.' -ForegroundColor DarkYellow
            Write-Host '      Build and deploy (self-contained) will still work.' -ForegroundColor DarkYellow
        }
    }

    return $status
}

# ── BUILD ───────────────────────────────────────────────────────────────────
function Invoke-Build {
    $status = Invoke-Install

    Write-Host ''
    Write-Host '=== Build ===' -ForegroundColor Cyan
    Write-Host ''

    if (-not $status.CanBuild) {
        Write-Error 'Cannot build - doctor reported failures. Fix them first.'
        return
    }

    # Cache MSBuild path for later phases (deploy)
    $script:CachedMSBuildPath = $status.MSBuildPath
    $msbuild = $status.MSBuildPath

    # 1. Code formatting
    Write-Host '[1/3] Formatting code ...' -ForegroundColor White
    & dotnet format $script:ProjectFile --verbosity quiet 2>$null
    if ($LASTEXITCODE -ne 0) {
        Write-Warning 'dotnet format reported issues (non-fatal).'
    }
    else {
        Write-Host '      Format OK' -ForegroundColor Green
    }

    # 2. Lint (verify formatting is clean)
    Write-Host '[2/3] Linting (verify formatting) ...' -ForegroundColor White
    & dotnet format $script:ProjectFile --verify-no-changes --verbosity quiet 2>$null
    if ($LASTEXITCODE -ne 0) {
        Write-Warning 'Lint: formatting drifted after format pass (check .editorconfig).'
    }
    else {
        Write-Host '      Lint OK' -ForegroundColor Green
    }

    # 3. Compile via VS MSBuild (required for PRI generation)
    Write-Host "[3/3] Compiling ($Configuration | $Platform) ..." -ForegroundColor White
    & $msbuild $script:ProjectFile `
        "-p:Configuration=$Configuration" `
        "-p:Platform=$Platform" `
        '-restore' `
        '-m' `
        '-verbosity:minimal'

    if ($LASTEXITCODE -ne 0) {
        Write-Error "Build FAILED (exit code $LASTEXITCODE)."
        return
    }

    Write-Host ''
    Write-Host 'Build succeeded.' -ForegroundColor Green
    Write-Host ''
}

# ── DEPLOY ──────────────────────────────────────────────────────────────────
function Invoke-Deploy {
    Invoke-Build

    Write-Host ''
    Write-Host '=== Deploy ===' -ForegroundColor Cyan
    Write-Host ''

    $msbuild = $script:CachedMSBuildPath
    $tfm     = $script:Requirements.TargetFramework

    # 1. Build self-contained, then copy full output to deploy dir.
    #    MSBuild Publish for WinUI 3 can skip the PRI and Content files,
    #    so we build normally with self-contained flag and copy everything.
    Write-Host 'Building self-contained ...' -ForegroundColor White
    & $msbuild $script:ProjectFile `
        "-p:Configuration=$Configuration" `
        "-p:Platform=$Platform" `
        "-p:WindowsAppSDKSelfContained=true" `
        '-restore' `
        '-m' `
        '-verbosity:minimal'

    if ($LASTEXITCODE -ne 0) {
        Write-Error "Build (self-contained) FAILED (exit code $LASTEXITCODE)."
        return
    }

    $binDir = Join-Path $script:ProjectDir "bin\$Platform\$Configuration\$tfm"
    if (-not (Test-Path $binDir)) {
        Write-Error "Build output not found at $binDir."
        return
    }

    # Stop any running instance before overwriting
    $procName = [System.IO.Path]::GetFileNameWithoutExtension($script:ExeName)
    $running  = Get-Process -Name $procName -ErrorAction SilentlyContinue
    if ($running) {
        Write-Host "Stopping running $procName ..." -ForegroundColor Yellow
        $running | Stop-Process -Force
        Start-Sleep -Seconds 1
    }

    Write-Host 'Copying to deploy directory ...' -ForegroundColor White
    if (Test-Path $script:DeployDir) { Remove-Item $script:DeployDir -Recurse -Force }
    New-Item $script:DeployDir -ItemType Directory -Force | Out-Null
    # Use robocopy for reliable recursive copy (Copy-Item can miss subdirectories).
    # robocopy exit codes: 0-7 = success, 8+ = error.
    & robocopy $binDir $script:DeployDir /E /NFL /NDL /NJH /NJS /R:1 /W:1
    if ($LASTEXITCODE -ge 8) {
        Write-Error "robocopy FAILED (exit code $LASTEXITCODE)."
        return
    }
    $LASTEXITCODE = 0  # Reset so PowerShell doesn't treat robocopy 1 as failure

    $exe = Join-Path $script:DeployDir $script:ExeName
    if (-not (Test-Path $exe)) {
        Write-Error "Executable not found at $exe after deploy."
        return
    }

    # 2. Start Menu shortcut
    Write-Host 'Creating Start Menu shortcut ...' -ForegroundColor White
    $lnk = New-StartMenuShortcut -ExePath $exe -ShortcutName $script:ProjectName
    Write-Host "      $lnk" -ForegroundColor DarkGray

    Write-Host ''
    Write-Host "Deployed to:   $script:DeployDir" -ForegroundColor Green
    Write-Host "Start Menu:    $lnk" -ForegroundColor Green
    Write-Host "Run:           $exe" -ForegroundColor Green
    Write-Host ''
}

# ── Main dispatch ───────────────────────────────────────────────────────────
switch ($Command) {
    'doctor'  { Invoke-Doctor  | Out-Null }
    'install' { Invoke-Install | Out-Null }
    'build'   { Invoke-Build }
    'deploy'  { Invoke-Deploy }
}
