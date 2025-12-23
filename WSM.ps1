# ===========================
# NSSM 服务管理菜单
# ===========================

# 检测管理员权限
function Ensure-RunAsAdmin {
    # 获取当前用户身份
    $currentIdentity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentIdentity)
    if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        Write-Warning "$($L.runAsAdminPrompt)"
        Read-Host
        Exit
    }
}

# 默认中文语言包的JSON字符串
$defaultLanguageJson = @"
{
  "appName": "Windows Service Manager (WSM)",
  "runAsAdminPrompt": "Please run this program as administrator",
  "serviceList": "Service List",
  "addService": "Add Service",
  "setting": "Settings",
  "about": "About",
  "exit": "Exit",
  "selectOperation": "Please select an operation",
  "return": "Return",
  "returnMainMenu": "Return to Main Menu",
  "noManagedServices": "No services available to manage",
  "pressEnterToContinue": "Press Enter to continue",
  "pressEnterToExit": "Press Enter to exit the program",
  "currentStatus": "Current Status",
  "startService": "Start Service",
  "start": "Starting...",
  "stop": "Stopping...",
  "restart": "Restarting...",
  "stopService": "Stop Service",
  "success": "Success",
  "failure": "Failure",
  "restartService": "Restart Service",
  "viewDetails": "View Details",
  "serviceDetails": "Service Details",
  "changeStartupType": "Change Startup Type",
  "changingStartupType": "Changing startup type...",
  "currentStartupType": "Current Startup Type: ",
  "selectNewStartupType": "Select new startup type",
  "invalidChoice": "Invalid choice",
  "startupTypeUpdated": "Startup type updated to:",
  "editService": "Edit Service",
  "openingServiceEditInterface": "Opening service edit interface...",
  "serviceEditComplete": "Service editing complete...",
  "deleteService": "Delete Service",
  "confirmDeleteService": "Are you sure you want to delete this service from the system? (y/N)",
  "deleteServiceWarning": "!!! Warning !!! This action cannot be undone",
  "confirmReallyDelete": "Are you really sure you want to delete? (y/N)",
  "cancelDelete": "Cancel deletion",
  "deletingService": "Deleting service...",
  "serviceNotExist": "Service does not exist",
  "serviceRunningStopping": "Service is running, stopping service...",
  "serviceDeleted": "Service has been deleted",
  "returnServiceList": "Return to Service List",
  "addNssmService": "Add NSSM Service",
  "nssmNotInstalledPrompt": "NSSM is not installed. Please install NSSM in the settings first",
  "addCustomServices": "Add Custom Service List",
  "addServicesToMonitor": "Add services to monitor here",
  "installNssm": "Install NSSM",
  "nssmNotFoundPrompt": "NSSM (nssm.exe) not found. Do you want to install it? (y/N)",
  "nssmInstallPrompt": "Would you like to download and install NSSM? (y/N)",
  "downloadingNssm": "Downloading NSSM...",
  "extractingNssm": "Download complete, extracting...",
  "notFound": "Not Found",
  "addNssmToPathPrompt": "Would you like to add NSSM to the system PATH? (y/N)",
  "pathAlreadyExists": "This directory already exists in PATH",
  "addedToPathSuccess": "Successfully added to system PATH",
  "downloadOrInstallFailed": "Download or installation of NSSM failed:",
  "nssmInstallComplete": "NSSM installation complete, path: ",
  "nssmAlreadyInstalled": "NSSM is already installed, path: ",
  "installationCancelled": "Installation cancelled",
  "toggleTheme": "Toggle Theme",
  "chooseTheme": "Choose Theme",
  "themeFolderNotFound": "Theme folder not found",
  "themeFileNotFoundInFolder": "No theme files found in the theme folder",
  "resetToDefaultTheme": "Reset to default theme",
  "aboutBody": [
    " [ Program Information ]",
    "  Name        : {0}",
    "  Version     : {1}",
    "  Author      : {2}",
    "  Update      : {3}",
    "  Path        : {4}",
    "  GitHub      : https://github.com/Zhu-junwei/wsm",
    "  NSSM Ver    : {5}",
    "  NSSM Path   : {6}",
    "  NSSM Web    : https://nssm.cc",
    " ",
    " [ Key Features ]",
    "  Manage all Windows services hosted by NSSM",
    "  Manage services defined in {7}",
    "  Support Start, Stop, Restart, and Delete operations",
    "  Native NSSM GUI to wrap executables (EXE, BAT, JAR, Python) as services",
    "  Change Startup Type (Automatic / Manual / Disabled)",
    "  View detailed parameters (Path, Args, Working Dir)",
    "  Auto-detect and online installation of NSSM",
    "  Support for custom UI themes",
    " ",
    " [ What is NSSM? ]",
    "  NSSM (Non-Sucking Service Manager) is a service helper that",
    "  encapsulates any executable as a standard Windows service.",
    " ",
    " [ Important Notes ]",
    "  Administrator privileges are required",
    "  Service deletion is permanent and irreversible",
    "  Please ensure correct paths before editing"
  ]
}
"@
function Load-Language {
    $osLanguage = [System.Globalization.CultureInfo]::CurrentCulture.Name
    $languageFilePath = Join-Path -Path $PSScriptRoot -ChildPath "languages\$osLanguage.json"
    if (Test-Path $languageFilePath) {
        try {
            $languageJson = Get-Content -Path $languageFilePath -Encoding UTF8 | ConvertFrom-Json
        } catch {
            $languageJson = $defaultLanguageJson | ConvertFrom-Json
        }
    } else {
        $languageJson = $defaultLanguageJson | ConvertFrom-Json
    }
    return $languageJson
}
$Global:L = Load-Language


# 加载UI插件
$BoxPath = Join-Path $PSScriptRoot 'plugins/Box.ps1'
if (Test-Path $BoxPath) { 
	. $BoxPath 
} else {
	Write-Warning "Missing plugins/Box.ps1 plugin file"
	Read-Host
	exit
}

# 初始化必要参数
function Initialize-Parameters() {
	$Global:ScriptName = "Windows服务管理(WSM)"
	$Global:ScriptUser = "zjw"
	$Global:ScriptVersion = "v1.3.0"
	$Global:ScriptUpdate = "20251223"
	$Global:ServiceFile = "services.txt"
	$Global:NssmInstallDir = "$env:ProgramFiles\NSSM"
	$Global:NssmZipUrl    = "https://nssm.cc/ci/nssm-2.24-103-gdee49fc.zip"
	$Global:ThemeDir = 'themes'
	$Global:ExitKeys = @("0", "q", "quit", "exit")
	$Global:WSMServiceStore = @()
	$Global:NssmInfo = [PSCustomObject]@{
        Path    = ''
        Version = ''
    }
}

# 定义默认的UI主题
function Initialize-DefaultTheme() {
	$Global:UI.Width = 50
	$Global:UI.BoxStyle = 'Single'
	$Global:UI.BorderColor = 'DarkCyan'
	$Global:UI.TextPaddingLeft = 2
	$Global:UI.TextColor  = 'Yellow'
	$Global:UI.MutedColor  = 'DarkGray'
}
# 加载生效的主题
function Load-SavedTheme {
	try {
		$ThemeDir = Join-Path $PSScriptRoot $Global:ThemeDir
		$ThemeSaveFile = Join-Path $ThemeDir "current_theme.txt"
		if (Test-Path $ThemeSaveFile) {
			$themeName = (Get-Content $ThemeSaveFile -Encoding UTF8).Trim()
			$themePath = Join-Path $ThemeDir "$themeName.ps1"
			if (Test-Path $themePath) {
				. "$themePath"
			}
		} else {
			Initialize-DefaultTheme
		}
	} catch {
        Initialize-DefaultTheme
    }
}

# ---------------------------
# 加载所有需要管理的服务
# ---------------------------
function Initialize-Services {
	$customServices = @()
	$ServiceFilePath = Join-Path $PSScriptRoot $Global:ServiceFile
    if (Test-Path $ServiceFilePath) {
		$customServices = Get-Content $ServiceFilePath | Where-Object {
			$_.Trim() -ne "" -and $_ -notmatch '^\s*(#|;|//)'
		}
	}
	$Global:WSMServiceStore = @(
		Get-CimInstance Win32_Service | Where-Object {
			($_.PathName -match "nssm.exe") -or ($customServices -contains $_.Name)
		}
	)
}

# 列表获取所有服务
function Get-WSMServices {
    return $Global:WSMServiceStore | Sort-Object DisplayName
}

# ---------------------------
# 读取菜单输入
# ---------------------------
function Read-MenuChoice {
    param([string]$Prompt = "$($L.selectOperation)")
    (Read-Host $Prompt).Trim().ToLower()
}

# ---------------------------
# 下载并安装 NSSM
# ---------------------------
function Install-Nssm {
    $tempZip = Join-Path $env:TEMP "nssm.zip"
    $tempExtractDir = Join-Path $env:TEMP "nssm_extract"

    try {
        # 下载 ZIP
        Write-Host "`n$($L.downloadingNssm)" -ForegroundColor Yellow
        Invoke-WebRequest -Uri $Global:NssmZipUrl -OutFile $tempZip -UseBasicParsing
        Write-Host "$($L.extractingNssm)" -ForegroundColor Yellow

        # 创建安装目录
        if (-not (Test-Path $Global:NssmInstallDir)) {
            New-Item -Path $Global:NssmInstallDir -ItemType Directory | Out-Null
        }

        # 清理临时解压目录
        if (Test-Path $tempExtractDir) { Remove-Item $tempExtractDir -Recurse -Force }
        New-Item -Path $tempExtractDir -ItemType Directory | Out-Null

        # 解压 ZIP
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        [System.IO.Compression.ZipFile]::ExtractToDirectory($tempZip, $tempExtractDir)

        # 找到所有 nssm.exe
        $nssmExes = Get-ChildItem -Path $tempExtractDir -Recurse -Filter "nssm.exe"
        if (-not $nssmExes -or $nssmExes.Count -eq 0) { throw "$($L.notFound) nssm.exe" }

        # 判断系统位数
        $is64 = [Environment]::Is64BitOperatingSystem
        $targetArch = if ($is64) { "win64" } else { "win32" }

        # 选择正确架构的 nssm.exe
        $sourceExe = $nssmExes | Where-Object { $_.Directory.Name -ieq $targetArch } | Select-Object -First 1
        if (-not $sourceExe) { throw "$($L.notFound) $targetArch\nssm.exe" }

        # 移动到安装目录根目录
        $destExe = Join-Path $Global:NssmInstallDir "nssm.exe"
        Move-Item -Path $sourceExe.FullName -Destination $destExe -Force
        Write-Host "$($L.nssmInstallComplete) $destExe" -ForegroundColor Green
		Write-Host
        $choice = Read-Host "$($L.addNssmToPathPrompt)"
        if ($choice -match '^[Yy]$') {
			$dir = $Global:NssmInstallDir.TrimEnd('\')
			$exists = @(
				[Environment]::GetEnvironmentVariable("PATH", "Machine"),
				[Environment]::GetEnvironmentVariable("PATH", "User")
			) -join ';' -split ';' |
				ForEach-Object { $_.TrimEnd('\') } |
				Where-Object { $_ -eq $dir } |
				Select-Object -First 1

			if ($exists) {
				Write-Host "$($L.pathAlreadyExists)" -ForegroundColor Yellow
			} else {
				$machinePath = [Environment]::GetEnvironmentVariable("PATH", "Machine")
				[Environment]::SetEnvironmentVariable("PATH", "$machinePath;$dir", "Machine")
				Write-Host "$($L.addedToPathSuccess)" -ForegroundColor Green
			}
		}
		Start-Sleep 1
        return $destExe
    } catch {
        Write-Host "$($L.downloadOrInstallFailed) $_" -ForegroundColor Red
        Read-Host "`n$($L.pressEnterToExit)"
        Exit
    } finally {
        # 清理临时文件和目录
        if (Test-Path $tempZip) { Remove-Item $tempZip -Force }
        if (Test-Path $tempExtractDir) { Remove-Item $tempExtractDir -Recurse -Force }
    }
}

# ---------------------------
# 初始化 NSSM（自动检测）
# ---------------------------
function Initialize-Nssm {
    # 临时加入当前 PowerShell 会话 PATH
    if ($env:PATH -notmatch [Regex]::Escape($Global:NssmInstallDir)) {
        $env:PATH = "$Global:NssmInstallDir;$env:PATH"
    }
    $env:PATH = "$PSScriptRoot;$env:PATH"

    # 尝试获取 nssm.exe 路径
    $nssmPath = (Get-Command nssm.exe -ErrorAction SilentlyContinue).Source

    if ($nssmPath) {
        # 已安装，设置全局信息
        $fullVersion = & $nssmPath version
        $Global:NssmInfo = [PSCustomObject]@{
            Path    = $nssmPath
            Version = $fullVersion
        }
    } else {
        # 未安装，启动时不提示
        $Global:NssmInfo = $null
    }
}

# ---------------------------
# 安装 NSSM（用户主动操作）
# ---------------------------
function Install-NssmIfMissing {
    if ($Global:NssmInfo -and (Test-Path $Global:NssmInfo.Path)) {
        Write-Host "$($L.nssmAlreadyInstalled) $($Global:NssmInfo.Path)" -ForegroundColor Green
		Read-Host
        return
    }

    Write-Host "$($L.nssmNotFoundPrompt)" -ForegroundColor Yellow
    $choice = Read-Host "$($L.nssmInstallPrompt)"
    if ($choice -match "^[Yy]$") {
        $nssmPath = Install-Nssm
        $fullVersion = & $nssmPath version
        $Global:NssmInfo = [PSCustomObject]@{
            Path    = $nssmPath
            Version = $fullVersion
        }
        Write-Host "$($L.nssmInstallComplete)$nssmPath" -ForegroundColor Green
        Start-Sleep 1
    } else {
        Write-Host "`n$($L.installationCancelled)" -ForegroundColor Red
        Read-Host "`n$($L.pressEnterToContinue)"
    }
}

# ---------------------------
# 编辑服务文件
# ---------------------------
function Edit-ServiceFile {
	$ServiceFilePath = Join-Path $PSScriptRoot $Global:ServiceFile
    if (-not (Test-Path $ServiceFilePath)) {
        "# $($L.addServicesToMonitor)" | Out-File -FilePath $ServiceFilePath -Encoding UTF8
    }
    Start-Process -FilePath "notepad.exe" -ArgumentList $ServiceFilePath
}

# ---------------------------
# 缩进输出函数
# ---------------------------
function Write-Indented {
    param(
        [string]$Text,
        [int]$Indent = 4,
        [ConsoleColor]$Color = "White"
    )
    $spaces = " " * $Indent
    Write-Host "$spaces$Text" -ForegroundColor $Color
}

function Get-DisplayWidth($text) {
    $width = 0
    foreach ($ch in $text.ToCharArray()) {
        $width += if ($ch -match '[\u1100-\u115F\u2E80-\uA4CF\uAC00-\uD7A3\uF900-\uFAFF\uFE10-\uFE6F\uFF00-\uFF60]') { 2 } else { 1 }
    }
    return $width
}
function PadRightWidth($text, $width) {
    $displayWidth = 0
    foreach ($ch in $text.ToCharArray()) {
        $displayWidth += if ($ch -match '[\u1100-\u115F\u2E80-\uA4CF\uAC00-\uD7A3\uF900-\uFAFF\uFE10-\uFE6F\uFF00-\uFF60]') { 2 } else { 1 }
    }
    $padding = $width - $displayWidth
    if ($padding -gt 0) {
        return $text + (' ' * $padding)
    } else {
        return $text
    }
}

# ---------------------------
# 显示服务列表菜单
# ---------------------------
function Show-ServiceListMenu {
	while ($true) {
		Initialize-Services
		$services = @(Get-WSMServices)
		Clear-Host
		if (-not $services -or $services.Count -eq 0) {
			$menuItems = @(
				$null,
				@{Text="$($L.noManagedServices)"; Color='Yellow';},
				$null,
				@{Text="0. $($L.returnMainMenu)"; Color=$Global:UI.MutedColor;}
			)
			Show-BoxMenu -Title "$($L.serviceList)" -MenuItems $menuItems -Wrap
		} else {
			# 计算服务名称列对齐长度
			$maxNameLength = ($services | ForEach-Object { Get-DisplayWidth $_.DisplayName } | Measure-Object -Maximum).Maximum
			$menuItems = @($null)
			$i = 1
			foreach ($svc in $services) {
				$nameText = "{0}. {1}" -f $i, (PadRightWidth $svc.DisplayName $maxNameLength)
				$menuItems += @{
					Text  = $nameText + "  $($svc.State)"
					Color = (Get-StateColor $svc.State)
				}
				$i++
			}

			# 添加空行和返回主菜单选项
			$menuItems += @(
				$null,
				@{Text="0. $($L.returnMainMenu)"; Color=$Global:UI.MutedColor; Align='Left'},
				$null
			)

			Show-BoxMenu -Title "$($L.serviceList)" -MenuItems $menuItems -Wrap
		}

		# 读取用户选择
		$selection = Read-MenuChoice
		if ($Global:ExitKeys -contains $selection) { return }
		if ($selection -match "^\d+$" -and $selection -ge 1 -and $selection -le $services.Count) {
			$svc = $services[$selection - 1]
			Show-ServiceManagementMenu $svc
		} else {
			continue
		}
	}
}

function Get-StateColor {
    param(
        [string]$State
    )
    switch ($State) {
        'Running' { [ConsoleColor]::Green }
        'Stopped' { [ConsoleColor]::Red }
        default   { [ConsoleColor]::Yellow }
    }
}

function Invoke-ServiceAction {
    param (
        [Parameter(Mandatory)]
        [string]$ServiceName,

        [Parameter(Mandatory)]
        [ValidateSet("Start", "Stop", "Restart")]
        [string]$Action,

        [int]$TimeoutSeconds = 5
    )

    $actionText = @{
        Start   = "$($L.start)"
        Stop    = "$($L.stop)"
        Restart = "$($L.restart)"
    }[$Action]

    Write-Host "`n${actionText}..." -ForegroundColor Yellow

    try {
        switch ($Action) {
            "Start" {
                Start-Service -Name $ServiceName -ErrorAction Stop
                (Get-Service $ServiceName).WaitForStatus('Running', "00:00:$TimeoutSeconds")
            }
            "Stop" {
                Stop-Service -Name $ServiceName -Force -ErrorAction Stop
                (Get-Service $ServiceName).WaitForStatus('Stopped', "00:00:$TimeoutSeconds")
            }
            "Restart" {
                Restart-Service -Name $ServiceName -Force -ErrorAction Stop
                (Get-Service $ServiceName).WaitForStatus('Running', "00:00:$TimeoutSeconds")
            }
        }

        Write-Host "${actionText} $($L.success)" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "`n${actionText} $($L.failure)" -ForegroundColor Red
        Write-Host "--------------------------------" -ForegroundColor DarkGray
         Write-Host "`n[Message]" -ForegroundColor Yellow
		Write-Host $_.Exception.Message -ForegroundColor DarkRed

		Write-Host "`n[Command]" -ForegroundColor Yellow
		Write-Host $_.InvocationInfo.MyCommand

		Write-Host "`n[Location]" -ForegroundColor Yellow
		Write-Host $_.InvocationInfo.PositionMessage

		if ($_.Exception.InnerException) {
			Write-Host "`n[InnerException]" -ForegroundColor Yellow
			Write-Host $_.Exception.InnerException.Message -ForegroundColor DarkRed
		}

		Write-Host "`n[StackTrace]" -ForegroundColor Yellow
		Write-Host $_.ScriptStackTrace -ForegroundColor DarkGray
        Write-Host "--------------------------------" -ForegroundColor DarkGray
		Show-ServiceDetails $svc
		Read-Host "`n$($L.pressEnterToContinue)"
		return $false
    }
}

# ---------------------------
# 服务管理菜单
# ---------------------------
function Show-ServiceManagementMenu {
    param($svc)
    while ($true) {
		$svc = Get-CimInstance -ClassName Win32_Service -Filter "Name='$($svc.Name)'"
		if (-not $svc) {
			Write-Host "`n$($svc.Name) $($L.serviceNotExist)" -ForegroundColor Red
			Read-Host
			return
		}
		Clear-Host
		$menuItems = @(
			$null,
			@{Text=" $($L.currentStatus) : $($svc.State)";Color=(Get-StateColor $svc.State)},
			$null,
			@{Text="1. $($L.startService)"},
			@{Text="2. $($L.stopService)"},
			@{Text="3. $($L.restartService)"},
			@{Text="4. $($L.viewDetails)"},
			@{Text="5. $($L.changeStartupType)"},
			@{Text="6. $($L.editService)"},
			@{Text="7. $($L.deleteService)"}
			$null,
			@{Text="0. $($L.returnServiceList)";Color=$Global:UI.MutedColor;},
			$null
		)
		Show-BoxMenu -Title "$($svc.DisplayName)" -MenuItems $menuItems
        $choice = Read-MenuChoice

        switch ($choice) {
            "1" {
				$null = Invoke-ServiceAction -ServiceName $svc.Name -Action Start
				#Start-Sleep 1
			}
            "2" {
                $null = Invoke-ServiceAction -ServiceName $svc.Name -Action Stop
            }
            "3" {
                $null = Invoke-ServiceAction -ServiceName $svc.Name -Action Restart
            }
            "4" { Show-ServiceDetails $svc; Read-Host "`n$($L.pressEnterToContinue)" }
            "5" {
                Write-Host "`n$($L.changingStartupType)" -ForegroundColor Yellow
                Change-ServiceStartMode $svc
                Start-Sleep 1
            }
			"6" {
				if (-not $Global:NssmInfo -or [string]::IsNullOrEmpty($Global:NssmInfo.Path) -or -not (Test-Path $Global:NssmInfo.Path)) {
					Write-Warning "$($L.nssmNotInstalledPrompt)"
					Read-Host
					continue
				}
                Write-Host "`n$($L.openingServiceEditInterface)" -ForegroundColor Yellow
                Start-Process "nssm.exe" -ArgumentList "edit $($svc.Name)" -Wait
				Write-Host "`n$($L.serviceEditComplete)" -ForegroundColor Yellow
                Start-Sleep 1
            }
            "7" { 
				Remove-ServiceWithConfirmation $svc.Name; 
				Initialize-Services
				return 
			}
            { $Global:ExitKeys -contains $_ } { return }
            default {
                continue
            }
        }
    }
}

# ---------------------------
# 显示详细参数
# ---------------------------
function Show-ServiceDetails {
    param($svc)
    $paramsPath = "HKLM:\SYSTEM\CurrentControlSet\Services\$($svc.Name)\Parameters"
    $params = if (Test-Path $paramsPath) { Get-ItemProperty $paramsPath } else { @{} }
    Write-Host "`n=== $($L.serviceDetails) ===" -ForegroundColor Cyan
    Write-Indented "Name         : $($svc.Name)"
    Write-Indented "DisplayName  : $($svc.DisplayName)"
    Write-Indented "Description  : $($svc.Description)"
    Write-Indented "State        : $($svc.State)"
    Write-Indented "StartMode    : $($svc.StartMode)"
    Write-Indented "PathName     : $($svc.PathName)"
    Write-Indented "Application  : $($params.Application)"
    Write-Indented "Parameters   : $($params.AppParameters)"
    Write-Indented "Directory    : $($params.AppDirectory)`n"
}

# ---------------------------
# 更改启动类型
# ---------------------------
function Change-ServiceStartMode {
    param($svc)
    Write-Host "`n$($L.currentStartupType) $($svc.StartMode)"
    Write-Indented "1. Automatic"
    Write-Indented "2. Manual"
    Write-Indented "3. Disabled"
    $choice = Read-Host "`n$($L.selectNewStartupType)"
    $mode = switch ($choice) {
        "1" { "Automatic" }
        "2" { "Manual" }
        "3" { "Disabled" }
        default { Write-Host "$($L.invalidChoice)"; return }
    }
    Set-Service $svc.Name -StartupType $mode
    Write-Host "$($L.startupTypeUpdated) $mode" -ForegroundColor Green
}

# ---------------------------
# 删除服务确认
# ---------------------------
function Remove-ServiceWithConfirmation {
    param($svcName)

    Write-Host
    $confirm1 = Read-Host "$svcName $($L.confirmDeleteService)"
    if ($confirm1 -notmatch "^[Yy]$") {
        Write-Host "$($L.cancelDelete)" -ForegroundColor Yellow
        Start-Sleep 1
        return
    }
	Write-Warning "$($L.deleteServiceWarning)"
    $confirm2 = Read-Host "$svcName $($L.confirmReallyDelete)"
    if ($confirm2 -notmatch "^[Yy]$") {
        Write-Host "$($L.cancelDelete)" -ForegroundColor Yellow
        Start-Sleep 1
        return
    }
    Write-Host "$($L.deletingService)" -ForegroundColor Yellow
    # 检查服务是否存在
    $service = Get-Service -Name $svcName -ErrorAction SilentlyContinue
    if (-not $service) {
        Write-Host "$($L.serviceNotExist)$svcName" -ForegroundColor Red
        Start-Sleep 2
        return
    }
    # 如果服务在运行，先停止它
    if ($service.Status -eq 'Running') {
        Write-Host "$($L.serviceRunningStopping)" -ForegroundColor Yellow
        Stop-Service -Name $svcName -Force -ErrorAction SilentlyContinue
        $service.WaitForStatus('Stopped', '00:00:10')  # 最多等待10秒
    }
    # 使用 nssm 删除服务
    & nssm.exe remove $svcName confirm
    Write-Host "$($L.serviceDeleted)" -ForegroundColor Green
    Start-Sleep 3
}

# ---------------------------
# 添加 NSSM 服务 GUI
# ---------------------------
function Add-NssmService {
	if (-not $Global:NssmInfo -or [string]::IsNullOrEmpty($Global:NssmInfo.Path) -or -not (Test-Path $Global:NssmInfo.Path)) {
		Write-Warning "$($L.nssmNotInstalledPrompt)"
		Read-Host
		return
	}
	Clear-Host
	Start-Process nssm.exe -ArgumentList "install"
}

function Show-ThemeMenu {
    while ($true) {
        Clear-Host
		$themes = Join-Path $PSScriptRoot $Global:ThemeDir
        # 确认主题目录存在
        if (-not (Test-Path $themes)) {
            Write-Warning "$($L.themeFolderNotFound) $($themes)"
			Read-Host
            return
        }

        # 获取所有 ps1 主题文件
        $themeFiles = Get-ChildItem -Path $themes -Filter *.ps1 | Sort-Object Name
        if ($themeFiles.Count -eq 0) {
            Write-Warning "$($L.themeFileNotFoundInFolder)"
			Read-Host
            return
        }

        # 构建菜单
        $menuItems = @($null)
        $i = 1
        foreach ($file in $themeFiles) {
            $menuItems += @{Text  = "$i. $($file.BaseName)"}
            $i++
        }
        $menuItems += @(
            $null,
            @{Text="0. $($L.resetToDefaultTheme)"},
            @{Text="q. $($L.return)";;Color=$Global:UI.MutedColor;},
            $null
        )

        # 显示菜单
        Show-BoxMenu -Title "$($L.chooseTheme)" -MenuItems $menuItems
        # 读取用户选择
        $selection = Read-MenuChoice
        # 处理退出
        if ($selection -eq 'q') {
            return
        }
		$ThemeDir = Join-Path $PSScriptRoot $Global:ThemeDir
		$ThemeSaveFile = Join-Path $ThemeDir "current_theme.txt"

        # 重置默认主题
        if ($selection -eq '0') {
            Remove-Item $ThemeSaveFile -Force -ErrorAction SilentlyContinue
			Load-SavedTheme
			continue
        }
		
        # 选择主题文件
        if ($selection -match '^\d+$' -and $selection -ge 1 -and $selection -le $themeFiles.Count) {
            $selectedTheme = $themeFiles[$selection - 1].FullName
            . $selectedTheme
			Set-Content -Path $ThemeSaveFile -Value $themeFiles[$selection - 1].BaseName -Encoding UTF8
        }
    }
}

# ---------------------------
# 添加服务菜单
# ---------------------------
function Add-ServiceMenu {
    while ($true) {
		Clear-Host
		$menuItems = @(
			$null,
			@{Text="1. $($L.addNssmService)"},
			@{Text="2. $($L.addCustomServices)$($Global:ServiceFile)"},
			$null,
			@{Text="0. $($L.returnMainMenu)";Color=$Global:UI.MutedColor;}
		)
		Show-BoxMenu -Title "$($L.addService)" -MenuItems $menuItems
        $choice = Read-MenuChoice
        switch ($choice) {
            "1" { Add-NssmService }
            "2" { Edit-ServiceFile }
            { $Global:ExitKeys -contains $_ } { return }
        }
    }
}

# ---------------------------
# 设置菜单
# ---------------------------
function Show-SettingsMenu {
    while ($true) {
		Clear-Host
		$menuItems = @(
			$null,
			@{Text="1. $($L.installNssm)"},
			@{Text="2. $($L.toggleTheme)"}
			$null,
			@{Text="0. $($L.returnMainMenu)";Color=$Global:UI.MutedColor;}
		)
		Show-BoxMenu -Title "$($L.setting)" -MenuItems $menuItems
        $choice = Read-MenuChoice
        switch ($choice) {
            "1" { Install-NssmIfMissing }
			"2" { Show-ThemeMenu }
            { $Global:ExitKeys -contains $_ } { return }
        }
    }
}

# ---------------------------
# 关于
# ---------------------------
function Show-AboutMenu {
    Clear-Host
    $oldWidth = $Global:UI.Width
    $Global:UI.Width = 84

    # 直接循环处理 JSON 里的每一行
    $menuItems = foreach ($lineTemplate in $L.aboutBody) {
        # 1. 变量替换（-f 会自动处理每一行的占位符）
        $formattedLine = $lineTemplate -f $L.appName, 
                                          $Global:ScriptVersion, 
                                          $Global:ScriptUser, 
                                          $Global:ScriptUpdate, 
                                          $PSCommandPath, 
                                          $Global:NssmInfo.Version, 
                                          $Global:NssmInfo.Path,
                                          $Global:ServiceFile

        # 2. 根据内容决定颜色
        if ($formattedLine -match '[\[【]') {
            @{ Text = $formattedLine; Color = 'Cyan' }
        } elseif ($formattedLine -match '^\s*\*') {
            @{ Text = $formattedLine; Color = 'Yellow' }
        } else {
            @{ Text = $formattedLine }
        }
    }

    # 3. 渲染
    Show-BoxMenu -Title "$($L.about)" -MenuItems $menuItems -Footer "$($L.returnMainMenu)" -Wrap
    
    Read-Host
    $Global:UI.Width = $oldWidth
}

# ---------------------------
# 主菜单
# ---------------------------
function Show-MainMenu {
    while ($true) {
		Clear-Host
		$menuItems = @(
			$null,
			@{Text="1. $($L.serviceList)"},
			@{Text="2. $($L.addService)"},
			@{Text="3. $($L.setting)"},
			@{Text="4. $($L.about)"},
			$null,
			@{Text="0. $($L.exit)";Color=$Global:UI.MutedColor;}
		)
		Show-BoxMenu -Title "$($L.appName)" -MenuItems $menuItems -Footer "$Global:ScriptVersion "
		$choice = Read-MenuChoice
		switch ($choice) {
			"1" { Show-ServiceListMenu }
			"2" { Add-ServiceMenu }
			"3" { Show-SettingsMenu }
			"4" { Show-AboutMenu }
			{ $Global:ExitKeys -contains $_ } { Exit }
		}
	}
}

# ===========================
# 启动应用
# ===========================
function Main{
	Ensure-RunAsAdmin
	Initialize-Parameters
	Load-SavedTheme
	Initialize-Services
	Initialize-Nssm
	Show-MainMenu
}
Main
