# ===========================
# NSSM 服务管理菜单
# ===========================

# 检测管理员权限，如果不是管理员则以管理员重新运行
function Ensure-RunAsAdmin {
    $currentIdentity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentIdentity)
    if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        Write-Host "正在以管理员权限重新启动脚本..." -ForegroundColor Yellow
        $scriptPathEscaped = $PSCommandPath -replace '(["`])', '``$1'
        Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPathEscaped`"" -Verb RunAs
        Exit
    }
}
Ensure-RunAsAdmin

$BoxPath = Join-Path $PSScriptRoot 'plugins/Box.ps1'
if (Test-Path $BoxPath) { . $BoxPath }

function Initialize-Parameters() {
	$Global:ScriptName = "Windows服务管理(WSM)"
	$Global:ScriptUser = "zjw"
	$Global:ScriptVersion = "v1.2.0"
	$Global:ScriptUpdate = "20251217"
	$Global:ServiceFile = "services.txt"
	$Global:NssmInstallDir = "$env:ProgramFiles\NSSM"
	$Global:NssmZipUrl    = "https://nssm.cc/ci/nssm-2.24-103-gdee49fc.zip"
	$Global:ThemeDir = 'themes'
	$Global:ExitKeys = @("0", "q", "exit")
	$Global:WSMServiceStore = @()
	$Global:NssmInfo = [PSCustomObject]@{
        Path    = ''
        Version = ''
    }
}
function Initialize-DefaultTheme() {
	$Global:UI.Width = 50
	$Global:UI.BoxStyle = 'Heavy'
	$Global:UI.BorderColor = 'DarkCyan'
	$Global:UI.TextPaddingLeft = 2
	$Global:UI.TextColor  = 'Yellow'
	$Global:UI.MutedColor  = 'DarkGray'
}

function Load-SavedTheme {
	try {
		$ThemeDir = Join-Path $PSScriptRoot $Global:ThemeDir
		$ThemeSaveFile = Join-Path $ThemeDir "current_theme.txt"
		if (Test-Path $ThemeSaveFile) {
			$themeName = (Get-Content $ThemeSaveFile -Raw).Trim()
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
    param([string]$Prompt = "请选择操作")
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
        Write-Host "`n正在下载 NSSM..." -ForegroundColor Yellow
        Invoke-WebRequest -Uri $Global:NssmZipUrl -OutFile $tempZip -UseBasicParsing
        Write-Host "下载完成，正在解压..." -ForegroundColor Yellow

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
        if (-not $nssmExes -or $nssmExes.Count -eq 0) { throw "未找到 nssm.exe" }

        # 判断系统位数
        $is64 = [Environment]::Is64BitOperatingSystem
        $targetArch = if ($is64) { "win64" } else { "win32" }

        # 选择正确架构的 nssm.exe
        $sourceExe = $nssmExes | Where-Object { $_.Directory.Name -ieq $targetArch } | Select-Object -First 1
        if (-not $sourceExe) { throw "未找到 $targetArch\nssm.exe" }

        # 移动到安装目录根目录
        $destExe = Join-Path $Global:NssmInstallDir "nssm.exe"
        Move-Item -Path $sourceExe.FullName -Destination $destExe -Force
        Write-Host "NSSM 安装完成: $destExe" -ForegroundColor Green
		Write-Host
        $choice = Read-Host "是否将 NSSM 添加到系统 PATH? (y/N)"
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
				Write-Host "PATH 中已存在该目录" -ForegroundColor Yellow
			} else {
				$machinePath = [Environment]::GetEnvironmentVariable("PATH", "Machine")
				[Environment]::SetEnvironmentVariable("PATH", "$machinePath;$dir", "Machine")
				Write-Host "已成功添加到系统 PATH" -ForegroundColor Green
			}
		}
		Start-Sleep 1
        return $destExe
    } catch {
        Write-Host "下载或安装 NSSM 失败: $_" -ForegroundColor Red
        Read-Host "`n按回车退出程序"
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
        Write-Host "`nNSSM 已安装，路径：$nssmPath" -ForegroundColor Green
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
        Write-Host "NSSM 已经安装，路径：$($Global:NssmInfo.Path)" -ForegroundColor Green
		Read-Host
        return
    }

    Write-Host "未找到 NSSM (nssm.exe)，需要安装吗？" -ForegroundColor Yellow
    $choice = Read-Host "是否下载并安装 NSSM ? (y/N)"
    if ($choice -match "^[Yy]$") {
        $nssmPath = Install-Nssm
        $fullVersion = & $nssmPath version
        $Global:NssmInfo = [PSCustomObject]@{
            Path    = $nssmPath
            Version = $fullVersion
        }
        Write-Host "NSSM 安装完成，路径：$nssmPath" -ForegroundColor Green
        Start-Sleep 1
    } else {
        Write-Host "`n安装已取消" -ForegroundColor Red
        Read-Host "`n按回车返回"
    }
}

# ---------------------------
# 编辑服务文件
# ---------------------------
function Edit-ServiceFile {
	$ServiceFilePath = Join-Path $PSScriptRoot $Global:ServiceFile
    if (-not (Test-Path $ServiceFilePath)) {
        "# 这里添加需要监控的服务" | Out-File -FilePath $ServiceFilePath -Encoding UTF8
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
                @{Text="当前没有可管理的服务。"; Color='Yellow';},
				@{Text="仅管理使用nssm添加的服务和在$($Global:ServiceFile)中定义的服务。"; Color='Yellow';},
                $null,
                @{Text="0. 返回主菜单"; Color=$Global:UI.MutedColor;}
            )
            Show-BoxMenu -Title "服务列表" -MenuItems $menuItems -Wrap
        } else {
            # 计算服务名称列对齐长度
            $maxNameLength = ($services | ForEach-Object { Get-DisplayWidth $_.DisplayName } | Measure-Object -Maximum).Maximum
			$menuItems = @($null)
            $i = 1
            foreach ($svc in $services) {
				$nameText = "{0}. {1}" -f $i, (PadRightWidth $svc.DisplayName $maxNameLength)
                $menuItems += @{
                    Text  = $nameText + "  $($svc.State)"  # 状态紧跟名字
                    Color = (Get-StateColor $svc.State)    # 状态颜色
                    Align = 'Left'
                }
                $i++
            }

            # 添加空行和返回主菜单选项
            $menuItems += @(
				$null,
				@{Text="0. 返回主菜单"; Color=$Global:UI.MutedColor; Align='Left'},
				$null
			)

            Show-BoxMenu -Title "服务列表" -MenuItems $menuItems
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
        Start   = "启动"
        Stop    = "停止"
        Restart = "重启"
    }[$Action]

    Write-Host "`n正在${actionText}服务..." -ForegroundColor Yellow

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

        Write-Host "服务已${actionText}" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "`n服务${actionText}失败" -ForegroundColor Red
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
		Read-Host "`n按回车键继续"
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
			Write-Host "`n服务 $($svc.Name) 不存在，可能已被删除。" -ForegroundColor Red
			Read-Host
			return
		}
		Clear-Host
		$menuItems = @(
			$null,
			@{Text=" 当前状态 : $($svc.State)";Color=(Get-StateColor $svc.State)},
			$null,
			@{Text='1. 启动服务'},
			@{Text='2. 停止服务'},
			@{Text='3. 重启服务'},
			@{Text='4. 查看详细参数'},
			@{Text='5. 更改启动类型'},
			@{Text='6. 编辑服务'},
			@{Text='7. 删除服务'}
			$null,
			@{Text='0. 返回服务列表';Color=$Global:UI.MutedColor;},
			$null
		)
		Show-BoxMenu -Title "管理服务: $($svc.DisplayName)" -MenuItems $menuItems
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
            "4" { Show-ServiceDetails $svc; Read-Host "`n按回车返回" }
            "5" {
                Write-Host "`n正在更改启动类型..." -ForegroundColor Yellow
                Change-ServiceStartMode $svc
                Start-Sleep 1
            }
			"6" {
				if (-not $Global:NssmInfo -or [string]::IsNullOrEmpty($Global:NssmInfo.Path) -or -not (Test-Path $Global:NssmInfo.Path)) {
					Write-Warning "nssm未安装，请到设置中先安装nssm"
					Read-Host
					continue
				}
                Write-Host "`n正在打开服务编辑界面..." -ForegroundColor Yellow
                Start-Process "nssm.exe" -ArgumentList "edit $($svc.Name)" -Wait
				Write-Host "`n服务编辑完成..." -ForegroundColor Yellow
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
    Write-Host "`n=== 服务详细参数 ===" -ForegroundColor Cyan
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
    Write-Host "`n当前启动类型: $($svc.StartMode)"
    Write-Indented "1. Automatic"
    Write-Indented "2. Manual"
    Write-Indented "3. Disabled"
    $choice = Read-Host "`n选择新的启动类型"
    $mode = switch ($choice) {
        "1" { "Automatic" }
        "2" { "Manual" }
        "3" { "Disabled" }
        default { Write-Host "`n无效选择"; return }
    }
    Set-Service $svc.Name -StartupType $mode
    Write-Host "`n启动类型已更新为 $mode" -ForegroundColor Green
}

# ---------------------------
# 删除服务确认
# ---------------------------
function Remove-ServiceWithConfirmation {
    param($svcName)

    Write-Host
    $confirm1 = Read-Host "确认从系统中删除服务 $svcName ? (y/N)"
    if ($confirm1 -notmatch "^[Yy]$") {
        Write-Host "`n取消删除" -ForegroundColor Yellow
        Start-Sleep 1
        return
    }
    $confirm2 = Read-Host "!!! 警告 !!! 此操作不可撤销，真的要删除 $svcName ? (y/N)"
    if ($confirm2 -notmatch "^[Yy]$") {
        Write-Host "`n取消删除" -ForegroundColor Yellow
        Start-Sleep 1
        return
    }
    Write-Host "`n正在删除服务..." -ForegroundColor Yellow
    # 检查服务是否存在
    $service = Get-Service -Name $svcName -ErrorAction SilentlyContinue
    if (-not $service) {
        Write-Host "服务 $svcName 不存在" -ForegroundColor Red
        Start-Sleep 2
        return
    }
    # 如果服务在运行，先停止它
    if ($service.Status -eq 'Running') {
        Write-Host "服务正在运行，正在停止服务..." -ForegroundColor Yellow
        Stop-Service -Name $svcName -Force -ErrorAction SilentlyContinue
        $service.WaitForStatus('Stopped', '00:00:10')  # 最多等待10秒
    }
    # 使用 nssm 删除服务
    & nssm.exe remove $svcName confirm
    Write-Host "服务已删除" -ForegroundColor Green
    Start-Sleep 3
}

# ---------------------------
# 添加 NSSM 服务 GUI
# ---------------------------
function Add-NssmService {
	if (-not $Global:NssmInfo -or [string]::IsNullOrEmpty($Global:NssmInfo.Path) -or -not (Test-Path $Global:NssmInfo.Path)) {
		Write-Warning "nssm未安装，请到设置中先安装nssm"
		Read-Host
		return
	}
	Clear-Host
	Write-Host "=== 添加 NSSM 服务 ===`n"
	Start-Process nssm.exe -ArgumentList "install"
}

function Show-ThemeMenu {
    while ($true) {
        Clear-Host
		$themes = Join-Path $PSScriptRoot $Global:ThemeDir
        # 确认主题目录存在
        if (-not (Test-Path $themes)) {
            Write-Warning "主题文件夹 $($themes) 不存在"
			Read-Host
            return
        }

        # 获取所有 ps1 主题文件
        $themeFiles = Get-ChildItem -Path $themes -Filter *.ps1 | Sort-Object Name
        if ($themeFiles.Count -eq 0) {
            Write-Warning "主题文件夹中没有找到主题文件"
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
            @{Text="0. 重置为默认主题"},
            @{Text="q. 返回上级菜单";;Color=$Global:UI.MutedColor;},
            $null
        )

        # 显示菜单
        Show-BoxMenu -Title "主题选择" -MenuItems $menuItems
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
			@{Text='1. 添加nssm服务'},
			@{Text="2. 添加自定义服务列表$($Global:ServiceFile)"},
			$null,
			@{Text='0. 返回主菜单';Color=$Global:UI.MutedColor;}
		)
		Show-BoxMenu -Title "添加新服务" -MenuItems $menuItems
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
			@{Text='1. 安装nssm'},
			@{Text='2. 切换主题'}
			$null,
			@{Text='0. 返回上级菜单';Color=$Global:UI.MutedColor;}
		)
		Show-BoxMenu -Title "设置" -MenuItems $menuItems
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
	$width = $Global:UI.Width
	$Global:UI.Width = 80
    $content = @"
 【程序信息】
  名称        : $Global:ScriptName
  版本        : $Global:ScriptVersion
  作者        : $Global:ScriptUser
  更新时间    : $Global:ScriptUpdate
  位置        : $PSCommandPath
  主页        : https://github.com/Zhu-junwei/wsm
  NSSM 版本   : $($Global:NssmInfo.Version)
  NSSM 位置   : $($Global:NssmInfo.Path)
  NSSM 官网   : https://nssm.cc
 
 【功能】
  管理所有由 NSSM 托管的 Windows 服务
  管理所有由 $($ServiceFile) 定义的服务
  提供服务状态查看、启动、停止、重启、删除等操作
  调用 NSSM 官方 GUI 编辑服务，适合将普通程序（EXE / BAT / JAR / Python 等）注册为系统服务
  修改服务启动类型（Automatic / Manual / Disabled）
  查看服务详细参数（程序、参数、工作目录）
  支持检测并在线安装 NSSM，无需手动配置
  切换程序主题，可自定义添加
 
 【什么是 NSSM】
  NSSM（Non-Sucking Service Manager）用于将普通程序
  封装为标准 Windows Service，比 sc.exe 更稳定、易用。
 
 【注意事项】
  本脚本需以管理员权限运行
  删除服务操作不可恢复，请谨慎
  编辑服务前请确认程序路径与参数正确
"@
    $menuItems = $content -split "[`r`n]+" | ForEach-Object {
		if ($_ -match '【') {
			@{ Text = $_; Color = 'Cyan' }
		} else {
			@{ Text = $_ }
		}
	}

    # 显示 Box 菜单
    Show-BoxMenu -Title "关于" `
                 -MenuItems $menuItems `
                 -Footer "按回车返回主菜单" `
                 -BoxStyle $Global:UI.BoxStyle `
                 -Wrap

    Read-Host
	$Global:UI.Width = $width
}

# ---------------------------
# 主菜单
# ---------------------------
function Show-MainMenu {
    while ($true) {
		Clear-Host
		$menuItems = @(
			$null,
			@{Text='1. 服务列表'},
			@{Text='2. 添加服务'},
			@{Text='3. 设置'},
			@{Text='4. 关于'},
			$null,
			@{Text='0. 退出';Color=$Global:UI.MutedColor;}
		)
		Show-BoxMenu -Title "$Global:ScriptName" -MenuItems $menuItems -Footer "$Global:ScriptVersion "
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
# 启动脚本
# ===========================
function Main{
	Initialize-Parameters
	Load-SavedTheme
	Initialize-Services
	Initialize-Nssm
	Show-MainMenu
}
Main
