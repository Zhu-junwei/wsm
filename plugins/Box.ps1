<#
.SYNOPSIS
    控制台美化工具与菜单渲染脚本（支持中英文对齐、边框样式、文本裁剪与换行）。

.DESCRIPTION
    本脚本提供了一套用于 PowerShell 控制台的 UI 渲染函数，可用于：
      - 绘制带边框的文本框
      - 显示菜单和子菜单
      - 支持中英文混排的宽度计算
      - 自动裁剪或换行文本
      - 多种边框样式（Double / Single / Heavy / Rounded / Ascii）
      - 可设置文字颜色、边框颜色、对齐方式和文本缩进

    核心功能模块：
      1. 全局 UI 配置 ($Global:UI)
         - 控制宽度、边框颜色、标题颜色、文本颜色等
      2. 边框样式定义 ($Global:BoxStyles)
         - 支持多种字符边框风格
      3. 字符串显示宽度计算 (Get-DisplayWidth)
         - 解决中英文字符宽度不一致问题
      4. 文本裁剪与填充 (Fit-Text)
         - 可设置对齐方式（Left/Center/Right）和换行
      5. 边框绘制 (Write-BoxBorder)
      6. 内容行绘制 (Write-BoxLine)
      7. 菜单渲染 (Show-BoxMenu)
         - 可自定义标题、菜单项、页脚和对齐方式
         - 支持菜单项颜色和换行显示

.PARAMETER Title
    菜单或文本框标题。

.PARAMETER MenuItems
    菜单项数组，每个菜单项可以包含：
      - Text  : 菜单文本
      - Color : 文本颜色（可选）
      - Align : 对齐方式（Left/Center/Right，可选）

.PARAMETER Footer
    页脚文本，可显示提示或版权信息。

.EXAMPLE
    # 显示一个简单菜单
    $items = @(
        @{ Text = '1. 启动服务'; Color = 'Green' },
        @{ Text = '2. 停止服务'; Color = 'Red' },
        @{ Text = '3. 退出'; Color = 'Yellow' }
    )
    Show-BoxMenu -Title '服务控制菜单' -MenuItems $items -Footer '请选择操作'

.NOTES
    - 作者: zjw
    - 创建日期: 2025-12-16
    - 版本: 1.0
    - 依赖: PowerShell 5.1 及以上
    - 适用场景: 控制台脚本、自动化工具、运维脚本、交互式菜单
#>

# ==========================================================
# 全局 UI 配置
# ==========================================================
# $Global:UI 保存全局界面风格设置
# Width           : 控制台内容宽度（字符数）
# BorderColor     : 边框颜色
# BoxStyle        : 默认边框样式（Double/Single/Heavy/Rounded/Ascii/Dotted）
# TitleColor      : 标题文本颜色
# TextColor       : 普通文本颜色
# TextPaddingLeft : 左右内边距
# MutedColor      : 页脚或辅助文本颜色
# AccentColor     : 高亮文本颜色
$Global:UI = @{
	Width       = 50
	BorderColor = 'DarkCyan'
	BoxStyle    = 'Rounded'
	TitleColor  = 'Cyan'
	TextColor   = 'White'
	TextPaddingLeft = 2
	MutedColor  = 'DarkGray'
	AccentColor = 'Cyan'
}
$Global:UI_Default = $Global:UI.Clone()
# ==========================================================
# 重置边框样式为默认
# ==========================================================
function Reset-BoxStyle {
    $Global:UI.BoxStyle = $Global:UI_Default.BoxStyle
}

# ==========================================================
# 边框样式定义
# ==========================================================
# $Global:BoxStyles 定义多种边框样式
# TL/TR/BL/BR : 左上/右上/左下/右下角符号
# H/V         : 水平/垂直线符号
# ML/MR       : 中间分隔符左右符号
$Global:BoxStyles = @{
    Double   = @{TL = '╔'; TR = '╗'; BL = '╚'; BR = '╝'; H  = '═'; V  = '║'; ML = '╠'; MR = '╣'}
    Single   = @{TL = '┌'; TR = '┐'; BL = '└'; BR = '┘'; H  = '─'; V  = '│'; ML = '├'; MR = '┤'}
    Heavy    = @{TL = '┏'; TR = '┓'; BL = '┗'; BR = '┛'; H  = '━'; V  = '┃'; ML = '┣'; MR = '┫'}
    Rounded  = @{TL = '╭'; TR = '╮'; BL = '╰'; BR = '╯'; H  = '─'; V  = '│'; ML = '├'; MR = '┤'}
    Ascii    = @{TL = '+'; TR = '+'; BL = '+'; BR = '+'; H  = '-'; V  = '|';ML = '+'; MR = '+'}
    Dotted   = @{TL = '┌'; TR = '┐'; BL = '└'; BR = '┘'; H = '┄'; V = '┆'; ML = '├'; MR = '┤'}
}

# ==========================================================
# 对齐方式
# ==========================================================
enum TextAlign { Left; Center; Right }

function Get-BoxStyle {
    $Global:BoxStyles[$Global:UI.BoxStyle]
}

# ==========================================================
# 【核心】计算字符串显示宽度（中英文对齐关键）
# ==========================================================
# 返回整数宽度
# 注意：CJK字符及全角符号宽度为2，其他字符为1
function Get-DisplayWidth {
    param([string]$Text)
    $width = 0
    foreach ($ch in $Text.ToCharArray()) {
        # CJK 统一表意字符 + 全角符号 → 宽度 2
        if ($ch -match '[\u1100-\u115F\u2E80-\uA4CF\uAC00-\uD7A3\uF900-\uFAFF\uFE10-\uFE6F\uFF00-\uFF60]') {
            $width += 2
        }
        else {
            $width += 1
        }
    }
    return $width
}

# ==========================================================
# 内容裁剪 + 填充（基于显示宽度）
# ==========================================================
# 功能：
#   - 根据指定宽度裁剪或换行文本
#   - 支持文本对齐（Left/Center/Right）
#   - 支持 Wrap 换行
# 参数：
#   - Text : 要显示的文本
#   - Width: 显示宽度
#   - Align: 对齐方式
#   - Wrap : 是否换行
# 返回值：
#   - 字符串数组，每个元素是一行文本
function Fit-Text {
    param(
        [string]$Text, 
        [int]$Width, 
        [TextAlign]$Align = [TextAlign]::Left,
        [switch]$Wrap
    )

    $lines = @()
    $padding = if ($Global:UI.TextPaddingLeft) { $Global:UI.TextPaddingLeft } else { 0 }
    $usableWidth = $Width - $padding * 2
    $remaining = $Text

    while ($remaining) {
        $trimmed = ''
        $used = 0

        foreach ($ch in $remaining.ToCharArray()) {
            $w = if ($ch -match '[\u1100-\u115F\u2E80-\uA4CF\uAC00-\uD7A3\uF900-\uFAFF\uFE10-\uFE6F\uFF00-\uFF60]') { 2 } else { 1 }
            if ($used + $w -gt $usableWidth) { break }
            $trimmed += $ch
            $used += $w
        }

        # 不换行时超过宽度用 '..'
        if (-not $Wrap -and $remaining.Length -gt $trimmed.Length) {
            $finalTrim = ''
            $trimmedLength = 0
            foreach ($ch in $remaining.ToCharArray()) {
                $w = if ($ch -match '[\u1100-\u115F\u2E80-\uA4CF\uAC00-\uD7A3\uF900-\uFAFF\uFE10-\uFE6F\uFF00-\uFF60]') { 2 } else { 1 }
                if ($trimmedLength + $w -gt ($usableWidth - 2)) { break }
                $finalTrim += $ch
                $trimmedLength += $w
            }
            $trimmed = $finalTrim + '..'
            $used = $usableWidth
            $remaining = ''
        } else {
            $remaining = $remaining.Substring($trimmed.Length)
        }

        $space = $usableWidth - $used
        switch ($Align) {
            'Right'  { $line = (' ' * $space) + $trimmed }
            'Center' { $line = (' ' * [Math]::Floor($space/2)) + $trimmed + (' ' * ([Math]::Ceiling($space/2))) }
            default  { $line = $trimmed + (' ' * $space) }
        }
        $lines += (' ' * $padding) + $line + (' ' * $padding)
        if (-not $Wrap) { break }
    }
	if ($lines.Count -eq 0) {
        $lines += ' ' * $Width
    }
    return $lines
}

# ==========================================================
# 边框绘制
# ==========================================================
# 绘制顶部/中间分隔线/底部边框
# 参数：
#   Type : Top / Sep / Bottom
# 使用 $Global:UI.BorderColor 进行着色
function Write-BoxBorder {
    param([ValidateSet('Top','Sep','Bottom')]$Type)

    $s = Get-BoxStyle
    $c = $Global:UI.BorderColor
    $w = $Global:UI.Width

    $line = switch ($Type) {
        'Top'    { "{0}{1}{2}" -f $s.TL, ($s.H * $w), $s.TR }
        'Sep'    { "{0}{1}{2}" -f $s.ML, ($s.H * $w), $s.MR }
        'Bottom' { "{0}{1}{2}" -f $s.BL, ($s.H * $w), $s.BR }
    }
    Write-Host $line -ForegroundColor $c
}

# ==========================================================
# 内容行（边框颜色 / 字体颜色完全分离）
# ==========================================================
# 绘制一行带左右边框的文本
# 参数：
#   Text      : 显示内容
#   TextColor : 文本颜色（可选）
#   Align     : 左/中/右对齐
#   Wrap      : 是否换行
function Write-BoxLine {
    param(
        [string]$Text = '',
        [ConsoleColor]$TextColor = $Global:UI.TextColor,
        [TextAlign]$Align = [TextAlign]::Left,
        [switch]$Wrap
    )

    $s = Get-BoxStyle
    $width = $Global:UI.Width
    $lines = Fit-Text $Text $width $Align -Wrap:$Wrap

    foreach ($line in $lines) {
        Write-Host $s.V -ForegroundColor $Global:UI.BorderColor -NoNewline
        Write-Host $line -ForegroundColor $TextColor -NoNewline
        Write-Host $s.V -ForegroundColor $Global:UI.BorderColor
    }
}

# ==========================================================
# 菜单渲染
# ==========================================================
# 功能：
#   - 渲染标题、菜单项、页脚
#   - 支持每项自定义颜色和对齐
# 参数：
#   Title       : 菜单标题
#   MenuItems   : 菜单项数组，每项支持 Text/Color/Align
#   Footer      : 页脚文本
#   TitleAlign  : 标题对齐方式
#   FooterAlign : 页脚对齐方式
#   BoxStyle    : 边框样式
#   Wrap        : 是否换行显示
function Show-BoxMenu {
    param(
        [string]$Title = '',
        [array]$MenuItems = @(),
        [string]$Footer = '',
        [TextAlign]$TitleAlign = [TextAlign]::Center,
        [TextAlign]$FooterAlign = [TextAlign]::Right,
        [string]$BoxStyle = $Global:UI.BoxStyle,
        [switch]$Wrap 
    )
    # 保存当前样式
    $oldStyle = $Global:UI.BoxStyle
    # 如果调用时传了 BoxStyle，则临时覆盖
    if ($BoxStyle) { $Global:UI.BoxStyle = $BoxStyle }
	try {
		Write-Host ""
		Write-BoxBorder Top
		if ($Title) {
			Write-BoxLine $Title -TextColor $Global:UI.TitleColor -Align $TitleAlign
			Write-BoxBorder Sep
		}
		foreach ($item in $MenuItems) {
			if ($item -and $item.Text) {
				$color = if ($item.Color) { $item.Color } else { $Global:UI.TextColor }
				$align = if ($item.Align) { $item.Align } else { [TextAlign]::Left }
				Write-BoxLine $item.Text -TextColor $color -Align $align -Wrap:$Wrap
			} else {
				Write-BoxLine ""
			}
		}
		if ($Footer) {
			Write-BoxLine "" # 空行分隔
			Write-BoxLine $Footer -TextColor $Global:UI.MutedColor -Align $FooterAlign
		}
		Write-BoxBorder Bottom
		Write-Host ""
	} finally {
        # 恢复原来的 UI 样式
        $Global:UI.BoxStyle = $oldStyle
    }
}