English | [中文](./README_zh-CN.md) 

# Windows Service Manager (WSM)

WSM (Windows Service Manager) is an interactive command-line tool based on PowerShell + [NSSM](https://nssm.cc/) for unified management of Windows services. It is especially suitable for managing **custom services hosted by NSSM** (EXE / BAT / JAR / Python, etc.).

This tool provides a complete menu interface, supporting service viewing, starting, stopping, restarting, deleting, editing, theme switching, and automatic installation of NSSM.

## Example Usage

![](./imgs/service_list.png)

![](./imgs/service_manage.png)

---

## ✨ Features

* Automatically detect and manage **NSSM-hosted services**
* Manage custom services added via `services.txt`
* Service operations:

  * Start / Stop / Restart
  * Delete (double confirmation to prevent accidental deletion)
  * Modify startup type (Automatic / Manual / Disabled)
* View detailed service parameters:

  * Program path
  * Startup arguments
  * Working directory
* One-click access to **NSSM official GUI** for service editing
* Automatically detect and **install NSSM online**
* Theme switching (supports custom theme scripts)
* Automatic elevation to administrator privileges
* Adapted for both Chinese and English displays with proper alignment

---

## 🧩 System Requirements

* Windows 10 / 11
* PowerShell 5.1
* Administrator privileges (script will request automatically)
* Internet access (only required when downloading NSSM)

---

## 📁 Directory Structure

```text
WSM.ps1                 # Main script
services.txt            # Custom service list to manage
plugins/
 └─ Box.ps1             # Console Box UI plugin
themes/
 ├─ xxx.ps1             # Theme file
 └─ current_theme.txt   # Current theme record file (automatically generated)
```

---

## 🚀 Usage

### 1️⃣ Run the Script

This script supports both **cmd** and **ps1** formats. `WSM.cmd` will run `WSM.ps1` with administrator privileges.

---

#### Method 1: Double-click to Run (Recommended)

Double-click the `WSM.cmd` file to run it.

---

#### Method 2: Run in PowerShell

First, disable the PowerShell script execution policy:

```
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Bypass
```

Then, run `WSM.ps1` in PowerShell with administrator privileges.

---

### 2️⃣ Main Menu Features

* **Service List**:

  * View and manage all NSSM-hosted services and those defined in `services.txt`
* **Add New Service**:

  * Open NSSM official GUI (`nssm install`)
  * Edit custom service list in `services.txt`
* **Program Settings**:

  * Install NSSM
  * Switch themes
* **About**:

  * Display program information, version, NSSM details, and feature descriptions

---

## 📄 `services.txt` Description

`services.txt` is used to supplement additional services that need to be managed. (Services added via NSSM will be managed automatically.)

Example:

```text
# Add services to monitor here
W32Time
MySQL
```

Explanation:

* One service name per line (Service Name)
* Supports comments (`#` / `;` / `//`)
* If the file does not exist, it will be automatically created

---

## 🎨 Theme System Description

* All theme files are located in the `themes` directory
* Each theme is a `.ps1` file
* The current theme is recorded in `current_theme.txt`
* Runtime switching is supported without restarting the script

Example theme configuration:

```powershell
# ==========================
# UI Theme Configuration
# ==========================
# The $Global:UI hashtable controls the style and display effects of the WSM menu interface
# Themes can be switched at runtime, supporting modification of color, border style, width, etc.
# Field explanation:
#   Width           : Menu/Box width (in characters)
#   BorderColor     : Border color
#   BoxStyle        : Border style (Double/Single/Heavy/Rounded/Ascii/Dotted)
#   TitleColor      : Menu title color
#   TextColor       : List text color
#   TextPaddingLeft : Left padding for text (spaces)
#   AccentColor     : Highlight or emphasized text color (not used yet)
#   MutedColor      : Auxiliary or prompt text color (like "Back" button, secondary information)
$Global:UI = @{
    Width       = 50
    BorderColor = 'DarkGray'
    BoxStyle    = 'Heavy'
    TitleColor  = 'DarkYellow'
    TextColor   = 'Cyan'
    TextPaddingLeft = 2
    AccentColor = 'DarkYellow'
    MutedColor  = 'Gray'
}
```

---

## ⚙️ NSSM Support Details

* The program will automatically detect the `nssm.exe` in the current directory and the `PATH` environment variable.
* If NSSM is not found, it will not be forcefully installed. You can still manage services added in `services.txt`, but the service editing features and the ability to add new services via NSSM will not be available.
* You can manually download NSSM via the **Settings** menu if needed.
* Supports adding NSSM to the system `PATH` during installation.

NSSM website: [nssm](https://nssm.cc)

---

## ⚠️ Important Notes

* This tool **must be run with administrator privileges**.
* Service deletion is irreversible, please confirm carefully.
* Before editing a service, ensure that the program path and arguments are correct.
* Modifying the `PATH` variable may affect system environment variables.

---

## 📌 Use Cases

* Registering regular programs as Windows services
* Managing background services for Java / Python / Node / Batch
* Unified service management for operations or development environments
* Replacing manual usage of `services.msc` or `sc.exe`

---

## 👤 Author Information

* Author: zjw
* Project homepage: [wsm](https://github.com/Zhu-junwei/wsm)

---

## 📜 License

This project (WSM) is licensed under the MIT License, allowing free use, copying, modification, and distribution of the code for personal or commercial purposes.

---

Let me know if you need any adjustments!
