# FREE-Microsoft-365 - Automated Installation & Activation Utility

> **Lifetime Activation Note:**  
> You can use this utility to get Microsoft Office for free forever. If Office or Windows ever becomes deactivated in the future (due to a Windows update or license reset), simply open **`Office.bat`**, select **Activate Office** or **Activate Windows**, and click **RUN SELECTED PROCESS** to instantly reactivate your system.

---

## 1. How to Setup and Run

### Step 1: Download the Script File
1. Download the **`Office.bat`** file to your computer.
2. Ensure the file is named **`Office.bat`** before running.

### Step 2: Run the Utility
1. Right-click **`Office.bat`** and select **Run as administrator** (or double-click it and accept the UAC prompt).
2. Choose your desired action from the menu interface:
   * **Install Office:** Downloads Microsoft 365 Apps for Enterprise (64-bit), configures default file settings, and permanently activates Office.
   * **Change / Add Office Edition:** Switch Office editions or add secondary components, followed by automatic re-activation.
   * **Change Windows Edition:** Upgrade or change your Windows SKU (e.g., Home to Pro) with automatic digital license (HWID) activation.
   * **Activate Office:** Runs a quick, non-destructive activation routine specifically for Microsoft Office using the permanent Ohook method.
   * **Activate Windows:** Runs a quick digital license (HWID) activation routine specifically for Windows 10/11.
3. Click **RUN SELECTED PROCESS**.
4. Monitor progress live in the integrated log panel.

---

## 2. Overview & Key Features

**Office** is a polyglot PowerShell and Windows Batch hybrid utility designed for automated installation, management, and activation of Microsoft Office and Windows operating systems.

- **Automated Office Installation:** Downloads, extracts, and installs Microsoft 365 Apps directly from Microsoft's Content Delivery Network (CDN).
- **Permanent Ohook Activation:** Employs the Ohook activation method for lifetime Office activation without altering local binaries.
- **Windows HWID Activation:** Permanently activates Windows 10/11 using official Microsoft Hardware ID digital licenses.
- **Interactive WinForms GUI:** Simple, lightweight graphical interface built directly into native Windows components.
- **Smart Process Polling:** Tracks background Click-to-Run installation processes to ensure activation occurs only after setup completes fully.
- **Automated Defender Exclusions:** Temporarily excludes the working folder (`C:\ODT`) to prevent security software false positives during installation.

---

## 3. Technical Specifications

| Feature / Parameter | Specification Details |
| :--- | :--- |
| **Tool Name** | Office |
| **Script Architecture** | Dual-Layer Polyglot (`.bat` Batch Wrapper + Embedded PowerShell 5.1 Host) |
| **GUI Framework** | `.NET Framework` (`System.Windows.Forms` & `System.Drawing`) |
| **Execution Directory** | `C:\ODT` (Staging & Installation Folder) |
| **Target Office Edition** | `O365ProPlusRetail` (Microsoft 365 Apps for Enterprise) |
| **Architecture** | 64-bit (`OfficeClientEdition="64"`) |
| **Update Channel** | `Current` Channel |
| **Office Activation Method** | Permanent Ohook Script Method (`/OHook`) |
| **Windows Activation Method** | Digital License Hardware ID (`/HWID`) |
| **Network Protocol** | TLS 1.2 / HTTPS Connection |

---

## 4. System Requirements

- **Supported Operating Systems:** Windows 10 (64-bit) or Windows 11 (64-bit).
- **Required Privileges:** Administrator rights (UAC elevation).
- **Connectivity:** Active internet connection (HTTPS TLS 1.2) to fetch installation files and activation tokens.
