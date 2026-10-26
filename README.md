# EzOffice — Deployment & Activation Utility (ODT)

> **Lifetime Activation Note:**  
> **ODT** allows you to deploy and permanently activate Microsoft Office and Windows. If your system or Office suite ever becomes deactivated in the future (due to a Windows update or license reset), simply open **`Office.bat`**, select **Activate Office** or **Activate Windows**, and click **RUN SELECTED PROCESS** to instantly reactivate your system. This lifetime activation functionality works due to **Massgrave**.

## 📢 Acknowledgments & Credits

Special thanks and credit go to the [**Massgrave (MAS / Microsoft Activation Scripts)**](https://massgrave.dev) open-source project. **ODT** relies on their activation framework and tools:

* **OHook Engine:** Permanent Office activation.
* **HWID Engine:** Permanent Windows Digital License activation bound to your hardware ID.
* **Edition Managers:** Interactive scripts powered by Massgrave for **Changing/Adding Office Editions** and **Changing Windows Editions**.

## 1. How to Setup and Run

### Step 1: Download the Script

1. Download or save the script as **`Office.bat`** (or **`Office.cmd`**).
2. Ensure it retains the `.bat` or `.cmd` extension.

### Step 2: Launch the Interface

1. Right-click **`Office.bat`** and select **Run as administrator** (if launched normally, the utility will automatically prompt for administrator permission).
2. From the Graphical User Interface (GUI), choose your desired task:

| Task | Description |
| --- | --- |
| **Install Office** | Select your Office edition, architecture (64-bit/32-bit), language, update channel, optional product keys, and specific apps. |
| **Change / Add Office Edition** | Opens Massgrave's menu to change or add extra Office editions without starting from scratch. |
| **Change Windows Edition** | Opens Massgrave's menu to upgrade your Windows edition (e.g., Home to Pro), activates it, and offers to restart. |
| **Activate Office** | Runs Massgrave's `/OHook` activation script to activate Office. |
| **Activate Windows** | Runs Massgrave's `/HWID` script to activate Windows 10/11. |
| **Check System Activation Status** | Displays your OS version, Windows license state, and Office activation status. |

3. Click **RUN SELECTED PROCESS**.
4. Track progress live via the text box at the bottom of the window.

## 2. Key Features

* **Dynamic Office Deployment:** Downloads official Office files directly from Microsoft using the official Office Deployment Tool (`setup.exe`).
* **Wide Product Support:** Supports Microsoft 365, Office 2024 / 2021 Perpetual (LTSC & Retail), Visio, Project, and standalone apps.
* **App Selection:** Pick only the apps you need (e.g., Word and Excel only, without Teams or OneDrive).
* **Architecture & Language Options:** Supports 64-bit/32-bit and over 30 languages.
* **Permanent Activation:** Uses Massgrave's `/OHook` and `/HWID` activation scripts.
* **Windows & Office Edition Switching:** Easily upgrade Windows or add Office products using Massgrave tools.
* **Fast Diagnostic Check:** Quickly inspects your system to show if Windows and Office are activated.
* **Optional Cleanup:** At the end of installation, you can choose to keep the temporary folder (`C:\ODT`) if you want to reinstall later using the exact same configuration, though deleting it is recommended to free up disk space.

## 3. Technical Specifications

### Architecture & Runtime

* **Executable Package:** Polyglot script (`.bat` wrapper encapsulating PowerShell 5.1 host script).
* **Privilege Elevation:** Auto-elevates via PowerShell `Start-Process -Verb RunAs` if executed without administrator rights.
* **GUI Engine:** Built natively using `.NET Framework` (`System.Windows.Forms` & `System.Drawing`) without external dependencies.
* **Networking Protocol:** Enforces TLS 1.2 / HTTPS communication (`[Net.SecurityProtocolType]::Tls12`) for security and CDN updates.

### Staging & File Management

* **Working Staging Folder:** `C:\ODT`
* **Installer Verification:** Downloads Microsoft's official `setup.exe` (Office Deployment Tool) and verifies its digital signature (`Microsoft Corporation`).
* **Antivirus Interoperability:** Automatically adds a temporary Defender folder exclusion to `C:\ODT` during execution to prevent deployment blockages.
* **Folder Cleanup Policy:** User-configurable at completion. You may keep `C:\ODT` to quickly reinstall later with identical configuration XML settings, though deletion is recommended to save storage space.

### Activation & Deployment Core

* **Office Installation Engine:** Dynamic `configuration.xml` generation paired with Microsoft Office Click-to-Run (C2R) CDN setup.
* **Office Activation Core:** Massgrave permanent Ohook DLL method (`/OHook`).
* **Windows Activation Core:** Massgrave permanent Digital License / Hardware ID activation (`/HWID`).
* **Process Tracking Engine:** Real-time polling of `OfficeC2RClient.exe` and `setup.exe` installer processes (up to 60-minute safety threshold).

## 4. System Requirements

* **Supported Operating Systems:** Windows 10 or Windows 11 (64-bit / 32-bit).
* **Privileges:** Administrator access (UAC prompt).
* **Connectivity:** Active internet connection (HTTPS TLS 1.2) for CDN package fetch and license validation.
