# Windows Unused OEM Driver Remover

A simple **Python GUI tool** to list and remove **obsolete OEM drivers** from Windows. This tool helps identify unused drivers that could pose **security risks** or **block Windows upgrades**. It provides an interactive **sortable table** with **multi-select functionality** and allows easy removal of drivers using the **Windows pnputil command**.

## 🚀 Why Remove Unused Drivers?

### 🔒 Security Risks
- Old drivers can have **unpatched vulnerabilities** that could be exploited by attackers.
- Outdated drivers may **bypass security policies**, leading to **system instability or compatibility issues**.

### ⛔ Windows Upgrade Issues
- Some obsolete drivers **block Windows Feature Updates**, causing errors like:
  - *Windows cannot be installed because of an incompatible driver.*
  - *This PC can't be upgraded to Windows 11 due to outdated drivers.*
- Removing these drivers ensures **smoother system upgrades**.

---

## 🎯 Features

✅ **Lists all unused OEM drivers** (from `pnputil /enum-drivers`)  
✅ **Sortable table** (click column headers to sort)  
✅ **Multi-select support** (`CTRL + Click` to select multiple drivers)  
✅ **Select All / Deselect All buttons** for bulk actions  
✅ **Administrator privilege check** (prevents execution if not run as admin)  
✅ **Asynchronous driver loading** (GUI remains responsive)  
✅ **Functional scrollbar** for large driver lists  
✅ **Displays total number of unused drivers**  

---

## 📦 Installation

### Option 1: Running the Compiled `.exe`
1. **Download** `Unused-Drivers-Remover.exe` from the [Releases](https://github.com/ondrejvysek/Windows-Unused-OEM-Driver-Remover/blob/main/build/unused-drivers-remover.exe) page.
2. **Right-click → Run as Administrator**.
3. The tool will scan and display unused drivers.
4. Select unwanted drivers and click **Remove Selected Drivers**.

### Option 2: Running from Source (`Python`)
#### Requirements
- **Windows OS**
- **Python 3.8+**
- **Pnputil (Built into Windows)**

#### Install Required Libraries:
```bash
pip install pyinstaller
