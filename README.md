# SSBPAC Traceability Generator — Web Application

A self-contained web application for generating dynamic, multi-stage SWRS Traceability Reports with support for complete input folders.

---

## How to Share with Others / Public Access

### Method 1: Share on the Local Network / Office Wi-Fi (No Internet Needed)
1. Double-click `start_app.bat` as usual.
2. The console window will display your network address:
   ```
   Local Network (Same Wi-Fi): http://192.168.1.74:8080
   ```
3. Any colleague or device connected to the same Wi-Fi, office network, or company VPN can immediately open that URL in their browser!

---

### Method 2: Fixed / Static Worldwide Public HTTPS Link
If colleagues or clients are outside your local network:
1. Make sure `start_app.bat` is running.
2. Double-click **`share_public_tunnel.bat`**.
3. It defaults to a fixed static subdomain:
   ```
   https://ssbpac-traceability.loca.lt
   ```
   (You can press Enter to keep this permanent URL, or type any custom name you prefer).
4. Send that URL to anyone — it stays the same every time!
   *(Note: The first time someone opens a loca.lt URL, it asks for a 'Tunnel Password'. This is simply your public IP address, which you can see at https://loca.lt/mytunnelpassword).*

---

### Method 3: Cloudflare Tunnel (Alternative)
Run from PowerShell:
```powershell
npx cloudflared tunnel --url http://localhost:8080
```
This gives an instant free `https://...trycloudflare.com` URL.

---

## Supported Input Stages

| Stage | Input Type | Description |
|---|---|---|
| **1. Master SWRS Requirements** | `.xlsx` File or Folder | **MANDATORY**. Base requirements template with Section, SL No, Requirement, and ID's. |
| **2. SHI Verification Reports** | Complete Folder or `.docx` Files | **OPTIONAL**. Automatically traverses `Single line` and `Double line` subfolders. |
| **3. Overall System Software Test** | Complete Folder or `.docx` Files | **OPTIONAL**. System testing report folder. |
| **4. Software Component Testing** | Complete Folder or `.tcf` Files | **OPTIONAL**. Unit testing folder (traverses `BVA`, `Statement_Coverage`, etc.). |
| **5. Software Integration Testing** | Complete Folder or `.tcf` Files | **OPTIONAL**. Integration testing folder with `.tcf` test suites. |

---

## Dynamic Matrix Generation

- **SWRS vs. SHI Only (6 Columns)**:
  - When only the SHI folder is active or provided:
    - Columns 1–4: `Section`, `SL No.`, `Requirement`, `ID's`
    - Columns 5–6: `SHI` (`Single line`, `Double line`)
- **SWRS vs. Full Suite (16 Columns)**:
  - When all folders are active, generates the complete 16-column matrix across all 4 stages.
- **Custom Combinations**:
  - Check or uncheck any stage checkbox (or choose/clear any folder) to customize the exact matrix columns to produce.

---

## How to Run

### Double-Click Launcher (Easiest)
Simply double-click:
```
start_app.bat
```
This starts the Python server and opens your browser to:
👉 **`http://localhost:8080`**
