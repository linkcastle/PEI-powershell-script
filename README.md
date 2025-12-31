# PEI Investigation Script
**PowerShell Forensic Tool for Suspicious Activity Detection**

# Overview
This script searches for fraudulent or suspicious activity in:
- **Files** (PDFs, Office docs, images, Excel)
- **Outlook Emails** (Inbox/Sent Items)
- **Browser History** (Chrome/Edge)
- **OneNote** notes

Output is organized in structured folders for legal/compliance reviews.

---

# Requirements
- **Run as Administrator** (for full file access).
- **Outlook installed** (for email search).
- **PowerShell 5.1+** (tested on Windows 10/11).

---




# HOW-TO-USE


1. **Edit the config section** (lines 10-15):
   ```powershell
    $ userProfile = "C:\Users\NAME.SURNAME"  # Target user profile
    $ outputFolder = "C:\PEI_Investigation_Results"  # Output directory
    $ startDate = "2025-01-01"  # Search timeframe
    $ endDate = "2025-09-30" # End timeframe




**2. Run the script:**
   .\Forensic-PEI-Scanner.ps1



**3. Check results in $outputFolder**
   ├── 01_Files/          # Copied suspicious files
├── 02_Emails/         # Exported .msg files
├── 03_OneNote/        # OneNote extracts
├── 04_Browser/        # Browser history DBs
└── 05_Reports/        # CSVs with hits
