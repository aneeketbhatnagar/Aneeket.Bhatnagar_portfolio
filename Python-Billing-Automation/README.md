# Python Billing Automation Suite

Full end-to-end billing workflow automation tool developed in Python for Clarivate's Billing & Revenue Analytics team.

## Overview
This suite automates the entire monthly billing cycle:
- Master file processing from SharePoint/OneDrive
- EWS (Early Warning System) report generation with client-wise grouping
- Old vs New comparison with change tracking
- Filtered data extraction with formulas
- Automatic email distribution via Outlook 365 (work account)
- GUI for easy execution (no coding required for end-users)

## Key Features
- **EWS Report Generation**: Client-wise grouping, full Excel formatting, auto-fit columns, color-coded headers
- **Old vs New Comparison**: Detects changes in monthly data, highlights modified line items
- **Data Filtering**: Monthly filtered view with calculated columns and formulas
- **Email Automation**: Sends personalized reports to recipients using filename-based email extraction
- **GUI Interface**: Tkinter-based user-friendly interface with progress tracking
- **Cloud Integration**: Designed for SharePoint/OneDrive paths (local fallback supported)

## Tech Stack
- Python 3.x
- openpyxl (Excel processing & formatting)
- win32com (Outlook 365 integration)
- tkinter (GUI)
- os, shutil, datetime (file operations)

## Business Impact
- Reduced manual reporting time by 80%
- Eliminated human errors in monthly billing data processing
- Enabled faster decision-making with automated change detection
- Streamlined communication with automatic email distribution

## Project Structure
- `Billing_automation.py` — Main GUI launcher
- `EWS_automation.py` — EWS report generation
- `OldVsNew.py` — Change comparison logic
- `Filter_automation.py` — Monthly filtered data
- `Email_Sender_GUI_Outlook365.py` — Automated email distribution

## Note
Code is optimized for internal use with Clarivate's file structure and permissions. Confidential data has been removed/anonymized for portfolio purposes.

---
*12+ years of experience in building production-grade automation solutions*
