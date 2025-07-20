# Excel-to-WhatsApp-Daily-Report-Automation

This project automates the daily process of generating and sending personalized performance summaries from an Excel report to individual employees via WhatsApp Web.

## 📌 Overview

The automation extracts daily data for each field executive, formats it into a customized message, and delivers it directly to each employee’s WhatsApp chat.
This eliminates manual daily reporting and ensures timely, consistent updates.

## ⚙️ Tech Stack

- **Python**
- **Selenium**
- **Pandas**
- **Babel** (for currency formatting)
- **Excel (ERP Data Source)**

## 📂 Main Script

- `whatsapp_dcr_group.py`  
  - Reads data from a shared Excel file.
  - Formats personalized messages including:
    - Systematic & approved penalties
    - Daily primary & secondary sales
    - Month-to-date sales
    - Total calls made
    - Login, first & last call times
    - Daily tour location
    - Salary deductions & absent count
  - Automates WhatsApp Web login and sends messages to 120+ contacts daily.
  - Includes retry logic for failed contacts.

## 🚀 Usage

1. Update `EXCEL_PATH` and `SHEET_NAME` as per your environment.
2. Adjust ChromeDriver path and user profile settings.
3. Run the script to start sending daily updates.
4. This version is sanitized — no actual contact names or credentials are included.

## ✅ Impact

This automation saves hours of manual messaging work for daily performance tracking and helps managers maintain accurate, timely communication with large field teams.

## 🔒 Note

This is a sample version for demonstration. Please ensure you have proper permission to automate messaging on WhatsApp Web.
