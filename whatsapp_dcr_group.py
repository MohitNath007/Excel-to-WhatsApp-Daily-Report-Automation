import os
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from babel.numbers import format_currency

# -----------------------------------------
# CONFIGURATION
# -----------------------------------------
EXCEL_PATH = r"\\192.168.10.10\Support\Daily Rep\23-24,24-25 BI report\DCR - PBI.xlsx.xlsm"
SHEET_NAME = "DCR - Automation"

df = pd.read_excel(EXCEL_PATH, sheet_name=SHEET_NAME)

# -----------------------------------------
# FUNCTION: Format message for each row
# -----------------------------------------
def format_message(row):
    def safe_get(value, decimal=False, currency=False):
        if pd.isna(value):
            return "N/A"
        if currency:
            return format_currency(round(value),'INR',locale='en_IN',format=u'¤#,##0',currency_digits=False)
        return f"{value:.2f}" if decimal else value

    date_formatted = row['Date'].strftime('%d-%m-%Y') if not pd.isna(row['Date']) else "N/A"
    Month_formatted = row['Date'].strftime('%B') if not pd.isna(row['Date']) else "N/A"
    total_penalty = row['Total (P)'] * 100 if not pd.isna(row['Total (P)']) else 0
    approved_penalty = row['Approved (P)'] * 100 if not pd.isna(row['Approved (P)']) else 0

    return f"""Dear {safe_get(row['Full Name'])},

Date: {date_formatted}

Your working for the day is as follows

Tour plan       : {safe_get(row['Tour Plan Status'])}
Location type   : {safe_get(row['Location Type'])}
First call      : {safe_get(row['First Call'])}
Last call       : {safe_get(row['Last Call'])}
Total calls     : {safe_get(row['Total Calls'])}

Penalty
Tour plan       : {safe_get(row['Tour Plan (P)'])}
First call      : {safe_get(row['First call (P)'])}
Last call       : {safe_get(row['Last Call (P)'])}
Total calls     : {safe_get(row['Total Call (P)'])}

Systematic penalty   : {total_penalty:.0f}%
Approved penalty     : {approved_penalty:.0f}%
Per Day Salary       : {safe_get(row['Today Salary'])}
{date_formatted} Deduction    : {safe_get(row['Approved Deduction'], decimal=True)}
Current Month Deduction : {safe_get(row['Monthly Deduction'], decimal=True)}
Total Absent (This Month) : {safe_get(row['Absent Count'])}

Sales for {date_formatted}
Primary Sales   : {safe_get(row['Primary Sales'], currency=True)}
Secondary Sales : {safe_get(row['Secondary Sales'], currency=True)}

Sales for {Month_formatted}
Primary Sales       : {safe_get(row['AON PS'], currency=True)}
Secondary Sales     : {safe_get(row['AON SS'], currency=True)}
"""

# -----------------------------------------
# FUNCTION: Send WhatsApp messages
# -----------------------------------------
def send_whatsapp_messages(retry_mode=False, failed_contacts_previous_run=[]):
    if not retry_mode:
        # Normal mode - filter for 'Mark' = "Yes"
        filtered_df = df[df['Contact Name'].notna() & (df['Contact Name'] != "") & (df['Mark'] == "Yes")]
        print(f"📩 Preparing to send {len(filtered_df)} messages (Marked 'Yes')...")
    else:
        # Retry mode - only target previously failed contacts
        filtered_df = df[df['Contact Name'].isin(failed_contacts_previous_run)]
        print(f"🔁 Retrying {len(filtered_df)} previously failed contacts...")

    if filtered_df.empty:
        print("❌ No valid contacts found for this mode.")
        return []

    successful_contacts = []
    failed_contacts = []

    chrome_options = Options()
    chrome_options.add_argument("--user-data-dir=D:/automation/Chrome")
    chrome_options.add_argument("--profile-directory=Profile 9")
    chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    
    service = Service("D:/automation/chromedriver.exe")
    driver = webdriver.Chrome(service=service, options=chrome_options)

    try:
        driver.get("https://web.whatsapp.com/")
        print("⏳ Waiting for WhatsApp Web to load...")
        WebDriverWait(driver, 60).until(
            EC.presence_of_element_located((By.XPATH, "//div[@contenteditable='true']"))
        )
        print("✓ WhatsApp Web ready")

        for index, row in filtered_df.iterrows():
            contact_name = str(row['Contact Name']).strip()
            message = format_message(row)

            print(f"\n💬 Attempting to send to {contact_name}...")

            try:
                # Clear search box thoroughly
                search_box = WebDriverWait(driver, 20).until(
                    EC.presence_of_element_located((By.XPATH, "//div[@contenteditable='true']"))
                )
                for _ in range(3):
                    search_box.click()
                    search_box.send_keys(Keys.CONTROL + "a")
                    search_box.send_keys(Keys.DELETE)
                    time.sleep(0.5)
                
                search_box.send_keys(contact_name)
                time.sleep(2)

                # Click contact
                contact_xpath = f"//span[@title='{contact_name}']"
                WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, contact_xpath))
                )
                driver.find_element(By.XPATH, contact_xpath).click()
                time.sleep(2)

                # Send message
                input_box = WebDriverWait(driver, 15).until(
                    EC.element_to_be_clickable((By.XPATH, "//div[@contenteditable='true'][@data-tab='10']"))
                )
                input_box.click()
                input_box.send_keys(Keys.CONTROL + "a")
                input_box.send_keys(Keys.DELETE)
                time.sleep(0.5)
                
                lines = message.split('\n')
                for i, line in enumerate(lines):
                    input_box.send_keys(line)
                    if i < len(lines) - 1:
                        input_box.send_keys(Keys.SHIFT + Keys.ENTER)
                    time.sleep(0.2)
                
                input_box.send_keys(Keys.ENTER)
                
                WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, "//span[@data-icon='msg-time'] | //span[@data-icon='msg-dblcheck']"))
                )
                
                print("✅ Sent successfully!")
                successful_contacts.append(contact_name)
                time.sleep(2)

            except Exception as e:
                print(f"❌ Failed to send to {contact_name}: {str(e)}")
                failed_contacts.append(contact_name)
                
                try:
                    driver.find_element(By.XPATH, "//div[@contenteditable='true']").send_keys(Keys.CONTROL + "a")
                    driver.find_element(By.XPATH, "//div[@contenteditable='true']").send_keys(Keys.DELETE)
                except:
                    pass
                continue

    except Exception as main_ex:
        print(f"🚨 Critical error: {str(main_ex)}")
        failed_contacts.extend([row['Contact Name'] for _, row in filtered_df.iterrows()])
    finally:
        driver.quit()
        
        # Print report
        print("\n📊 REPORT:")
        print("----------------")
        print(f"✅ Successfully sent ({len(successful_contacts)}):")
        for name in successful_contacts:
            print(f"  - {name}")
        
        print(f"\n❌ Failed ({len(failed_contacts)}):")
        for name in failed_contacts:
            print(f"  - {name}")
        
        print("\n🎉 Batch completed!")
        return failed_contacts

# -----------------------------------------
# MAIN EXECUTION
# -----------------------------------------
if __name__ == "__main__":
    first_run_failures = send_whatsapp_messages(retry_mode=False)
    
    if first_run_failures:
        while True:
            retry = input("\nRetry failed contacts? (yes/no): ").strip().lower()
            if retry in ('yes', 'y'):
                first_run_failures = send_whatsapp_messages(
                    retry_mode=True,
                    failed_contacts_previous_run=first_run_failures
                )
                if not first_run_failures:
                    break
            elif retry in ('no', 'n'):
                print("Skipping retry.")
                break
            else:
                print("Please enter 'yes' or 'no'")