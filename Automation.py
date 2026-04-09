import os
import glob
import pdfplumber
import re
import win32com.client
import openpyxl
import datetime

# 1. SETUP: Get the exact folder where this script is saved
current_folder = os.path.dirname(os.path.abspath(__file__))
print(f"TROUBLESHOOT: Script running from: {current_folder}")

def run_valet_automation():
    # ---------------------------------------------------------
    # PHASE 1: DOWNLOAD ALL PDFS FROM THE PAST WEEK
    # ---------------------------------------------------------
    print("\n--- STARTING PHASE 1: DOWNLOAD (PAST 7 DAYS) ---")
    downloaded_files = []
    try:
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        inbox = outlook.GetDefaultFolder(6) 
        root_folder = inbox.Parent
        
        target_folder = root_folder.Folders.Item("Revenue Tracking")
        print(f"TROUBLESHOOT: Accessed folder: {target_folder.Name}")

        last_week = datetime.datetime.now() - datetime.timedelta(days=7)
        date_filter = last_week.strftime("%m/%d/%Y %I:%M %p")
        
        sender_email = "noreply@notifications.parkingmgt.com"
        sFilter = f"[SenderEmailAddress] = '{sender_email}' AND [ReceivedTime] >= '{date_filter}'"
        
        messages = target_folder.Items.Restrict(sFilter)
        messages.Sort("[ReceivedTime]", True)

        print(f"TROUBLESHOOT: Found {messages.Count} potential emails from the last 7 days.")

        for msg in messages:
            subject = msg.Subject
            if "Step#1" in subject and "Trilogy Huntsville" in subject:
                received_date = msg.ReceivedTime.strftime("%Y-%m-%d")
                for attachment in msg.Attachments:
                    if attachment.FileName.lower().endswith(".pdf"):
                        unique_filename = f"{received_date}_{attachment.FileName}"
                        pdf_path = os.path.join(current_folder, unique_filename)
                        attachment.SaveAsFile(pdf_path)
                        downloaded_files.append(pdf_path)
                        print(f"SUCCESS: Downloaded '{unique_filename}'")
                        break 
        
    except Exception as e:
        print(f"TROUBLESHOOT ERROR (Outlook): {e}")

    if not downloaded_files:
        print("TROUBLESHOOT: No new PDFs found. Exiting.")
        return

    # ---------------------------------------------------------
    # PHASE 2: EXTRACTION & UPDATE
    # ---------------------------------------------------------
    print(f"\n--- STARTING PHASE 2: PROCESSING {len(downloaded_files)} FILES ---")
    excel_path = r"C:\Users\AltonJones\OneDrive - Parking Management Company LLC\Operations - Accounting - Huntsville AL Hotels\2026\2026 RSS-102742 Trilogy Hotel Huntsville.xlsx"
    
    try:
        print(f"TROUBLESHOOT: Opening workbook at {excel_path}")
        wb_read = openpyxl.load_workbook(excel_path, data_only=True)
        wb_save = openpyxl.load_workbook(excel_path)

        for pdf_file in downloaded_files:
            print(f"\n--- Processing: {os.path.basename(pdf_file)} ---")
            with pdfplumber.open(pdf_file) as pdf:
                text = pdf.pages[0].extract_text()
            
            date_match = re.search(r'Overnight Revenue Summary\s+([A-Z][a-z]{2})\s+(\d{1,2})', text, re.IGNORECASE)
            if not date_match:
                print(f"TROUBLESHOOT: Date not found in PDF text.")
                continue

            month_str = date_match.group(1).capitalize()
            day_num = int(date_match.group(2))
            month_num = datetime.datetime.strptime(month_str, "%b").month
            print(f"TROUBLESHOOT: PDF indicates date is {month_str} {day_num} (Month #{month_num})")
            
            vehicles = re.findall(r'Vehicle Count\s+(\d+)', text)
            revenues = re.findall(r'Revenue\s+\$?([\d,]+\.\d{2})', text)
            
            if len(vehicles) >= 2 and len(revenues) >= 2:
                dv, ov = int(vehicles[0]), int(vehicles[1])
                dr = float(revenues[0].replace(',', ''))
                or_rev = float(revenues[1].replace(',', ''))

                sheet_map = {"Jan": "RSS_Jan", "Feb": "RSS_Feb", "Mar": "RSS_Mar", "Apr": "RSS_Apr", "May": "RSS_May", "Jun": "RSS_June", "Jul": "RSS_July", "Aug": "RSS_Aug", "Sep": "RSS_Sept", "Oct": "RSS_Oct", "Nov": "RSS_Nov", "Dec": "RSS_Dec"}
                sheet_name = sheet_map.get(month_str, f"RSS_{month_str}")
                
                if sheet_name in wb_read.sheetnames:
                    ws_save = wb_save[sheet_name]
                    
                    print(f"TROUBLESHOOT: Calculating target column for Day {day_num}...")
                    target_col = day_num + 4
                    
                    ws_save.cell(row=73, column=target_col).value = dv
                    ws_save.cell(row=74, column=target_col).value = dr
                    ws_save.cell(row=79, column=target_col).value = ov
                    ws_save.cell(row=80, column=target_col).value = or_rev
                    print(f"SUCCESS: Data for {month_str} {day_num} updated in Column {target_col}")
            else:
                print(f"TROUBLESHOOT ERROR: Extraction failed for {pdf_file}")

        # --- UPDATED SAVE LOGIC ---
        wb_save.save(excel_path)
        wb_save.close() # Closes the file handle to trigger OneDrive sync
        wb_read.close()
        print(f"\nSUCCESS: Workbook saved and closed. OneDrive will now sync the updates.")

        print("\n--- FINAL CLEANUP ---")
        for pdf_file in downloaded_files:
            try:
                os.remove(pdf_file)
                print(f"Deleted: {os.path.basename(pdf_file)}")
            except: pass

    except Exception as e:
        print(f"CRITICAL ERROR: {e}")

if __name__ == "__main__":
    run_valet_automation()