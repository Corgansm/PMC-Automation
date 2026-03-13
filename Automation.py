import os
import glob
import pdfplumber
import csv
import re
import win32com.client

# 1. SETUP: Get the exact folder where this script is saved
current_folder = os.path.dirname(os.path.abspath(__file__))
print(f"TROUBLESHOOT: Script running from: {current_folder}")

def run_valet_automation():
    pdf_path = None
    
    # ---------------------------------------------------------
    # PHASE 1: DOWNLOAD PDF FROM OUTLOOK
    # ---------------------------------------------------------
    print("\n--- STARTING PHASE 1: DOWNLOAD ---")
    try:
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        inbox = outlook.GetDefaultFolder(6) # 6 = Inbox
        messages = inbox.Items
        messages.Sort("[ReceivedTime]", True) # Newest first

        found_email = False
        for msg in messages:
            if "Step#1" in msg.Subject:
                print(f"TROUBLESHOOT: Found matching email! Subject: {msg.Subject}")
                for attachment in msg.Attachments:
                    if attachment.FileName.lower().endswith(".pdf"):
                        pdf_path = os.path.join(current_folder, attachment.FileName)
                        attachment.SaveAsFile(pdf_path)
                        print(f"SUCCESS: Downloaded '{attachment.FileName}'")
                        found_email = True
                        break 
                if found_email: break 
        
        if not found_email:
            print("TROUBLESHOOT: Could not find an email with 'Step#1' and a PDF.")
            
    except Exception as e:
        print(f"TROUBLESHOOT ERROR (Outlook): {e}")
        print("Tip: Make sure Classic Outlook remains open on your desktop.")

    # Fallback: If download failed, check if a PDF is already in the folder
    if not pdf_path or not os.path.exists(pdf_path):
        pdf_files = glob.glob(os.path.join(current_folder, "*.pdf"))
        if pdf_files:
            pdf_path = pdf_files[0]
            print(f"TROUBLESHOOT: Using existing PDF found in folder: {os.path.basename(pdf_path)}")
        else:
            print("TROUBLESHOOT ERROR: No PDF available to process. Exiting script.")
            return

    # ---------------------------------------------------------
    # PHASE 2: EXTRACT TEXT AND UPDATE CSV
    # ---------------------------------------------------------
    print("\n--- STARTING PHASE 2: EXTRACTION & UPDATE ---")
    try:
        # Extract text from the PDF
        with pdfplumber.open(pdf_path) as pdf:
            text = pdf.pages[0].extract_text()
            
        # Extract the Target Date Strictly
        date_match = re.search(r'Overnight Revenue Summary\s+([A-Z][a-z]{2})\s+(\d{1,2})', text, re.IGNORECASE)
        if date_match:
            month = date_match.group(1).capitalize()
            day = int(date_match.group(2))
            target_date = f"{day:02d}-{month}"
            print(f"TROUBLESHOOT: Regex found Report Date -> {date_match.group(0)}")
        else:
            target_date = "12-Mar" # Fallback
            print("TROUBLESHOOT: Regex FAILED to find the date. Using fallback 12-Mar.")
            
        print(f"TROUBLESHOOT: Formatted Target Date for CSV -> {target_date}")
            
        # Extract Vehicles and Revenue
        vehicles = re.findall(r'Vehicle Count\s+(\d+)', text)
        revenues = re.findall(r'Revenue\s+\$?([\d,]+\.\d{2})', text)
        
        if len(vehicles) >= 2 and len(revenues) >= 2:
            dv = int(vehicles[0]) 
            ov = int(vehicles[1]) 
            dr = float(revenues[0].replace(',', '')) 
            or_rev = float(revenues[1].replace(',', '')) 
            
            print(f"TROUBLESHOOT: Parsed Data -> DV: {dv}, DR: {dr}, OV: {ov}, OR: {or_rev}")

            # Open and Update the CSV File
            csv_name = "2026 RSS-102742 Trilogy Hotel Huntsville(RSS_Mar).csv"
            csv_path = os.path.join(current_folder, csv_name)
            
            if os.path.exists(csv_path):
                with open(csv_path, 'r', newline='', encoding='utf-8') as file:
                    reader = csv.reader(file)
                    rows = list(reader)
                
                target_col = None
                header_row = rows[1]
                
                print(f"TROUBLESHOOT: Searching CSV Row 2 for '{target_date}'...")
                
                for col_idx, cell_val in enumerate(header_row):
                    if target_date in str(cell_val):
                        target_col = col_idx
                        print(f"TROUBLESHOOT: MATCH FOUND! '{target_date}' is at CSV Index {target_col}")
                        break
                        
                if target_col is not None:
                    rows[72][target_col] = dv
                    rows[73][target_col] = dr
                    rows[78][target_col] = ov
                    rows[79][target_col] = or_rev
                    
                    with open(csv_path, 'w', newline='', encoding='utf-8') as file:
                        writer = csv.writer(file)
                        writer.writerows(rows)
                        
                    print(f"\nSUCCESS! CSV fully updated in the '{target_date}' column.")
                else:
                    print(f"TROUBLESHOOT ERROR: Could not find '{target_date}' in Row 2 of the CSV.")
            else:
                print(f"TROUBLESHOOT ERROR: CSV file '{csv_name}' not found in folder.")
        else:
            print("TROUBLESHOOT ERROR: Could not find all vehicle/revenue data in the PDF.")
            
    except Exception as e:
        print(f"TROUBLESHOOT ERROR (Extraction/CSV Update): {e}")

if __name__ == "__main__":
    run_valet_automation()