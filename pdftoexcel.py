
import win32com.client
import os
import datetime
import re
import camelot
import pandas as pd
import requests
import urllib3
import pdfplumber


urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# === CONFIGURATION ===
sender_email = "mustafa.sayyed@ariantechsolutions.com"
output_folder = r"C:\Users\ssing232\OneDrive\OneDrive - JLR\Desktop\Tableau\Reports"  # Change if needed

# === Ensure Output Folder Exists ===
if not os.path.exists(output_folder):
    os.makedirs(output_folder)

# === Get Outlook Namespace & Inbox ===
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)  # 6 = Inbox
messages = inbox.Items
messages.Sort("[ReceivedTime]", True)  # Sort latest first

# === Loop through Unread Messages from Mustafa ===
pdf_path = None
base_file_name_with_timestamp = None

for message in messages:
    if message.Class != 43:  # skip non-mail items
        continue
    if message.UnRead and message.SenderEmailAddress.lower() == sender_email.lower():
        attachments = message.Attachments
        if attachments.Count > 0:
            for i in range(1, attachments.Count + 1):
                attachment = attachments.Item(i)
                original_name = attachment.FileName

                # Clean filename from illegal characters
                original_name = re.sub(r'[\\/*?:"<>|]', "_", original_name)

                # Add timestamp to filename
                timestamp = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
                base_name = os.path.splitext(original_name)[0]
                extension = os.path.splitext(original_name)[1]
                new_filename = f"{base_name}_{timestamp}{extension}"
                full_path = os.path.join(output_folder, new_filename)

                attachment.SaveAsFile(full_path)
                print(f"✅ Saved: {full_path}")

                if full_path.lower().endswith(".pdf"):
                    pdf_path = full_path
                    base_file_name_with_timestamp = os.path.splitext(os.path.basename(full_path))[0]

            message.UnRead = False
        break  # Stop after processing the first unread email from Mustafa
else:
    print("📭 No unread emails from Mustafa with attachments.")

# === Process the PDF if found ===
if pdf_path:
    try:
        print("🔍 Extracting table using Camelot...")
        base_dir = os.path.dirname(pdf_path)
        output_csv = os.path.join(base_dir, f"{base_file_name_with_timestamp}.csv")
        output_json = os.path.join(base_dir, f"{base_file_name_with_timestamp}.json")

        tables = camelot.read_pdf(
            pdf_path,
            pages='all',
            flavor='stream',
            strip_text='\n',
            edge_tol=200
        )

        if tables and tables.n > 0:
            dfs = [table.df for table in tables]
            df = pd.concat(dfs, ignore_index=True)

            df = df.dropna(how='all')
            df = df.loc[:, ~df.columns.duplicated()]
            df = df.iloc[:, 2:]

            new_columns = [
                "user_email", "enquiry", "unique_td",
                "enquiry_to_UTD", "enquiry_to_TD", "new_orders",
                "td_to_retail", "cancellations", "cancellation_contro"
            ]
            df.columns = new_columns

            df["data_score"] = df.iloc[:, 1:].apply(
                lambda r: r.astype(str).str.strip().str.len().sum(), axis=1
            )

            header_keywords = [
                "user_email", "ps", "enquiry", "unique", "conversion",
                "orders", "ratio", "cancel", "cc", "sa - mtd", "retailer"
            ]
            df = df[~df["user_email"].astype(str).str.lower().apply(
                lambda x: any(keyword in x for keyword in header_keywords)
            )]

            df = df[~((df["user_email"].isna()) | (df["user_email"].astype(str).str.strip() == ""))]

            df = df[df["user_email"].str.contains(
                r"^[a-zA-Z0-9._%+-]+@[A-Za-z0-9.-]+\.[a-zA-Z]{2,}$", na=False
            )]

            df = df.sort_values("data_score", ascending=False).drop_duplicates("user_email", keep="first")
            df = df.drop(columns=["data_score"]).reset_index(drop=True)

            # Save CSV
            df.to_csv(output_csv, index=False)
            print(f"✅ Clean CSV saved at:\n{output_csv}")

            # Save JSON
            df.to_json(output_json, orient="records", indent=2)
            print(f"📤 Clean JSON saved at:\n{output_json}")

            # Upload to API
            with open(output_json, 'r', encoding='utf-8') as f:
                json_data = f.read()

            headers = {
                "Content-Type": "application/json"
            }

            response = requests.post(
                "https://api.smartassistapp.in/api/bulk-insert/analytics",
                data=json_data,
                headers=headers,
                verify= False
            )

            if response.status_code in [200, 201]:
                print("✅ Data upload successful!")
            else:
                print(f"❌ Upload failed. Status: {response.status_code}")
                print(response.text)

        else:
            print("❌ No tables found in the PDF. Try adjusting Camelot settings.")

    except Exception as e:
        print("❌ Exception during PDF to CSV/JSON processing:", str(e))