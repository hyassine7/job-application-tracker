import os
import sys
# ensure this script’s folder is first on the import path
sys.path.insert(0, os.path.dirname(__file__))

import win32com.client
import pandas as pd
from datetime import datetime, timedelta

from formatting import save_with_formatting
from ui_helpers import ask_and_open


def get_outlook_folder(account_name: str, folder_name: str):
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    for store in outlook.Folders:
        if store.Name.lower() == account_name.lower():
            for subfolder in store.Folders:
                if subfolder.Name.lower() == folder_name.lower():
                    return subfolder
    raise ValueError(f"Could not find '{folder_name}' under '{account_name}'.")


def extract_matching_phrases(folder):
    messages = folder.Items
    messages.Sort("[ReceivedTime]", True)

    data = []
    now = datetime.now()
    cutoff = now - timedelta(days=90)
    now_str = now.strftime("%Y-%m-%d %H:%M:%S")

    target_phrases = [
        "application submitted",
        "thank you for applying",
        "application sent",
        "thanks for applying",
        "thanks for applying to google!",
        "we have received your application",
        "application received",
        "application was received",
        "your application has been sent"
    ]

    for item in messages:
        if item.Class == 43 and hasattr(item, "ReceivedTime") and item.ReceivedTime:
            recv = item.ReceivedTime
            if recv.tzinfo is not None:
                recv = recv.replace(tzinfo=None)
            if recv < cutoff:
                break

            subj = (item.Subject or "").lower()
            body = (item.Body or "").lower()
            if not any(phrase in subj or phrase in body for phrase in target_phrases):
                continue

            sender_email = ""
            if hasattr(item, "Sender") and item.Sender is not None:
                sender_email = getattr(item.Sender, "Address", "")

            data.append({
                "Date Received": recv.strftime("%Y-%m-%d %H:%M:%S"),
                "Sender Name": item.SenderName or "",
                "Sender Email": sender_email,
                "Subject": item.Subject or "",
                "Has Attachments": bool(item.Attachments.Count) if hasattr(item, "Attachments") else False,
                "Status": "",
                "Last Updated": now_str
            })

    return data


def main():
    OUTLOOK_STORE_NAME = "hassanyassine.work@gmail.com"
    FOLDER_NAME        = "Inbox"
    OUTPUT_FILE        = r"C:\Scripts\job_applications_tracker.xlsx"

    try:
        inbox = get_outlook_folder(OUTLOOK_STORE_NAME, FOLDER_NAME)
    except ValueError as e:
        print(f"ERROR: {e}")
        return

    records = extract_matching_phrases(inbox)
    if not records:
        print("ERROR: No matching application emails in the last 90 days.")
        return

    df = pd.DataFrame(records)
    cols = ["Date Received", "Sender Name", "Sender Email", "Subject", "Has Attachments", "Status", "Last Updated"]
    df = df[[c for c in cols if c in df.columns]]

    if os.path.isfile(OUTPUT_FILE):
        old = pd.read_excel(OUTPUT_FILE, engine="openpyxl")
        combined = pd.concat([old, df], ignore_index=True)
        combined.drop_duplicates(subset=["Date Received", "Subject"], inplace=True)
        combined["Date Received"] = pd.to_datetime(combined["Date Received"])
        combined.sort_values("Date Received", ascending=False, inplace=True)
        combined["Date Received"] = combined["Date Received"].dt.strftime("%Y-%m-%d %H:%M:%S")
        final_df = combined[cols]
    else:
        final_df = df.copy()

    save_with_formatting(final_df, OUTPUT_FILE)
    print(f"Tracker updated: now contains {len(final_df)} rows.")
    ask_and_open(OUTPUT_FILE)


if __name__ == "__main__":
    main()
