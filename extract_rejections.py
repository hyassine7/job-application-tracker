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


def extract_rejection_emails(folder):
    messages = folder.Items
    messages.Sort("[ReceivedTime]", True)

    data = []
    now = datetime.now()
    cutoff = now - timedelta(days=90)
    now_str = now.strftime("%Y-%m-%d %H:%M:%S")

    decline_phrases = [
        "unfortunately",
        "we regret to inform you",
        "not selected",
        "will not be moving forward",
        "no longer in consideration",
        "we won’t be moving forward",
        "we’re unable to proceed with your application"
        "we have decided not to move forward",
        "we have decided not to proceed",
        "we have decided not to continue",
        "we have decided not to advance",
        "we have decided not to take your application further",
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
            if not any(phrase in subj or phrase in body for phrase in decline_phrases):
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
                "Status": "No longer in consideration",
                "Last Updated": now_str
            })

    return data


def main():
    OUTLOOK_STORE_NAME = "hassanyassine.work@gmail.com"
    FOLDER_NAME        = "Inbox"
    OUTPUT_FILE        = r"C:\Scripts\job_applications_rejections.xlsx"

    try:
        inbox = get_outlook_folder(OUTLOOK_STORE_NAME, FOLDER_NAME)
    except ValueError as e:
        print(f"ERROR: {e}")
        return

    rejects = extract_rejection_emails(inbox)
    if not rejects:
        print("No rejection emails found in the last 90 days.")
        return

    df = pd.DataFrame(rejects)
    cols = ["Date Received", "Sender Name", "Sender Email", "Subject", "Has Attachments", "Status", "Last Updated"]
    df = df[[c for c in cols if c in df.columns]]

    save_with_formatting(df, OUTPUT_FILE)
    print(f"Rejection register created with {len(df)} rows.")
    ask_and_open(OUTPUT_FILE)


if __name__ == "__main__":
    main()
