# export_job_applications.py

import os
import sys
import argparse
# ensure this script’s folder is on the import path
sys.path.append(os.path.dirname(__file__))

import win32com.client
import pandas as pd
from datetime import datetime, timedelta

from formatting import save_with_formatting
from ui_helpers import ask_and_open


def get_outlook_folder(account_name: str, folder_name: str):
    """
    Connect to Outlook and return the specified subfolder.
    - account_name: e.g. "Gmail – youremail@example.com" or "[Gmail]"
    - folder_name: e.g. "Inbox" or "Job Applications"
    """
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    for store in outlook.Folders:
        if store.Name.lower() == account_name.lower():
            for sub in store.Folders:
                if sub.Name.lower() == folder_name.lower():
                    return sub
    raise ValueError(f"Could not find folder '{folder_name}' under store '{account_name}'.")


def extract_matching_phrases(folder, days_back=int):
    """
    Pulls emails from the last N days whose Subject or Body
    contains any exact substring in `TARGET_PHRASES` (case-insensitive).
    """
    MINDAYS = 90
    TARGET_PHRASES = [
        "application submitted",
        "thank you for applying",
        "application sent",
        "thanks for applying",
        # add or remove as desired...
    ]

    items = folder.Items
    items.Sort("[ReceivedTime]", True)

    cutoff = datetime.now() - timedelta(days=MINDAYS)
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    records = []
    for msg in items:
        if msg.Class != 43 or not hasattr(msg, "ReceivedTime"):
            continue

        rec_time = msg.ReceivedTime
        if rec_time.tzinfo:
            rec_time = rec_time.replace(tzinfo=None)
        if rec_time < cutoff:
            break

        subj = (msg.Subject or "").lower()
        body = (msg.Body or "").lower()
        if not any(phrase in subj or phrase in body for phrase in TARGET_PHRASES):
            continue

        sender = getattr(msg, "Sender", None)
        email = getattr(sender, "Address", "") if sender else ""
        records.append({
            "Date Received": rec_time.strftime("%Y-%m-%d %H:%M:%S"),
            "Sender Name": msg.SenderName or "",
            "Sender Email": email,
            "Subject": msg.Subject or "",
            "Has Attachments": bool(getattr(msg, "Attachments", [])) and msg.Attachments.Count > 0,
            "Status": "",
            "Last Updated": now_str
        })

    return records
def parse_args():
    p = argparse.ArgumentParser(
        description="Track job-application emails in Outlook and export to Excel."
    )
    p.add_argument(
        "--store", "-s",
        required=True,
        help="Outlook store/account name (e.g. '[Gmail]' or 'youremail@gmail.com')"
    )
    p.add_argument(
        "--folder", "-f",
        default="Inbox",
        help="Name of the folder under that store to scan (default: Inbox)"
    )
    p.add_argument(
        "--output", "-o",
        default="job_applications_tracker.xlsx",
        help="Path to the Excel output file"
    )
    p.add_argument(
        "--days", "-d",
        type=int,
        default=90,
        help="How many days back to scan (default: 90)"
    )
    return p.parse_args()


def main():
    args = parse_args()

    OUTLOOK_STORE = args.store
    FOLDER_NAME   = args.folder
    OUTPUT_FILE   = args.output
    DAYS_BACK     = args.days

    try:
        inbox = get_outlook_folder(OUTLOOK_STORE, FOLDER_NAME)
    except ValueError as e:
        print(f"ERROR: {e}")
        return

    # pass DAYS_BACK into extract…
    new_records = extract_matching_phrases(inbox, DAYS_BACK)

    try:
        folder = get_outlook_folder(OUTLOOK_STORE, FOLDER_NAME)
    except ValueError as e:
        print(f"ERROR: {e}")
        return

    records = extract_matching_phrases(folder)
    if not records:
        print("ERROR: no matching emails found in the last period.")
        return

    df = pd.DataFrame(records)
    cols = ["Date Received", "Sender Name", "Sender Email", "Subject", "Has Attachments", "Status", "Last Updated"]
    df = df[[c for c in cols if c in df]]

    # if file exists, read, concatenate & dedupe
    if os.path.isfile(OUTPUT_FILE):
        old = pd.read_excel(OUTPUT_FILE, engine="openpyxl")
        combined = pd.concat([old, df], ignore_index=True)
        combined.drop_duplicates(subset=["Date Received", "Subject"], inplace=True)
        combined["Date Received"] = pd.to_datetime(combined["Date Received"])
        combined.sort_values("Date Received", ascending=False, inplace=True)
        combined["Date Received"] = combined["Date Received"].dt.strftime("%Y-%m-%d %H:%M:%S")
        final = combined[cols]
    else:
        final = df.copy()

    save_with_formatting(final, OUTPUT_FILE)
    print(f"✔ Tracker updated: {len(final)} rows total.")
    ask_and_open(OUTPUT_FILE)


if __name__ == "__main__":
    main()
