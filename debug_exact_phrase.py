import os
import win32com.client
import pandas as pd
from datetime import datetime, timedelta   # ← added

def get_outlook_folder(account_name: str, folder_name: str):
    """
    Connect to Outlook and return the specified folder under the given account.
    """
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    for store in outlook.Folders:
        if store.Name.lower() == account_name.lower():
            for subfolder in store.Folders:
                if subfolder.Name.lower() == folder_name.lower():
                    return subfolder
    raise ValueError(f"Could not find '{folder_name}' under '{account_name}'.")

def extract_matching_phrases(folder):
    """
    Only include emails from the last 90 days whose Subject or Body
    contains any of the exact substrings in `target_phrases` (case-insensitive).
    """
    messages = folder.Items
    messages.Sort("[ReceivedTime]", True)  # newest → oldest

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
        "we have recieved your application",
        "application recieved",
        "application was recieved",
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

            data.append({
                "Date Received": recv.strftime("%Y-%m-%d %H:%M:%S"),
                "Sender Name": item.SenderName or "",
                "Sender Email": getattr(item, "Sender", None) and item.Sender.Address or "",
                "Subject": item.Subject or "",
                "Has Attachments": bool(item.Attachments.Count) if hasattr(item, "Attachments") else False,
                "Status": "",
                "Last Updated": now_str
            })

    return data

if __name__ == "__main__":
    # Example usage for debugging:
    ACCOUNT = "hassanyassine.work@gmail.com"
    FOLDER  = "Inbox"
    try:
        inbox = get_outlook_folder(ACCOUNT, FOLDER)
    except ValueError as e:
        print(f"ERROR: {e}")
    else:
        results = extract_matching_phrases(inbox)
        print(f"Found {len(results)} matching emails:")
        for r in results:
            print(r)
