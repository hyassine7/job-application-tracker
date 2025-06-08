import win32com.client

ol = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

print("=== Top‐Level Stores and Their Subfolders ===\n")
for store in ol.Folders:
    print(f"Store: {store.Name}")
    # List only the immediate child folders of each store
    for sub in store.Folders:
        print(f"    └─ Subfolder: {sub.Name}")
    print()
