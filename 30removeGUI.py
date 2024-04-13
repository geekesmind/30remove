import win32com.client
import datetime
import tkinter as tk
from tkinter import messagebox

def delete_old_emails():
    # Connect to Outlook
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

    # Select the inbox folder (you can change this to any other folder as needed)
    inbox = outlook.GetDefaultFolder(6)  # 6 represents the index of the inbox folder

    # Get all emails in the inbox
    emails = inbox.Items

    # Get today's date
    today = datetime.datetime.today()

    # Count of deleted emails
    deleted_count = 0

    # Iterate through each email
    for email in list(emails):
        # Get the ReceivedTime of the email as a datetime object
        received_time = email.ReceivedTime

        # Convert the ReceivedTime to a Python datetime object
        received_time = datetime.datetime(
            received_time.year,
            received_time.month,
            received_time.day
        )

        # Check if the email received time is older than 60 days
        if (today - received_time).days >= 60:
            email.Delete()  # Delete the email
            deleted_count += 1

    messagebox.showinfo("Deletion Complete", f"{deleted_count} emails deleted.")

def delete_old_emails_gui():
    # Create the main window
    window = tk.Tk()
    window.title("Outlook Email Deletion")

    # Create a label
    label = tk.Label(window, text="Click the button to delete outlook emails older than 30 days.")
    label.pack(pady=10)

    # Create a button to trigger the email deletion
    button = tk.Button(window, text="Delete Emails", command=delete_old_emails)
    button.pack(pady=10)

    # Run the GUI loop
    window.mainloop()

if __name__ == "__main__":
    delete_old_emails_gui()