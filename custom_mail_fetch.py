import os
import win32com.client
from datetime import datetime
import tkinter as tk
from tkinter import messagebox

# Function to fetch emails based on user input
def fetch_emails(email_type):
    """Fetch emails from Outlook based on the user-defined type."""
    items = []

    try:
        # Connect to Outlook
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        inbox = outlook.GetDefaultFolder(6)  # "6" refers to the inbox folder
        messages = inbox.Items
        messages.Sort("[ReceivedTime]", True)  # Sort by received time, newest first

        # Filter emails based on the user's input
        for message in messages:
            if email_type.lower() in message.SenderEmailAddress.lower():
                msg = {
                    "Subject": message.Subject,
                    "SentOn": message.SentOn,
                    "EntryID": message.EntryID,
                    "Sender": message.SenderName,
                    "Size": message.Size,
                    "Body": message.Body,
                }
                items.append(msg)

    except Exception as ex:
        print(f"Error accessing Outlook: {ex}")

    return items

# Function to save fetched emails to text files
def save_emails_to_files(emails, email_type):
    """Save fetched emails to files."""
    folder_name = f"{email_type}_Outlook_Emails"
    if not os.path.exists(folder_name):
        os.makedirs(folder_name)

    for i, email in enumerate(emails):
        # Format the sent time for the filename
        sent_time_str = email["SentOn"].strftime("%Y-%m-%d_%H-%M-%S") if email["SentOn"] else "UNKNOWN_DATE"
        filename = f"{sent_time_str}_{i + 1}.txt"
        filepath = os.path.join(folder_name, filename)
        
        # Write the email details into the file
        with open(filepath, 'w', encoding='utf-8') as f:
            f.write(f"Sender: {email['Sender']}\n")
            f.write(f"Subject: {email['Subject']}\n")
            f.write(f"Size: {email['Size']} bytes\n")
            f.write(f"Body:\n{email['Body']}\n")
        
        print(f"Email saved to: {filepath}")

# Main function that links to the Tkinter interface
def main():
    """Main function to fetch and save emails based on user input."""
    
    # Get user input from the Tkinter entry box
    email_type = email_input.get()
    
    if email_type:
        emails = fetch_emails(email_type)
        if emails:
            save_emails_to_files(emails, email_type)
            messagebox.showinfo("Success", f"Emails related to '{email_type}' have been saved.")
        else:
            messagebox.showinfo("No Emails", f"No emails found for '{email_type}'.")
    else:
        messagebox.showwarning("Input Error", "Please enter a valid email type.")

# Tkinter setup
root = tk.Tk()
root.title("Email Fetcher")
root.geometry("400x200")

# Label for the input field
label = tk.Label(root, text="Enter email type (e.g., LinkedIn, Twitter, etc.):")
label.pack(pady=10)

# Input field
email_input = tk.Entry(root, width=50)
email_input.pack(pady=10)

# Button to start the email fetching process
fetch_button = tk.Button(root, text="Fetch Emails", command=main)
fetch_button.pack(pady=10)

# Run the Tkinter event loop
root.mainloop()
