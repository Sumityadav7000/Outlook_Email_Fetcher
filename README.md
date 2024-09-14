The Outlook Email Fetcher is a Python project that automates fetching specific types of emails (e.g., LinkedIn, Twitter, Instagram, YouTube, Naukri, etc.) from Microsoft Outlook. The user can input the type of email to be fetched using a simple Tkinter-based graphical interface. The emails are then saved to text files for easy viewing and analysis.
Fetch emails from Microsoft Outlook based on keywords such as LinkedIn, Twitter, Instagram, YouTube, Naukri, and more.
Easy-to-use graphical interface built with Tkinter.
Save fetched emails to text files with detailed information (subject, sender, body, etc.).
Supports multiple types of emails, customizable by the user.

How to Use
Launch the application by running email_fetcher.py.
A graphical interface will appear, allowing you to input the type of emails you'd like to fetch (e.g., LinkedIn, Twitter, etc.).
Click the "Fetch Emails" button, and the program will search for emails matching the input and save them as .txt files in a folder named Fetched_Emails.
The saved emails will include details such as:
Sender name and email address.
Subject of the email.
The body of the email.
Size of the email.


Outlook_Email_Fetcher/
│
├── email_fetcher.py         # Main Python script that fetches the emails
├── README.md                # Project documentation
├── requirements.txt         # List of required Python packages
├── Fetched_Emails/          # Folder where fetched emails are stored
│
