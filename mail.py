import time
import win32com.client
import pygetwindow
from datetime import datetime



def sendmail(filename, breach_count):

    # Create an instance of the Outlook application
    outlook = win32com.client.Dispatch("Outlook.Application")

    # Create a new email
    mail = outlook.CreateItem(0)
    n = int(datetime.now().strftime("%d"))
    suffix = {1: 'st', 2: 'nd', 3: 'rd'}.get(n if n < 20 else n % 10, 'th')
    current_date = str(n) + suffix + datetime.now().strftime(" %B %Y")

    # Set the subject
    subject = f"TTO/TTIR/TTR breached incidents info for the last 24 hours {current_date}"
    mail.Subject = subject

    # Set the body with HTML formatting
    body = (
        "<p>Hi Team,</p>"
        "<p>Please find the attached TTO/TTIR/TTR breached incidents info for the last 24 hours.</p>"
        f"<p><span style='background-color:yellow'>There are {breach_count} breached incidents from the last 24 hours. </span></p>"
    )
    
    # Get the email inspector
    inspector = mail.GetInspector
    word_editor = inspector.WordEditor
    signature = mail.HTMLBody
    mail.HTMLBody = body + signature  # Append signature to the body

    # Add recipients
    mail.To = "gidcind_vpc_storage@dxc.com"
    mail.CC = "ifthikhar-ali.khan@dxc.com"

    # Attach the Excel file
    attachment = filename
    mail.Attachments.Add(attachment)

    # Display the email
    mail.Display()

    # Send the email
    # mail.Send()

    time.sleep(1)

    # Get the titles of all visible windows
    windows = pygetwindow.getAllTitles()
        
    # Check if any window title contains "Excel"
    for window_title in windows:
        if "TTO" in window_title:
            # Switch to the Excel window
            mail_window = pygetwindow.getWindowsWithTitle(window_title)
            if mail_window:
                time.sleep(1)
                mail_window[0].activate()
                break  # Stop searching for other Excel windows
