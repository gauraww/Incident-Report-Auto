import win32com.client
from datetime import datetime

def sendmail(filename, breach_count):

    # Create an instance of the Outlook application
    outlook = win32com.client.Dispatch("Outlook.Application")

    # Create a new email
    mail = outlook.CreateItem(0)
    current_date = datetime.now().strftime("%d %B %Y")

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

    return True


