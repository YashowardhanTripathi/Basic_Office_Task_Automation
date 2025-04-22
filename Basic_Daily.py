import time
import win32com.client
import pythoncom
import os
import pandas as pd
import tempfile
import re
import datetime
import threading

# Assuming these are your custom modules
import SampleApplication_Termination
import SampleApplication_Access
import GoAnywhereTool


def MFT_Error():
    pythoncom.CoInitialize()
    results = []

    try:
        print("In MFT Error Function")
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        inbox = outlook.GetDefaultFolder(6)
        messages = inbox.Items

        last_day = (datetime.datetime.now() - datetime.timedelta(hours=10)).strftime("%m/%d/%Y %H:%M %p")
        print(f"The Last Day Mails: {last_day}")

        unread_messages = messages.Restrict(f"[Unread] = True AND [ReceivedTime] >= '{last_day}'")
        message_list = [msg for msg in unread_messages]
        num_unread = len(message_list)
        print(f"In MFT Count num_unread: {num_unread}")

        if num_unread == 0:
            print("No unread messages found.")
            results.append("No unread messages found in MFT Error")
        else:
            for message in message_list:
                print(f"Processing message with subject: {message.Subject}")
                Target_Sub = "Error - [MFTPRD] - USON_INV_PAT_REF_STATUS_UPDATE"

                try:
                    attachments = message.Attachments
                    has_attachments = bool(attachments)
                    print(f"Has attachments: {has_attachments}")
                except Exception as e:
                    print(f"Error accessing attachments: {e}")
                    has_attachments = False

                if has_attachments and Target_Sub in message.Subject:
                    print(f"Processing Email: {message.Subject}")
                    processed = False  # Flag to track if we processed the email
                    for attachment in attachments:
                        print(f"Found Attachment: {attachment.FileName}")

                        with tempfile.NamedTemporaryFile(delete=False) as temp_file:
                            temp_filename = temp_file.name
                            attachment.SaveAsFile(temp_filename)

                        if attachment.FileName.endswith(".csv"):
                            df = pd.read_csv(temp_filename, sep=",")
                            print(f"CSV Data:\n{df}")
                            POCHeck = df[df["Type"] == "abc"]
                            if not POCHeck.empty:
                                df = df.applymap(
                                    lambda x: re.sub(r"[!@#$%^&*()_+={}\[\]:;\"'|<>.?/~`]", "", str(x)) if isinstance(x,
                                                                                                                      str) else x)
                                df = df.iloc[:, 2:]
                                df.to_csv(rf"\\Please add server/local path here\{attachment.FileName}",
                                          index=False)
                                reply = message.Reply()
                                recipients = ["abc@abc.com", "abc2@abc.com"]
                                reply.To = "; ".join(recipients)

                                reply.Body = (
                                    rf"This file {attachment.FileName} is Not ok. \n "
                                    rf"And The file is checked and cleaned saved in the \n "
                                    rf"path: localpath\{attachment.FileName}. \n "
                                    rf"GoAnywhere tool run sucessfully Please follow up with mail. \n "
                                    rf"Note:- This message is system generated \n Thank you"
                                )
                                reply.Send()
                                GoAnywhereTool.goAnywhere_tool()
                                print("Filtered CSV saved, mail sent, and tool has run")
                                results.append(f"NONPO Found in {attachment.FileName}, Process activated")
                                processed = True  # Mark as processed
                            else:
                                reply = message.Reply()
                                recipients = ["abc@abc.com", "abc2@abc.com"]
                                reply.To = "; ".join(recipients)
                                reply.Body = (
                                    f"This file {attachment.FileName} is ok. \n "
                                    f"No need to Worry \n\n\n "
                                    f"Note:- This message is system generated \n\n\n\n Thank you"
                                )
                                reply.Send()
                                print("PO found its OK")
                                processed = True  # Mark as processed

                        os.remove(temp_filename)

                    # Mark as read if processed (either NONPO or PO case)
                    if processed:
                        message.Unread = False
                        message.Save()
                        time.sleep(0.5)
                        print(f"Marked message with subject '{message.Subject}' as read, status: {message.Unread}")
                else:
                    print(f"No Attachment or Target Subject Found in: {message.Subject}")
                    results.append(f"No attachment or target subject found in {message.Subject}")

        print("Finished processing all unread messages in MFT_Error")
        return results

    except Exception as e:
        print(f"Error in MFT_Error: {e}")
        return [str(e)]
    finally:
        pythoncom.CoUninitialize()


def badGateway():
    pythoncom.CoInitialize()
    results = []

    try:
        # Connect to Outlook
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        abc = win32com.client.Dispatch("Outlook.Application")
        namespace = abc.GetNamespace("MAPI")
        inbox = outlook.GetDefaultFolder(6)  # 6 = Inbox
        messages = inbox.Items

        # Filter emails from the last 3 hours
        last_day = (datetime.datetime.now() - datetime.timedelta(hours=10)).strftime("%m/%d/%Y %H:%M %p")
        print(f"The last mails are: {last_day}")

        # Restrict to unread messages within the last 3 hours
        unread_messages = messages.Restrict(f"[Unread] = True AND [ReceivedTime] >= '{last_day}'")
        messages_list = [msg for msg in unread_messages]
        print(f"Length of unread messages is {len(messages_list)}")

        if len(messages_list) == 0:
            results.append("No mail found for GAError")
        else:
            for messages in messages_list:
                print("In the function")

                if "Error - [MFTPRD] - USON_INVOICE" in messages.Subject:
                    attachments = messages.Attachments
                    for attachment in attachments:
                        print(f"The attachment name is: {attachment.FileName}")

                        # Check if the attachment is a .log file
                        if attachment.FileName.endswith('.csv'):
                            # Define a path to save the attachment
                            save_path = os.path.join(os.getcwd(), attachment.FileName)

                            # Save the attachment to disk
                            attachment.SaveAsFile(save_path)
                            print(f"Saved attachment to: {save_path}")

                            # Read the .log file and search for "host cannot connect"
                            try:
                                with open(save_path, 'r', encoding='utf-8') as log_file:
                                    log_content = log_file.read()
                                    #  502  Bad Gateway make sure Bad request come only in a file or something more is need to watch out.
                                    pattern = r"\bBad Gateway\b"
                                    matches = re.findall(pattern, log_content, re.IGNORECASE)
                                    if matches:  # Case-insensitive search
                                        print(f"The attachment content is {log_content}")
                                        print(f"'Email sending to USON team {attachment.FileName}")
                                        try:
                                            # Assuming 'messages' is a single email item to reply to
                                            # Replace with your logic to get the specific email if needed
                                            # Example: messages = namespace.GetDefaultFolder(6).Items.Restrict("[Subject] = 'Error - [MFTPRD] - USON_INVOICE'").Item(1)

                                            reply = messages.Reply()
                                            reply.Subject = "Re: Error - [MFTPRD] - USON_INVOICE"
                                            recipients = ["abc@abc.com", "abc2@abc.com"]
                                            reply.To = "; ".join(recipients)

                                            # Use HTML for better formatting and consistent line breaks
                                            reply.HTMLBody = (
                                                "<html><body>"
                                                "<p>Dear Team,</p>"
                                                "<p>Can you please check if the above-listed record(s) are in PEOPLESOFT?</p>"
                                                "<ul>"
                                                "<li>If the record <b>is in PeopleSoft</b>, it should be listed as <b>Status 95</b> in Alevate.</li>"
                                                "<li>If the record <b>is not in PeopleSoft</b>, it should be listed as <b>Status 91</b> in Alevate, and it will be sent the next time the job runs and added to PeopleSoft.</li>"
                                                "<li>If the record <b>is in PeopleSoft</b> but listed as <b>Status 91</b> in Alevate, please inform us <b>immediately</b> so we can update it to Status 95 before the interface runs again.</li>"
                                                "</ul>"
                                                "<p>Thank you,<br> Team</p>"
                                                "</body></html>"
                                            )

                                            # Validate and add attachment
                                            save_path = Path(save_path)  # Convert to Path object for robust handling
                                            if save_path.is_file():
                                                reply.Attachments.Add(str(save_path))  # Convert Path to string
                                                print(f"Attached file: {save_path}")
                                            else:
                                                print(f"Warning: Attachment file not found at {save_path}")
                                                # Decide whether to proceed without attachment or raise an error
                                                # raise FileNotFoundError(f"Attachment file not found: {save_path}")

                                            # Send the email
                                            reply.Send()
                                            print("Email sent successfully")

                                            # Mark original email as read and save
                                            messages.Unread = False
                                            messages.Save()

                                            # Log result
                                            results.append(f"'Couldn't connect to host' found in {attachment.FileName}")

                                        except Exception as e:
                                            print(f"Error sending email: {e}")
                                            # Optionally log the error to a file or take other actions
                                    else:
                                        print(f"'Bad Gateway error:-  {attachment.FileName}")

                                        reply = messages.Reply()
                                        reply.Subject = "Action Required:: Error - [MFTPRD] - USON_INVOICE"
                                        recipients = ["abc@abc.com", "abc2@abc.com"]
                                        reply.To = "; ".join(recipients)
                                        reply.HTMLBody = (
                                            "<html><body>"
                                            "<p>Dear User,</p>"
                                            "<p>Error - [MFTPRD]</p>"
                                            "<p>Bad Gateway error <b>Not Found</b> in attachment please check it manually </p>"
                                            "<p>Thank you,<br>Team</p>"
                                            "</body></html>"
                                        )
                                        reply.Attachments.Add(save_path)
                                        reply.Send()
                                        messages.Unread = False
                                        messages.Save()
                                        results.append(f"'Couldn't connect to host' not found in {attachment.FileName}")
                            except Exception as e:
                                print(f"Error reading the .log file: {e}")
                                results.append(f"Error reading {attachment.FileName}: {str(e)}")
                            finally:
                                # Clean up the temporary file
                                if os.path.exists(save_path):
                                    os.remove(save_path)
                                    print(f"Deleted temporary file: {save_path}")
    except Exception as e:
        print(f"Error in GAHost_NotConnect: {e}")
        results.append(f"Error in GAHost_NotConnect: {str(e)}")
    finally:
        pythoncom.CoUninitialize()

    return results


def GAHost_NotConnect():
    pythoncom.CoInitialize()
    results = []

    try:
        # Connect to Outlook
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        abc = win32com.client.Dispatch("Outlook.Application")
        namespace = abc.GetNamespace("MAPI")
        inbox = outlook.GetDefaultFolder(6)  # 6 = Inbox
        messages = inbox.Items

        # Filter emails from the last 3 hours
        last_day = (datetime.datetime.now() - datetime.timedelta(hours=10)).strftime("%m/%d/%Y %H:%M %p")
        print(f"The last mails are: {last_day}")

        # Restrict to unread messages within the last 3 hours
        unread_messages = messages.Restrict(f"[Unread] = True AND [ReceivedTime] >= '{last_day}'")
        messages_list = [msg for msg in unread_messages]
        print(f"Length of unread messages is {len(messages_list)}")

        if len(messages_list) == 0:
            results.append("No mail found for GAError")
        else:
            for messages in messages_list:
                print("In the function")
                # "Project ECM_DEA222EmailProcessing Failed" 1
                if "Project ECM_DEA222EmailProcessing Failed" in messages.Subject:
                    attachments = messages.Attachments
                    for attachment in attachments:
                        print(f"The attachment name is: {attachment.FileName}")

                        # Check if the attachment is a .log file
                        if attachment.FileName.endswith('.log'):
                            # Define a path to save the attachment
                            save_path = os.path.join(os.getcwd(), attachment.FileName)

                            # Save the attachment to disk
                            attachment.SaveAsFile(save_path)
                            print(f"Saved attachment to: {save_path}")

                            # Read the .log file and search for "host cannot connect"
                            try:
                                with open(save_path, 'r', encoding='utf-8') as log_file:
                                    log_content = log_file.read()
                                    pattern = r"\b(?:couldn't|could\s+not)\s+connect\s+to\s+(?:host|server)\b"
                                    matches = re.findall(pattern, log_content, re.IGNORECASE)
                                    if matches:  # Case-insensitive search
                                        print(f"The attachment content is {log_content}")
                                        print(f"'Couldn't connect to host' found in {attachment.FileName}")

                                        # Move the email to the Deleted Items folder
                                        deleted_items = namespace.GetDefaultFolder(3)  # 3 = Deleted Items
                                        messages.Move(deleted_items)
                                        messages.Unread = False
                                        messages.Save()
                                        results.append(f"'Couldn't connect to host' found in {attachment.FileName}")
                                    else:
                                        print(f"'Couldn't connect to host' not found in {attachment.FileName}")
                                        print(f"The attachment content is {log_content}")
                                        reply = messages.Reply()
                                        reply.Subject = "Action Required:: Project ECM_DEA222EmailProcessing Failed"
                                        recipients = ["abc@abc.com", "abc2@abc.com"]
                                        reply.To = "; ".join(recipients)
                                        reply.Body = f"Project EmailProcessing Failed \n (This one is new please go through the mail log file) {attachment.FileName}"
                                        reply.Send()
                                        messages.Unread = False
                                        messages.Save()
                                        results.append(f"'Couldn't connect to host' not found in {attachment.FileName}")
                            except Exception as e:
                                print(f"Error reading the .log file: {e}")
                                results.append(f"Error reading {attachment.FileName}: {str(e)}")
                            finally:
                                # Clean up the temporary file
                                if os.path.exists(save_path):
                                    os.remove(save_path)
                                    print(f"Deleted temporary file: {save_path}")

                # Scheduled job HR Email Ingestion failed  2
                if "Scheduled job HR Email Ingestion failed" in messages.Subject:
                    attachments = messages.Attachments
                    for attachment in attachments:
                        print(f"The attachment name is: {attachment.FileName}")

                        # Check if the attachment is a .log file
                        if attachment.FileName.endswith('.log'):
                            # Define a path to save the attachment
                            save_path = os.path.join(os.getcwd(), attachment.FileName)

                            # Save the attachment to disk
                            attachment.SaveAsFile(save_path)
                            print(f"Saved attachment to: {save_path}")

                            # Read the .log file and search for "host cannot connect"
                            try:
                                with open(save_path, 'r', encoding='utf-8') as log_file:
                                    log_content = log_file.read()
                                    pattern = r"\b(?:couldn't|could\s+not)\s+connect\s+to\s+(?:host|server)\b"
                                    matches = re.findall(pattern, log_content, re.IGNORECASE)
                                    if matches:  # Case-insensitive search
                                        print(f"The attachment content is {log_content}")
                                        print(f"'Couldn't connect to host' found in {attachment.FileName}")

                                        # Move the email to the Deleted Items folder
                                        deleted_items = namespace.GetDefaultFolder(3)  # 3 = Deleted Items
                                        messages.Move(deleted_items)
                                        messages.Unread = False
                                        messages.Save()
                                        results.append(f"'Couldn't connect to host' found in {attachment.FileName}")
                                    else:
                                        print(f"'Couldn't connect to host' not found in {attachment.FileName}")
                                        print(f"The attachment content is {log_content}")
                                        reply = messages.Reply()
                                        reply.Subject = "Action Required:: Scheduled job HR Email Ingestion failed"
                                        recipients = recipients = ["abc@abc.com", "abc2@abc.com"]
                                        reply.To = "; ".join(recipients)
                                        reply.Body = f"Scheduled job HR Email Ingestion failed \n (This one is new please go through the mail log file) {attachment.FileName}"
                                        reply.Send()
                                        messages.Unread = False
                                        messages.Save()
                                        results.append(f"'Couldn't connect to host' not found in {attachment.FileName}")
                            except Exception as e:
                                print(f"Error reading the .log file: {e}")
                                results.append(f"Error reading {attachment.FileName}: {str(e)}")
                            finally:
                                # Clean up the temporary file
                                if os.path.exists(save_path):
                                    os.remove(save_path)
                                    print(f"Deleted temporary file: {save_path}")

                # Scheduled job MMS_Email_NewAccountSetup failed  3
                if "Scheduled job MMS_Email_NewAccountSetup failed" in messages.Subject:
                    attachments = messages.Attachments
                    for attachment in attachments:
                        print(f"The attachment name is: {attachment.FileName}")

                        # Check if the attachment is a .log file
                        if attachment.FileName.endswith('.log'):
                            # Define a path to save the attachment
                            save_path = os.path.join(os.getcwd(), attachment.FileName)

                            # Save the attachment to disk
                            attachment.SaveAsFile(save_path)
                            print(f"Saved attachment to: {save_path}")

                            # Read the .log file and search for "host cannot connect"
                            try:
                                with open(save_path, 'r', encoding='utf-8') as log_file:
                                    log_content = log_file.read()
                                    pattern = r"\b(?:couldn't|could\s+not)\s+connect\s+to\s+(?:host|server)\b"
                                    matches = re.findall(pattern, log_content, re.IGNORECASE)
                                    if matches:  # Case-insensitive search
                                        print(f"The attachment content is {log_content}")
                                        print(
                                            f"'Couldn't connect to host' found in {attachment.FileName}")

                                        # Move the email to the Deleted Items folder
                                        deleted_items = namespace.GetDefaultFolder(
                                            3)  # 3 = Deleted Items
                                        messages.Move(deleted_items)
                                        messages.Unread = False
                                        messages.Save()
                                        results.append(
                                            f"'Couldn't connect to host' found in {attachment.FileName}")
                                    else:
                                        print(
                                            f"'Couldn't connect to host' not found in {attachment.FileName}")
                                        print(f"The attachment content is {log_content}")
                                        reply = messages.Reply()
                                        reply.Subject = "Action Required:: Scheduled job MMS_Email_NewAccountSetup failed"
                                        recipients = ["abc@abc.com", "abc2@abc.com"]
                                        reply.To = "; ".join(recipients)
                                        reply.Body = f"Scheduled job MMS_Email_NewAccountSetup failed \n (This one is new please go through the mail log file) {attachment.FileName}"
                                        reply.Send()
                                        messages.Unread = False
                                        messages.Save()
                                        results.append(
                                            f"'Couldn't connect to host' not found in {attachment.FileName}")
                            except Exception as e:
                                print(f"Error reading the .log file: {e}")
                                results.append(f"Error reading {attachment.FileName}: {str(e)}")
                            finally:
                                # Clean up the temporary file
                                if os.path.exists(save_path):
                                    os.remove(save_path)
                                    print(f"Deleted temporary file: {save_path}")

                # PROD Email Handler Error email_mhs  4
                if "PROD Email Handler Error email_mhs" in messages.Subject:
                    attachments = messages.Attachments
                    for attachment in attachments:
                        print(f"The attachment name is: {attachment.FileName}")

                        # Check if the attachment is a .log file
                        if attachment.FileName.endswith('.log'):
                            # Define a path to save the attachment
                            save_path = os.path.join(os.getcwd(), attachment.FileName)

                            # Save the attachment to disk
                            attachment.SaveAsFile(save_path)
                            print(f"Saved attachment to: {save_path}")

                            # Read the .log file and search for "host cannot connect"
                            try:
                                with open(save_path, 'r', encoding='utf-8') as log_file:
                                    log_content = log_file.read()
                                    pattern = r"\b(?:couldn't|could\s+not)\s+connect\s+to\s+(?:host|server)\b"
                                    matches = re.findall(pattern, log_content, re.IGNORECASE)
                                    if matches:  # Case-insensitive search
                                        print(f"The attachment content is {log_content}")
                                        print(
                                            f"'Couldn't connect to host' found in {attachment.FileName}")

                                        # Move the email to the Deleted Items folder
                                        deleted_items = namespace.GetDefaultFolder(
                                            3)  # 3 = Deleted Items
                                        messages.Move(deleted_items)
                                        messages.Unread = False
                                        messages.Save()
                                        results.append(
                                            f"'Couldn't connect to host' found in {attachment.FileName}")
                                    else:
                                        print(
                                            f"'Couldn't connect to host' not found in {attachment.FileName}")
                                        print(f"The attachment content is {log_content}")
                                        reply = messages.Reply()
                                        reply.Subject = "Action Required:: PROD Email Handler Error email_mhs"
                                        recipients = ["abc@abc.com", "abc2@abc.com"]
                                        reply.To = "; ".join(recipients)
                                        reply.Body = f"PROD Email Handler Error email_mhs \n (This one is new please go through the mail log file) {attachment.FileName}"
                                        reply.Send()
                                        messages.Unread = False
                                        messages.Save()
                                        results.append(
                                            f"'Couldn't connect to host' not found in {attachment.FileName}")
                            except Exception as e:
                                print(f"Error reading the .log file: {e}")
                                results.append(f"Error reading {attachment.FileName}: {str(e)}")
                            finally:
                                # Clean up the temporary file
                                if os.path.exists(save_path):
                                    os.remove(save_path)
                                    print(f"Deleted temporary file: {save_path}")

                # PROD Email Handler Error  5
                if "PROD Email Handler Error" in messages.Subject:
                    attachments = messages.Attachments
                    for attachment in attachments:
                        print(f"The attachment name is: {attachment.FileName}")

                        # Check if the attachment is a .log file
                        if attachment.FileName.endswith('.log'):
                            # Define a path to save the attachment
                            save_path = os.path.join(os.getcwd(), attachment.FileName)

                            # Save the attachment to disk
                            attachment.SaveAsFile(save_path)
                            print(f"Saved attachment to: {save_path}")

                            # Read the .log file and search for "host cannot connect"
                            try:
                                with open(save_path, 'r', encoding='utf-8') as log_file:
                                    log_content = log_file.read()
                                    pattern = r"\b(?:couldn't|could\s+not)\s+connect\s+to\s+(?:host|server)\b"
                                    matches = re.findall(pattern, log_content, re.IGNORECASE)
                                    if matches:  # Case-insensitive search
                                        print(f"The attachment content is {log_content}")
                                        print(
                                            f"'Couldn't connect to host' found in {attachment.FileName}")

                                        # Move the email to the Deleted Items folder
                                        deleted_items = namespace.GetDefaultFolder(
                                            3)  # 3 = Deleted Items
                                        messages.Move(deleted_items)
                                        messages.Unread = False
                                        messages.Save()
                                        results.append(
                                            f"'Couldn't connect to host' found in {attachment.FileName}")
                                    else:
                                        print(
                                            f"'Couldn't connect to host' not found in {attachment.FileName}")
                                        print(f"The attachment content is {log_content}")
                                        reply = messages.Reply()
                                        reply.Subject = "Action Required:: PROD Email Handler Error"
                                        recipients = ["abc@abc.com", "abc2@abc.com"]
                                        reply.To = "; ".join(recipients)
                                        reply.Body = f"PROD Email Handler Error \n (This one is new please go through the mail log file) {attachment.FileName}"
                                        reply.Send()
                                        messages.Unread = False
                                        messages.Save()
                                        results.append(
                                            f"'Couldn't connect to host' not found in {attachment.FileName}")
                            except Exception as e:
                                print(f"Error reading the .log file: {e}")
                                results.append(f"Error reading {attachment.FileName}: {str(e)}")
                            finally:
                                # Clean up the temporary file
                                if os.path.exists(save_path):
                                    os.remove(save_path)
                                    print(f"Deleted temporary file: {save_path}")

                # PROD Rebate Email Handler Error  6
                if "PROD Rebate Email Handler Error" in messages.Subject:
                    attachments = messages.Attachments
                    for attachment in attachments:
                        print(f"The attachment name is: {attachment.FileName}")

                        # Check if the attachment is a .log file
                        if attachment.FileName.endswith('.log'):
                            # Define a path to save the attachment
                            save_path = os.path.join(os.getcwd(), attachment.FileName)

                            # Save the attachment to disk
                            attachment.SaveAsFile(save_path)
                            print(f"Saved attachment to: {save_path}")

                            # Read the .log file and search for "host cannot connect"
                            try:
                                with open(save_path, 'r', encoding='utf-8') as log_file:
                                    log_content = log_file.read()
                                    pattern = r"\b(?:couldn't|could\s+not)\s+connect\s+to\s+(?:host|server)\b"
                                    matches = re.findall(pattern, log_content, re.IGNORECASE)
                                    if matches:  # Case-insensitive search
                                        print(f"The attachment content is {log_content}")
                                        print(
                                            f"'Couldn't connect to host' found in {attachment.FileName}")

                                        # Move the email to the Deleted Items folder
                                        deleted_items = namespace.GetDefaultFolder(
                                            3)  # 3 = Deleted Items
                                        messages.Move(deleted_items)
                                        messages.Unread = False
                                        messages.Save()
                                        results.append(
                                            f"'Couldn't connect to host' found in {attachment.FileName}")
                                    else:
                                        print(
                                            f"'Couldn't connect to host' not found in {attachment.FileName}")
                                        print(f"The attachment content is {log_content}")
                                        reply = messages.Reply()
                                        reply.Subject = "Action Required:: PROD Rebate Email Handler Error"
                                        recipients = ["abc@abc.com", "abc2@abc.com"]
                                        reply.To = "; ".join(recipients)
                                        reply.Body = f"PROD Rebate Email Handler Error \n (This one is new please go through the mail log file) {attachment.FileName}"
                                        reply.Send()
                                        messages.Unread = False
                                        messages.Save()
                                        results.append(
                                            f"'Couldn't connect to host' not found in {attachment.FileName}")
                            except Exception as e:
                                print(f"Error reading the .log file: {e}")
                                results.append(f"Error reading {attachment.FileName}: {str(e)}")
                            finally:
                                # Clean up the temporary file
                                if os.path.exists(save_path):
                                    os.remove(save_path)
                                    print(f"Deleted temporary file: {save_path}")

    except Exception as e:
        print(f"Error in GAHost_NotConnect: {e}")
        results.append(f"Error in GAHost_NotConnect: {str(e)}")
    finally:
        pythoncom.CoUninitialize()

    return results

def alevateAccess():
    pythoncom.CoInitialize()
    results = []

    try:
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        inbox = outlook.GetDefaultFolder(6)
        messages = inbox.Items

        last_day = (datetime.datetime.now() - datetime.timedelta(hours=10)).strftime("%m/%d/%Y %H:%M %p")
        print(f"The Last Day Mails: {last_day}")

        unread_messages = messages.Restrict(f"[Unread] = True AND [ReceivedTime] >= '{last_day}'")
        message_list = [msg for msg in unread_messages]
        print(f"Unread Messages Count: {len(message_list)}")

        if len(message_list) == 0:
            results.append("No unread messages found in alevateAccess")
        else:
            for message in message_list:
                if "Catalog Task" in message.Subject and "TCS_Serrala_Support" in message.Subject:
                    file_content = message.Body.lower()
                    if ("create a new profile" in file_content or "create new profile" in file_content or
                            "create profile" in file_content or "create a profile" in file_content or
                            "add new alevate user" in file_content or "add new" in file_content or
                            "new user" in file_content or "new user access" in file_content):

                        match = re.search(r"email\s*\n\s*([\w\.-]+@[\w\.-]+\.\w+)", file_content)
                        if match:
                            email = match.group(1)
                            name = (email.split("@")[0]).upper()
                            print(f"Found Email: {email}, Name: {name}")
                        else:
                            print("No mail ID found")
                            results.append("No email found in message")
                            continue

                        match = re.search(r"what level of access is needed\?\s*\n\s*(.*)", file_content)
                        if match:
                            access_level = match.group(1)
                            print(f"Access Level: {access_level}")
                            if "view" in access_level.lower():
                                Auths = """[{"TYPE":"AUTH","VALUE":["MSH_PS","MSH_PSFT_B"]}]"""
                                print(f"Auths: {Auths}")
                            else:
                                Auths = None
                        else:
                            print("No access level found")
                            Auths = None

                        CoreContentAccess = SampleApplication_Access.AlevateAccessSelenium(email, name, Auths)
                        reply = message.Reply()
                        reply.Subject = "User Alevate Access"
                        recipients = ["abc@abc.com", "abc2@abc.com"]
                        reply.To = "; ".join(recipients)
                        reply.Body = f"New user Status. Here is details: {CoreContentAccess}"
                        reply.Send()
                        results.append(f"Alevate Access processed for {email}: {CoreContentAccess}")
                        # Mark the message as read and save
                        message.Unread = False
                        message.Save()
                        time.sleep(0.5)  # Small delay to allow Outlook to sync
                        print(f"Marked message with subject '{message.Subject}' as read, status: {message.Unread}")

        print("Finished processing alevateAccess")
        return results

    except Exception as e:
        print(f"Error in alevateAccess: {e}")
        return [str(e)]
    finally:
        pythoncom.CoUninitialize()


def terminateUser():
    pythoncom.CoInitialize()
    results = []

    def run_terminate(email, result_list):
        pythoncom.CoInitialize()
        try:
            print(f"Starting termination for {email}, this may take ~1 minute...")
            CoreContent = SampleApplication_Termination.TerminateUser(email)
            print(f"Completed termination for {email}")
            result_list.append(f"Termination processed for {email}: {CoreContent}")

            # Create a new email instead of replying
            outlook = win32com.client.Dispatch("Outlook.Application")
            new_mail = outlook.CreateItem(0)  # 0 = olMailItem
            new_mail.Subject = "User Alevate Terminate Status"
            recipients = ["abc@abc.com", "abc2@abc.com"]
            new_mail.To = "; ".join(recipients)
            new_mail.Body = f"User terminated: {CoreContent}"
            new_mail.Send()

        except Exception as e:
            result_list.append(f"Error terminating {email}: {e}")
        finally:
            pythoncom.CoUninitialize()

    try:
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        inbox = outlook.GetDefaultFolder(6)
        messages = inbox.Items

        last_day = (datetime.datetime.now() - datetime.timedelta(hours=10)).strftime("%m/%d/%Y %H:%M %p")
        print(f"The Last Day Mails: {last_day}")

        unread_messages = messages.Restrict(f"[Unread] = True AND [ReceivedTime] >= '{last_day}'")
        message_list = [msg for msg in unread_messages]
        print(f"Unread Messages Count: {len(message_list)}")

        if len(message_list) == 0:
            results.append("No unread messages found in terminateUser")
        else:
            threads = []
            for message in message_list:
                if "Catalog Task" in message.Subject and "TCS_Serrala_Support" in message.Subject:
                    print(f"Subject Found: {message.Subject}")
                    file_content = message.Body
                    if "Terminated" in file_content:
                        match = re.search(r"Email\s*\n\s*([\w\.-]+@[\w\.-]+\.\w+)", file_content)
                        if match:
                            email = match.group(1)
                            print(f"Email for termination: {email}")
                            t = threading.Thread(target=run_terminate, args=(email, results))
                            threads.append(t)
                            t.start()
                            # Mark the message as read and save
                            message.Unread = False
                            message.Save()
                            time.sleep(0.5)  # Small delay to allow Outlook to sync
                            print(f"Marked message with subject '{message.Subject}' as read, status: {message.Unread}")

            for t in threads:
                t.join()

        print("Finished processing terminateUser")
        return results

    except Exception as e:
        print(f"Error in terminateUser: {e}")
        return [str(e)]
    finally:
        pythoncom.CoUninitialize()


def main():
    print("Starting main")

    mft_results = MFT_Error()
    print(f"MFT_Error Results: {mft_results}")

    GAResult = GAHost_NotConnect()
    print(f"GAError result is : {GAResult}")

    BadGateway = badGateway()
    print(f"BadGate way message:- {BadGateway}")

    access_results = alevateAccess()
    print(f"AlevateAccess Results: {access_results}")

    terminate_results = terminateUser()
    print(f"TerminateUser Results: {terminate_results}")

    print("Finished main")


if __name__ == "__main__":
    main()