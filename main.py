import os
import win32com.client

local_folder = "C:\\Users\\blpic\\PycharmProjects\\OutlookDownload"

outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")

inbox_folder = namespace.GetDefaultFolder(6)
out_folder = inbox_folder.Folders("OutFolder")

emails = out_folder.Items.Restrict("[Unread]=True AND [Subject] = 'data from workday'")

if len(emails) > 0:
    email = emails.GetFirst()
    if email.Attachments.Count > 0:
        attachment = email.Attachments.Item(1)
        attachment.SaveAsFile(os.path.join(local_folder, attachment.FileName))

        print("Attachment dl")
    else:
        print("No Attachment")
else:
    print("No unread email")