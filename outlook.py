import win32com.client
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

inbox = outlook.GetDefaultFolder(6)

messages = inbox.Items
messages.Sort("[ReceivedTime]", True)
msg = messages.Find("[SenderEmailAddress]='rajeev.bhosle.222@gmail.com'")
a=msg.subject
print(a)
