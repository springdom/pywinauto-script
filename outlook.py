import win32com.client


outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# "6" refers to the index of a folder - in this case,
# the inbox. You can change that number to reference
# any other folder
inbox = outlook.GetDefaultFolder(6) 

messages = inbox.Items
message = messages.GetFirst()
rec_time = message.CreationTime
body_content = message.body
subj_line = message.subject

print(message.subject, message.CreationTime)
print(body_content)

"""
while message:
    print(message.subject, message.CreationTime)
    message = messages.GetNext()
"""
