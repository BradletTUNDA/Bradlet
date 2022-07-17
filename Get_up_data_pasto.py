from pathlib import Path #As we are going to deal with different file paths ...
import win32com.client

#Before interacting with outlook application...
#Create output folder for emails and attachements
output_dir = Path.cwd() / "Output"#Make at the same working directory with the current compte "Path.cwd()"
output_dir.mkdir(parents=True,exist_ok=True)#mkdirfor make directory create he folder from pathlib ; by setting parents to true pathlib will create any missing parents directory and if the output folder already exists, "exist_ok", path lib will araise error which we can egnore by setting exist_ok to True

#connect to outlook
outlook = win32com.client.Dispatch("outlook.Application")#"outlook.Application" is the application name ; from this application we want to get Namespace "MAPI", Messaging Application Programming Interface, the Microsoft outlook messaging API, for connecting to the interface our inbox

#connect to folder
inbox = outlook.GetDefaultFolder(6)#outlook come already with a couple of default folders, the number represent the type of folder, the number represent the inbox folder
#inbox = outlook.Folders("tudabradley@gmail.com").Folders("Inbox")

#Get messages
messages = inbox.Items
for message in range(5):
    #Get informations from messages ; subject, body, attachments for example
    subject = message.Subject
    body = message.body
    attachments = message.Attachments

    # Create separate folder for each message and respectif attachments
    target_folder = output_dir / str(subject)#Use pathlib to create new folder within the output_directory and use the subject as the folder name 
    target_folder.mkdir(parents=True,exist_ok=True)

    #Wtite email body to text file with pathlib
    Path(target_folder / "EMAIL_BODY.txt").write_text(str(body))

    #Save attachments
    for attachment in attachments:
        attachment.SaveAsFile(target_folder / str(attachment))#Save each attachment to the target_folder
