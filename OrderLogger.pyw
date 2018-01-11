from __future__ import print_function
import httplib2
import os
import io
import base64
from apiclient import errors, discovery

from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.audio import MIMEAudio
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import mimetypes

from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage

from appJar import gui

import json
import datetime

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/gmail-python-quickstart.json
SCOPES = 'https://www.googleapis.com/auth/gmail.send'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'Gmail API Python Quickstart'


def get_credentials():
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    print(credential_dir)
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir,
                                   'gmail-python-quickstart.json')

    store = Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else: # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials

def create_message_with_attachment(sender, to, subject, message_text, file):
  """Create a message for an email.

  Args:
    sender: Email address of the sender.
    to: Email address of the receiver.
    subject: The subject of the email message.
    message_text: The text of the email message.
    file: The path to the file to be attached.

  Returns:
    An object containing a base64url encoded email object.
  """
  message = MIMEMultipart()
  message['to'] = to
  message['from'] = sender
  message['subject'] = subject

  msg = MIMEText(message_text)
  message.attach(msg)

  content_type, encoding = mimetypes.guess_type(file)

  if content_type is None or encoding is not None:
    content_type = 'application/octet-stream'
  main_type, sub_type = content_type.split('/', 1)
  if main_type == 'text':
    fp = open(file, 'rb')
    msg = MIMEText(fp.read(), _subtype=sub_type)
    fp.close()
  elif main_type == 'image':
    fp = open(file, 'rb')
    msg = MIMEImage(fp.read(), _subtype=sub_type)
    fp.close()
  elif main_type == 'audio':
    fp = open(file, 'rb')
    msg = MIMEAudio(fp.read(), _subtype=sub_type)
    fp.close()
  else:
    fp = open(file, 'rb')
    msg = MIMEBase(main_type, sub_type)
    msg.set_payload(fp.read())
    fp.close()
  filename = os.path.basename(file)
  msg.add_header('Content-Disposition', 'attachment', filename=filename)
  message.attach(msg)

  return {'raw': base64.urlsafe_b64encode(message.as_string())}

def send_message(service, user_id, message):
  """Send an email message.

  Args:
    service: Authorized Gmail API service instance.
    user_id: User's email address. The special value "me"
    can be used to indicate the authenticated user.
    message: Message to be sent.

  Returns:
    Sent Message.
  """
  global app
  print(type(message))
  try:
    message = (service.users().messages().send(userId=user_id, body=message)
               .execute())
    print ("Message Id: %s" % message['id'])
    print(base64.urlsafe_b64decode(message['raw']))
    app.infoBox("Email Notification","Email was sent succesfully")
    return message

  except errors.HttpError as error:
    print ("An error occurred: %s" % error)
    app.errorBox("Email Notification","Error occured while sending message")

def sendMessage(to,subject,attachment):

    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    service = discovery.build('gmail', 'v1', http=http)

    with open('emailBody.txt', 'r') as myfile:
      bodyText = myfile.read()

    message = create_message_with_attachment("lewissavage88@gmail.com",to,subject,bodyText,attachment)
    send_message(service,"me",message)

def attachFile(widjet):
    global app
    file = app.openBox(title=None, dirName=None, fileTypes=None, asFile=True, parent=None)
    global filePath
    filePath = file.name
    app.setButton("Attach File",filePath)


def sendButton(widjet):
    print(filePath)
    global app
    add = contacts[app.getOptionBox("Select company to order from")]
    subject = "Enibas order %s" % datetime.datetime.now().strftime('%m/%d/%Y')
    sendMessage(add, subject, filePath)

def find_nth(haystack, needle, n):
    start = haystack.find(needle)
    while start >= 0 and n > 1:
        start = haystack.find(needle, start+len(needle))
        n -= 1
    return start

def loadContacts():
  global contacts
  with open("contacts.json",'r') as f:
    try:
      loadedContacts = json.load(f)
      print(loadedContacts);
      contacts = loadedContacts
    except ValueError as e:
      print("No contacts saved")
  
def saveContacts():
  global contacts
  with open("contacts.json", 'w') as f:
    f.write(json.dumps(contacts))

def addDeleteContact(w):
  global app
  global contacts
  name = app.getEntry("Contact Name")
  add = app.getEntry("Contact Address")
  if (name == "") or (add == ""):
    return
  else:
    if w == "Delete Contact":
      try:
        del contacts[name]
      except KeyError as e:
        print("Contact does not exist")
      saveContacts()
    else:
      contacts[name] = add
      saveContacts()
    app.clearListBox("ContactList", callFunction=False)
    app.addListItems("ContactList",contacts);
    updateInterfaces()

def changeContact(w):
  global app
  try:
    nm = app.getListBox("ContactList")[0]
    ad = contacts[str(app.getListBox("ContactList")[0])]
    app.setEntry("Contact Name", nm)
    app.setEntry("Contact Address", ad)
  except IndexError as e:
    print("No contacts to choose from")

def updateInterfaces():
  global app
  global contacts
  app.changeOptionBox("Select company to order from",contacts,callFunction=False)
  app.changeOptionBox("Company",contacts,callFunction=False)

contacts = {}
filePath = ""

if __name__ == '__main__':
    loadContacts()
    app = gui()
    app.setTitle("Workshop Ordering")
    app.setFont(10,"Arial")
    app.startTabbedFrame("TabbedFrame")
    app.setResizable(canResize=False)
    app.startTab("Order")

    app.startLabelFrame("Ordering")

    app.addLabelOptionBox("Select company to order from",contacts)
    app.setOptionBoxWidth("Select company to order from",10);
    app.addButton("Attach File",attachFile)
    app.setButtonSticky("Attach File","left")

    app.addButton("Send", sendButton)
    app.setButtonSticky("Send","left")

    app.stopLabelFrame()

    app.stopTab()

    app.startTab("Review")

    app.addLabelOptionBox("Company",contacts)

    app.stopTab()

    app.startTab("Contacts")

    app.addListBox("ContactList",contacts)
    app.setListBoxFunction("ContactList",changeContact)
    app.addLabelEntry("Contact Name")
    app.addLabelEntry("Contact Address")
    app.addButtons(["Delete Contact","Add Contact"],addDeleteContact)

    app.stopTab()
    app.stopTabbedFrame()
    app.go()
