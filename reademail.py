from __future__ import print_function
import httplib2
import os
import base64
import email
from email import message
from email.parser import Parser
from apiclient import discovery
from apiclient import errors
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage
from pprint import pprint
import re
from lxml.html import tostring, fromstring, etree
from lxml.cssselect import CSSSelector
from mongoengine import StringField,ListField,DateTimeField,DynamicDocument,DynamicField,connect, Document, DictField


parser = Parser()

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

SCOPES = 'https://mail.google.com/'
CLIENT_SECRET_FILE = 'client_secret2.json'
APPLICATION_NAME = 'GMAIL TEST'

def get_credentials():

    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir,'gmail-python-quickstart.json')

    store = Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else:
            credentials = tools.run(flow,store)
        print('Storing credentials to' + credential_path)
    return credentials

def ListLabels(service, user_id='me'):

    try:
        response = service.users().labels().list(userId=user_id).execute()
        labels = response['labels']
        for label in labels:
            print('Label id: %s - Label name - %s' % (label['id'],label['name']))
        return labels
    except errors.HttpError as error:
        print('An error has occurred: %s' % error)


def DeleteMessage(service, msg_id, user_id='me'):

    try:
        service.users().messages().delete(userId=user_id, id=msg_id).execute()

    except errors.HttpError as error:
        print('An error occurred: %s' % error)


def GetMessage(service, msg_id, user_id='me'):

    try:
        message = service.users().messages().get(userId=user_id,id=msg_id).execute()
        #DeleteMessage(service=service,msg_id=msg_id)
        return message

    except errors.HttpError as error:
        print('An error occurred: %s' % error)


def GetMimeMessage(service, msg_id, user_id='me'):

    try:
        message = service.users().messages().get(userId=user_id, id=msg_id, format='full').execute()

        #pprint(message)
        msg_str = base64.urlsafe_b64decode(message['payload']['parts'][1]['body']['data'].encode('ASCII'))

        m = email.message_from_bytes(msg_str)

        return msg_str

    except errors.HttpError as error:
        print('An error occurred: %s' %error)


def ListMessagesMatchingQuery(service, query='',user_id='me'):

    try:
        response = service.users().messages().list(userId=user_id,q=query).execute()

        messages = []

        if 'messages' in response:
            messages.extend(response['messages'])

        while 'nextPageToken' in response:
            page_token = response['nextPageToken']
            response = service.users().messages().list(userId=user_id,q=query,pageToken=page_token).execute()

            messages.extend(response['messages'])

        return messages

    except errors.HttpError as error:
        print('An error occurred: %s' % error)


def ListMessagesWithLabels(service, label_ids=[], user_id='me'):

    try:
        response = service.users().messages().list(userId=user_id,labelIds=label_ids).execute()

        messages = []
        if 'messages' in response:
            messages.extend(response['messages'])

        while 'nextPageToken' in response:
            page_token = response['nextPageToken']
            response = service.users().messages().list(userId=user_id,labelIds=label_ids,pageToken=page_token).execute()

            messages.extend(response['messages'])

        return messages

    except errors.HttpError as error:
        print('An error occurred: %s' % error)

def connect_to_db():

    try:
        connect('work', host='mongodb://192.168.1.103:27017/work')
    except:
        print("Connection Failed")

class ChangeRequest(Document):

    rtc = StringField(required=True, primary_key=True)
    alert = StringField()
    type = StringField()
    summary = StringField()
    project = StringField()
    focus = StringField()
    status = StringField()
    owner = StringField()
    lockdown_date = DateTimeField(null=True)
    sizing_date = DateTimeField(null=True)
    uat = DateTimeField(null=True)
    deploy = DateTimeField(null=True)
    sizing = StringField()
    updates = DictField()




def convert_dates(date_str):

    from datetime import datetime

    if date_str != 'Unassigned':
        return datetime.strptime(date_str, '%A, %B %d, %Y')
    else:
        return None


def load_dict(items):

    cr = ChangeRequest()
    cr.rtc = items[1]
    cr.alert = items[0]
    cr.type = items[2]
    cr.summary = items[3]
    cr.project = items[4]
    cr.status = items[5]
    cr.focus = items[6]
    cr.owner = items[7]
    cr.lockdown_date = convert_dates(items[8])
    cr.sizing_date = convert_dates(items[9])
    cr.uat = convert_dates(items[10])
    cr.deploy = convert_dates(items[11])
    cr.sizing = items[12]

    cr.save()



   # for c in change_request_change.items():
       # print(c)
    #print(len(change_request_change))

def ParseChangeRequest():

    change_request_label = ['Label_1'] #Label where Lenovo Change Requests are found
    service = GetService()

    messages = ListMessagesWithLabels(service=service, label_ids=change_request_label)
    #print(type(messages), len(messages))
    #print(messages)
    span_selector = CSSSelector('span')
    span_items = []

    for message in messages:
        #print(message)
        msg = GetMimeMessage(service=service,msg_id=message['id'])
        #print(msg)
        parsed_message = fromstring(msg.decode('utf-8'))
        #print('Parsed message:', parsed_message)
        results = span_selector(parsed_message)
        #print(results)
        for result in results:
            #print(result.text)
            span_items.append(str(result.text).strip())
            #print(span_items)
        load_dict(span_items)
        span_items.clear()
            #DeleteMessage(service=service,msg_id=message['id'])

def GetService():

    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    service = discovery.build('gmail', 'v1', http=http)

    return service


def main():

    connect_to_db()
    ParseChangeRequest()

    label_ids = ['Label_1']
    #credentials = get_credentials()
    #print(type(credentials))
    #http = credentials.authorize(httplib2.Http())
    #service = discovery.build('gmail','v1',http=http)
    #parser = etree.HTMLParser()

    #ListLabels(service=service)


    #messages = ListMessagesWithLabels(service=service,label_ids=label_ids)
    #print(len(messages))
    #print(type(messages),len(messages))
    #print(messages)

    #html_messages = []
    #for message in messages:
        #print(message)
        #msg = GetMimeMessage(service=service,msg_id=message['id'])
       # print(msg)
        #html_messages.append(fromstring(msg.decode('utf-8')))


    #print(type(html_messages[0]))

    #html_text = tostring(html_messages[0]).decode('utf-8')
    #print(html_text)

    #sel = CSSSelector('span')


    #alert_items = []

    #results = sel(html_messages[0])
    #print(len(results))
    #for result in results:
       # alert_items.append(str(result.text).strip())


    #pprint(alert_items)
    #load_dict(alert_items)
    #DeleteMessage(service=service,msg_id=message['id'])

if __name__ == '__main__':
    main()

