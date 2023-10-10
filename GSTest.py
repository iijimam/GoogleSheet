#https://www.kumilog.net/entry/2018/03/22/090000

from __future__ import print_function
import httplib2
import os

from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage
import configparser
config = configparser.RawConfigParser()
config.read("config.ini")

import sys
sys.path+=[config.get("section1","mgr"),config.get("section1","lib")]

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

SCOPES = 'https://www.googleapis.com/auth/spreadsheets'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'SheetTest2'


def get_credentials():
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir,
                                   'sheets.googleapis.com-python-quickstart.json')

    store = Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else:
            credentials = tools.run(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials


#spreadsheetId = '1YbPs8yJRiNCrMi7NlHDenAPqXATZYFAWM0jBMDB-8Js'
#rangeName = 'sales2!A1:C3'
def getData(spreadsheetId,rangeName):
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    discoveryUrl = ('https://sheets.googleapis.com/$discovery/rest?'
                    'version=v4')
    service = discovery.build('sheets', 'v4', http=http,
                            discoveryServiceUrl=discoveryUrl)

    result = service.spreadsheets().values().get(
            spreadsheetId=spreadsheetId, range=rangeName).execute()
    values = result.get('values', [])
    print(values)

#batchUpdateを実行
def updateData(spreadsheetId):
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    discoveryUrl = ('https://sheets.googleapis.com/$discovery/rest?'
                    'version=v4')
    service = discovery.build('sheets', 'v4', http=http,
                            discoveryServiceUrl=discoveryUrl)

    import iris
    import json
    obj=iris.cls("GoogleSheet.SheetTemplate").CreateJSON2()
    body=obj._ToJSON()
    body=json.loads(body)
    print(type(body))
    result = service.spreadsheets().values().batchUpdate(
        spreadsheetId=spreadsheetId, body=body).execute()

    print(result)

def clearSheet(spreadsheetId,rangeName):
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    discoveryUrl = ('https://sheets.googleapis.com/$discovery/rest?'
                    'version=v4')
    service = discovery.build('sheets', 'v4', http=http,
                            discoveryServiceUrl=discoveryUrl)
    result = service.spreadsheets().values().clear(
        spreadsheetId=spreadsheetId,range=rangeName).execute()
    print(result)

