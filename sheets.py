#!/usr/bin/python

#from __future__ import print_function
import httplib2
import os

from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage
from nsetools import Nse

from twilio.rest import Client

# Your Account SID from twilio.com/console
account_sid = "AC3aec34ad6af5ec23401b7872e5152b28"
# Your Auth Token from twilio.com/console
auth_token  = "30fc5a14c0d75ac268f3c383c3efa7ec"

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/sheets.googleapis.com-python-quickstart.json
SCOPES = 'https://www.googleapis.com/auth/drive'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'Google Sheets API Python Quickstart'


def get_credentials():
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """
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
        else: # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials

def main():
    """Shows basic usage of the Sheets API.

    Creates a Sheets API service object and prints the names and majors of
    students in a sample spreadsheet:
    https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit
    """
    count = 2;
    nse = Nse()
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    discoveryUrl = ('https://sheets.googleapis.com/$discovery/rest?'
                    'version=v4')
    service = discovery.build('sheets', 'v4', http=http,
                              discoveryServiceUrl=discoveryUrl)

    spreadsheetId = '1yPn9AS9QhWyQprXWrI6LfdheKbH8tDbN3eHVk6MD1X0'

    rangeName = 'Actual!A2:J'
    value_input_option='RAW'
    
    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheetId, range=rangeName).execute()
    actual = result.get('values', [])
    #print values

    rangeName = 'Config!A2:C'
    value_input_option='RAW'
    
    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheetId, range=rangeName).execute()
    config = result.get('values', [])
    print config

    if not actual:
        print('No data found.')
    else:
        for row in actual:
            # Print columns A and E, which correspond to indices 0 and 4.
            q = nse.get_quote(str(row[0]))
            
            '''
            print 'Symbol: ',q['symbol']
            print 'Company Name: ',q['companyName']
            print 'Buy Price: ',q['buyPrice1']
            print 'Sell Price: ',q['sellPrice1']
            print 'Change: ',q['change']
            print 'Open: ',q['open']
            print 'High: ',q['dayHigh']
            print 'Low: ',q['dayLow']
            print 'Close: ',q['closePrice']
            print '% Change: ',q['pChange']
            print '52 wk high: ',q['high52']
            print '52 wk low: ',q['low52']
             
                body = {
                'values': values
                }
                rangeName = 'test!D2:G'
                result = service.spreadsheets().values().append(spreadsheetId=spreadsheetId,range=rangeName,valueInputOption=value_input_option,body=body).execute()
                print result
            '''
            
            if q['buyPrice1'] is None :
                values = [
                [
                    q['closePrice'], q['high52'], q['low52']
                ]
                ]
            else:
                values = [
                [
                    q['buyPrice1'], q['high52'], q['low52']
                ]
                ]
            
            body = {
            'values': values
            }
            rangeName = 'Actual!C'+str(count)+':E'
            result = service.spreadsheets().values().update(spreadsheetId=spreadsheetId,range=rangeName,valueInputOption=value_input_option,body=body).execute()
            
            if row[5] == 'Yes' :
                client = Client(account_sid, auth_token)

                call = client.calls.create(    
                to=config[int(row[9])-1][2],
                from_="+19722036935",
                url="http://demo.twilio.com/docs/voice.xml")
                #print config[int(row[9])-1][2]
                values = [
                [
                    'Notified!'
                ]
                ]
                body = {
                'values': values
                }
                rangeName = 'Actual!F'+str(count)+':F'
                result = service.spreadsheets().values().update(spreadsheetId=spreadsheetId,range=rangeName,valueInputOption=value_input_option,body=body).execute()
            count += 1

if __name__ == '__main__':
    main()
    print 'Successfully updated!'
