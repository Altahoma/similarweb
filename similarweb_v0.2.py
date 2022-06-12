from googleapiclient.discovery import build
from google.oauth2 import service_account
from urllib.request import Request, urlopen
import urllib.error
import json
import time
import pycountry
import os


sheet_name = 'your sheet name'
user_agent = 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10.15; rv:96.0) Gecko/20100101 Firefox/96.0'
spreadsheet_id = 'your spreadsheet id'


print(f'Current sheet: {sheet_name}')
start = int(input('Start: '))
end = int(input('End: '))


URL = "https://data.similarweb.com/api/v1/data?domain="
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SERVICE_ACCOUNT_FILE = 'keys.json'
creds = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
range_read = '{}!A{}:A{}'.format(sheet_name, start, end)
service = build('sheets', 'v4', credentials=creds)
sheet = service.spreadsheets()
result = sheet.values().get(spreadsheetId=spreadsheet_id, range=range_read).execute()
sites = result.get('values', [])
print(sites)


def parse(num, val):
    for i, site in enumerate(val):
        time.sleep(1)
        value_range_body = []
        range_update = '{}!B{}'.format(sheet_name, num)
        req = Request(URL + site[0], headers={'User-Agent': user_agent})
        try:
            web_byte = urlopen(req).read()
        except urllib.error.HTTPError as inst:
            if inst.__str__() == 'HTTP Error 403: Forbidden':
                value_range_body.append('Captcha')
                request = service.spreadsheets().values().update(spreadsheetId=spreadsheet_id, range=range_update,
                                                                 valueInputOption='USER_ENTERED',
                                                                 body={'values': [value_range_body]})
                request.execute()
                print('Captcha')
                os.system("say капча сука бля")
                input('Press to continue: Enter')
                parse(num=num, val=val[i:])
                break
            if inst.__str__() == 'HTTP Error 404: Not Found':
                value_range_body.append('')
                value_range_body.append('Not Found')
                request = service.spreadsheets().values().update(spreadsheetId=spreadsheet_id, range=range_update,
                                                                 valueInputOption='USER_ENTERED',
                                                                 body={'values': [value_range_body]})
                request.execute()
                print('Not Found')
                num += 1
            if inst.__str__() == 'HTTP Error 400: Bad Request':
                value_range_body.append('')
                value_range_body.append('Not Domain')
                request = service.spreadsheets().values().update(spreadsheetId=spreadsheet_id, range=range_update,
                                                                 valueInputOption='USER_ENTERED',
                                                                 body={'values': [value_range_body]})
                request.execute()
                print('Not Domain')
                num += 1
            continue
        except UnicodeEncodeError as uee:
            value_range_body.append('')
            value_range_body.append('Unicode Error')
            request = service.spreadsheets().values().update(spreadsheetId=spreadsheet_id, range=range_update,
                                                             valueInputOption='USER_ENTERED',
                                                             body={'values': [value_range_body]})
            request.execute()
            print('Unicode Error')
            num += 1
            continue

        # webpage = web_byte.decode('utf-8')
        data_json = json.loads(web_byte)

        value_range_body.append(data_json['SiteName'])
        try:
            value_range_body.append(data_json['Category'])
        except KeyError:
            value_range_body.append('Null')
            print('Null Category')
        value_range_body.append(data_json['CategoryRank']['Rank'])
        value_range_body.append(str(list(data_json['EstimatedMonthlyVisits'].values())[-1]))

        for country in data_json['TopCountryShares']:
            value_range_body.append(pycountry.countries.get(numeric=str("{:03}".format(country['Country']))).alpha_2)
            value_range_body.append("{:.2%}".format(country['Value']))

        for traffic, value in data_json['TrafficSources'].items():
            value_range_body.append("{:.2%}".format(value))

        request = service.spreadsheets().values().update(spreadsheetId=spreadsheet_id, range=range_update,
                                                         valueInputOption='USER_ENTERED',
                                                         body={'values': [value_range_body]})
        request.execute()
        print(data_json)
        num += 1


parse(num=start, val=sites)
print('Finish!')
