from googleapiclient.discovery import build
from google.oauth2 import service_account
import requests
from bs4 import BeautifulSoup
import time


sheet_name = 'your sheet name'
spreadsheet_id = 'your spreadsheet id'


print(f'Current sheet: {sheet_name}')
num = int(input('Start: '))
end = int(input('End: '))


SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SERVICE_ACCOUNT_FILE = 'keys.json'
creds = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
range_read = '{}!A{}:A{}'.format(sheet_name, num, end)
service = build('sheets', 'v4', credentials=creds)
sheet = service.spreadsheets()
result = sheet.values().get(spreadsheetId=spreadsheet_id, range=range_read).execute()
values = result.get('values', [])
print(values)


url = 'https://www.alexa.com/minisiteinfo/'

for i, site in enumerate(values):
    time.sleep(1)
    value_range_body = []
    range_update = '{}!B{}'.format(sheet_name, num)

    page = requests.get(url + site[0])
    soup = BeautifulSoup(page.content, "html.parser")
    results = soup.find(style="width:35%;")
    sites_elements = results.find_all("a", class_="Block truncation Link")
    for site_element in sites_elements:
        site_name = site_element.text
        value_range_body.append(site_name)
    print(f'{site} = {value_range_body}')

    request = service.spreadsheets().values().update(spreadsheetId=spreadsheet_id, range=range_update,
                                                     valueInputOption='USER_ENTERED',
                                                     body={'values': [value_range_body]})
    response = request.execute()
    num += 1
