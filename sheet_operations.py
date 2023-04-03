from gspread import authorize, Client, Spreadsheet, Worksheet
from gspread.utils import column_letter_to_index, ValueInputOption
from gspread.cell import Cell

from oauth2client.service_account import ServiceAccountCredentials

from configurations import (
    REG_SHEET_URL, REG_WORKSHEET_NAME, SheetColumns, WHATSAPP_WAIT_TIME, TEST
)
from utils.messages import coguide_message
from utils.helpers import get_today_date_formatted, get_name_with_prefix
from typing import List

from pywhatkit import sendwhatmsg_instantly

formatted_date = get_today_date_formatted('%d %b %y')

class SheetOperations:
    def __init__(self) -> None:
        self.client = self.authenticate_google_api()

    @staticmethod
    def authenticate_google_api() -> Client:
        # set up credentials to access Google Sheets API
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
        client = authorize(creds)
        return client

    def connect_spreadsheet(self, sheet_url) -> Spreadsheet:
        # open the Google Sheet by its link
        sheet = self.client.open_by_url(sheet_url)
        return sheet

    def connect_worksheet(self, spreadsheet_client: Spreadsheet, worksheet_name) -> Worksheet:
        # get the specific worksheet by its grid ID
        return spreadsheet_client.worksheet(worksheet_name)

    def get_worksheet(self, sheet_url, sheet_name) -> Worksheet:
        spreadsheet_client = self.connect_spreadsheet(sheet_url)
        return self.connect_worksheet(spreadsheet_client, sheet_name)

    def get_sheet_data(self, sheet_url, sheet_name):
        spreadsheet_client = self.connect_spreadsheet(sheet_url)
        worksheet = self.connect_worksheet(spreadsheet_client, sheet_name)
        return worksheet.get_all_values()


def get_col_index(col: str):
    return column_letter_to_index(col)-1

def is_eligible_for_msg(item):
    return (
        item[get_col_index(SheetColumns.NAME_COL)] and
        item[get_col_index(SheetColumns.MSG_COL)] and
        item[get_col_index(SheetColumns.MSG_STATUS_COL)] == "FALSE"
    )

def get_name_index():
    return get_col_index(SheetColumns.NAME_COL)

def get_mobile_index():
    return get_col_index(SheetColumns.MOBILE_COL)

def get_name(item):
    return get_name_with_prefix(item[get_name_index()])

def get_mobile(item):
    return item[get_mobile_index()]

def sendmsg_and_update_list(item, index, cells_list: List):
    try:
        sendwhatmsg_instantly(
            phone_no="+91" + item[get_mobile_index()],
            message=coguide_message.format(name=get_name(item)),
            wait_time=WHATSAPP_WAIT_TIME,
            tab_close=True
        )
        cells_list.append(Cell(row=index+1, col=get_col_index(SheetColumns.MSG_STATUS_COL)+1, value="TRUE"))
        cells_list.append(Cell(row=index+1, col=get_col_index(SheetColumns.DATE_COL)+1, value=formatted_date))
    except Exception as e:
            print(f"failed to send message to {get_mobile(item)} due to {str(e)}")

def send_wa_msg_and_update_sheet(worksheet: Worksheet):
    data = worksheet.get_all_values()
    headers = data[0]
    cells_list = []
    print(f"Started sending {headers[get_col_index(SheetColumns.MSG_COL)]}")
    for index, item in enumerate(data):
        if is_eligible_for_msg(item):
            if TEST != 1:
                print("Script is in test mode")
                print("Starting candidate of Day 1", get_name(item), get_mobile(item))
                return
            print(f"Sending Message to {item[get_mobile_index()]}")
            sendmsg_and_update_list(item, index, cells_list)
    if len(cells_list):
        print(f"{headers[get_col_index(SheetColumns.MSG_COL)]} Messages Send to {len(cells_list)//2} candidates")
        worksheet.update_cells(cells_list, value_input_option=ValueInputOption.user_entered)
    else:
        print("Nothing to update")

def process_messages_and_update_sheet():
    sheet_operations = SheetOperations()
    worksheet = sheet_operations.get_worksheet(REG_SHEET_URL, REG_WORKSHEET_NAME)
    send_wa_msg_and_update_sheet(worksheet)
    return

if __name__=="__main__":
    process_messages_and_update_sheet()
