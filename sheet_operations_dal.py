from gspread import authorize, Client, Spreadsheet, Worksheet

from oauth2client.service_account import ServiceAccountCredentials


class SheetOperations:
    def __init__(self) -> None:
        self.client = self.authenticate_google_api()

    @staticmethod
    def authenticate_google_api() -> Client:
        # set up credentials to access Google Sheets API
        scope = [
            "https://spreadsheets.google.com/feeds",
            "https://www.googleapis.com/auth/drive",
        ]
        creds = ServiceAccountCredentials.from_json_keyfile_name(
            "credentials.json", scope
        )
        client = authorize(creds)
        return client

    def connect_spreadsheet(self, sheet_url) -> Spreadsheet:
        # open the Google Sheet by its link
        sheet = self.client.open_by_url(sheet_url)
        return sheet

    def connect_worksheet(
        self, spreadsheet_client: Spreadsheet, worksheet_name
    ) -> Worksheet:
        # get the specific worksheet by its grid ID
        return spreadsheet_client.worksheet(worksheet_name)

    def get_worksheet(self, sheet_url, sheet_name) -> Worksheet:
        spreadsheet_client = self.connect_spreadsheet(sheet_url)
        return self.connect_worksheet(spreadsheet_client, sheet_name)

    @staticmethod
    def get_data_by_worksheet(worksheet: Worksheet):
        return worksheet.get_all_values()

    def get_sheet_data(self, sheet_url, sheet_name):
        spreadsheet_client = self.connect_spreadsheet(sheet_url)
        worksheet = self.connect_worksheet(spreadsheet_client, sheet_name)
        return worksheet.get_all_values()


class sheetOpsWithDayData(SheetOperations):
    def __init__(self, daywise_sheet_url, daywise_worksheet_name) -> None:
        self.daywise_sheet_url = daywise_sheet_url
        self.daywise_worksheet_name = daywise_worksheet_name
        self.daywise_data = []
        super().__init__()

    def get_daywise_details(self):
        sheet_data = self.get_sheet_data(
            self.daywise_sheet_url,
            self.daywise_worksheet_name,
        )
        self.daywise_data = [
            tup for tup in sheet_data if len(tup) >= 6 and all(tup[:6])
        ]
        return self.daywise_data
