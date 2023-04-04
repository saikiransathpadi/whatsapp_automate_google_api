from gspread import Worksheet
from gspread.utils import column_letter_to_index, ValueInputOption
from gspread.cell import Cell

from sheet_operations_dal import sheetOpsWithDayData

from configurations import (
    REG_SHEET_URL,
    REG_WORKSHEET_NAME,
    SheetColumns,
    WHATSAPP_WAIT_TIME,
    TEST,
    DAYWISE_WORKSHEET_NAME,
)
from utils.messages import coguide_message
from utils.helpers import get_today_date_formatted, get_name_with_prefix, print_colored
from typing import List

from pywhatkit import sendwhatmsg_instantly

formatted_date = get_today_date_formatted("%d %b %y")


def get_col_index(col: str):
    return column_letter_to_index(col) - 1


def is_eligible_for_msg(item, msg_col, msg_status_col):
    return (
        item[get_col_index(SheetColumns.NAME_COL)]
        and item[get_col_index(msg_col)]
        and item[get_col_index(msg_status_col)] == "FALSE"
    )


def get_name_index():
    return get_col_index(SheetColumns.NAME_COL)


def get_mobile_index():
    return get_col_index(SheetColumns.MOBILE_COL)


def get_name(item):
    return get_name_with_prefix(item[get_name_index()])


def get_mobile(item):
    return item[get_mobile_index()]


def sendmsg_and_update_list(
    item, index, cells_list: List, mesage, msg_status_col, date_col
):
    try:
        sendwhatmsg_instantly(
            phone_no="+91" + get_mobile(item),
            message=mesage.format(name=get_name(item)),
            wait_time=WHATSAPP_WAIT_TIME,
            tab_close=True,
        )
        cells_list.append(
            Cell(row=index + 1, col=get_col_index(msg_status_col) + 1, value="TRUE")
        )
        cells_list.append(
            Cell(row=index + 1, col=get_col_index(date_col) + 1, value=formatted_date)
        )
    except Exception as e:
        print(f"failed to send message to {get_mobile(item)} due to {str(e)}")


def send_wa_msg_and_update_each_day(
    worksheet: Worksheet, data, msg_col, msg_status_col, date_col, message
):
    headers = data[0]
    cells_list = []
    # data = worksheet.get_all_values()
    for index, item in enumerate(data):
        if is_eligible_for_msg(item, msg_col, msg_status_col):
            if TEST != 0:
                print("Script is in test mode")
                print(
                    f"Starting candidate of {headers[get_col_index(msg_col)]}",
                    get_name(item),
                    get_mobile(item),
                )
                return
            print(f"Sending Message to {item[get_mobile_index()]}")
            sendmsg_and_update_list(
                item, index, cells_list, message, msg_status_col, date_col
            )
    if len(cells_list):
        print(
            f"{headers[get_col_index(msg_col)]} Sent to {len(cells_list)//2} candidates updating the sheet...."
        )
        worksheet.update_cells(
            cells_list, value_input_option=ValueInputOption.user_entered
        )
        print(f"Updated the sheet for {headers[get_col_index(msg_col)]}")
    else:
        print("Nothing to update")


def send_msg_and_update_daywise():
    sheet_operations = sheetOpsWithDayData(REG_SHEET_URL, DAYWISE_WORKSHEET_NAME)
    worksheet = sheet_operations.get_worksheet(REG_SHEET_URL, REG_WORKSHEET_NAME)
    data = sheet_operations.get_data_by_worksheet(worksheet)
    if len(data) <= 1:
        print("Sheet is empty or Some error occured")
        return
    daywise_data = sheet_operations.get_daywise_details()[1:]
    for day, msg_col, msg_status_col, date_col, message, enable in daywise_data:
        if enable == "TRUE":
            print_colored(f"Started processing for day {day}", "green")
            msg = coguide_message.format(message=message)
            # print(day, msg_col, msg_status_col, date_col, msg, enable)
            send_wa_msg_and_update_each_day(
                worksheet, data, msg_col, msg_status_col, date_col, msg
            )
        else:
            print_colored(f"Disabled for day {day}", "red")


if __name__ == "__main__":
    send_msg_and_update_daywise()
