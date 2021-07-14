import time
from tqdm import tqdm

import gspread
from google.oauth2.service_account import Credentials


def main():
    g_sheet_name = "Your Google Sheet Here"
    sheet = connect_google(g_sheet_name)

    reference = "File/path/here"
    check_sheet(reference, sheet)


def connect_google(sheet_name):
    """
    This connects to the chosen sheet, with the hardcoded details
    and credentials below.

    :return: The chosen sheet
    """
    # Sets the scope of API access
    scope = ['https://www.googleapis.com/auth/spreadsheets',
             'https://www.googleapis.com/auth/drive']

    # Gets the key for the service account from the JSON file
    credentials = Credentials.from_service_account_file("yourJSONfilehere", scopes=scope)

    # Uses the key to authorise access
    client = gspread.authorize(credentials)

    # Gets the sheet (Note you need to grant access to the service account email)
    google_sh = client.open(sheet_name)

    return google_sh.get_worksheet(5)  # Returns the sheet


def check_sheet(reference, sheet) -> None:
    """
    The below logic is what loops through each row, and then checks a specific cell
    in that row.

    You need to update the specific cell you are checking, depending on how the
    database spreadsheet is set up. For example:

    If the data we are looking for is contained in cell 'C', we will check `row[2]`

    :param reference The Excel sheet we will be checking against
    :param sheet: The GSheet to check
    """
    # The start and end values for the sheet rows
    start_val = 0
    end_val = 500

    # Loops over the rows, if the row content is in the opt-out document
    # then it highlights the row Red, and adds "Opt-Out" to cell 'J'
    # NB* This loop cannot handle blank spaces!
    for r in tqdm(range(start_val, end_val)):  # tqdm is a progress bar module
        try:
            row = sheet.row_values(r)
            # print("Row {}: {}".format(r, row[4]))     # For debugging/testing

            # This if statement deals with any numbers which are incomplete (but not empty)
            # Empty number cells are dealt with in the 'Except' block below
            if len(row[4]) < 5:
                # This sets the bg color to orange/brown if incomplete number
                sheet.format("E{}".format(r, r),
                             {"backgroundColor": {
                                 "red": 0.8,
                                 "green": 0.4,
                                 "blue": 0.0,
                                 "alpha": 0.5
                             }
                             })
                continue
            # Checks if the number is too long and flags it
            elif len(row[4]) > 9:
                print("Row {}: Number too long.".format(r))

            # Checks for any spaces in the number formatting
            if ' ' in row[4]:
                print("Row {}: Number not formatted correctly.".format(r))
                continue

            # Checks if the current number matches any numbers on the opt-out list
            for row_ in reference:
                # print(type(row_))
                if row[4] in str(row_):
                    print(
                        "Found 1 opt out: {}".format(row[4]))  # Prints the number (remove from opt_out list after run)
                    sheet.update_cell(r, 11, "Opt-Out")
                    # This sets the bg color to red for all the opt-outs
                    sheet.format("A{}:L{}".format(r, r),
                                 {"backgroundColor": {
                                     "red": 1.0,
                                     "green": 0.0,
                                     "blue": 0.0,
                                     "alpha": 1.0
                                 }
                                 })
        # If the cell contains no number, we highlight the cell and continue
        except IndexError:
            # print("No number for row {}".format(r))   # For debugging
            # This sets the bg color to orange/brown if missing number
            sheet.format("E{}".format(r, r),
                         {"backgroundColor": {
                             "red": 0.8,
                             "green": 0.4,
                             "blue": 0.0,
                             "alpha": 0.5
                         }
                         })
            continue

        time.sleep(1)  # This slows the checking down, to not exceed the API call limits


if __name__ == '__main__':
    main()
