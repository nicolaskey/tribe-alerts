import gspread
import pprint as pp
from oauth2client.service_account import ServiceAccountCredentials
from person import Person

class Workbook(object):

    # Column A corresponds to number 1. Index 0 has no meaning.
    num_to_letter = [0,'A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y', 'Z']

    def __init__(self, workbook_name):
        self.scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        self.credentials = ServiceAccountCredentials.from_json_keyfile_name('TribeAlerts-76b6655ce5f4.json', self.scope)
        self.gc = gspread.authorize(self.credentials)
        self.workbook_name = workbook_name
        self.wb = self.gc.open(workbook_name)

    def _reautorize(self):
        # TODO: make a decorator that will do this
        self.gc = gspread.authorize(self.credentials)

    def get_sheet(self, sheet):
        ''' Get specified sheet
            in:  sheet (str) - the wanted sheet name
            out: worksheet object - if found
                 None - no worksheet by sheet name
        '''
        if sheet in self.get_all_worksheets():
            return self.wb.worksheet(sheet)
        else:
            return None

    def get_all_worksheets(self):
        # return a list of all worksheet titles

        worksheets = self.wb.worksheets()
        return [sheet._properties['title'] for sheet in worksheets]

    def _scrub_phone(self, ph_num):
        return ph_num.replace('(', '').replace(')', '').replace(' ', '').replace('-', '')

    def get_name_phone_dues(self):
        """ Find all names, phone numbers, and current amount due
            in:  none
            out: List of Person objects
        """
        dues_members_sheet = self.get_sheet('Dues - Members')
        if not dues_members_sheet:
            return {}

        title_cell = dues_members_sheet.find('Tribe Member')
        phone_cell = dues_members_sheet.find('Phone')
        total_cell = dues_members_sheet.find('Total')
        dues_cell = dues_members_sheet.find('Amount Due')

        first_row = title_cell._row + 2
        last_row = total_cell._row - 1
        name_col = title_cell._col
        ph_col = phone_cell._col
        dues_col = dues_cell._col

        first_cell_names_str = self.num_to_letter[name_col] + str(first_row)
        last_cell_names_str = self.num_to_letter[ph_col] + str(last_row)

        first_cell_dues_str = self.num_to_letter[dues_col] + str(first_row)
        last_cell_dues_str = self.num_to_letter[dues_col] + str(last_row)

        data1 = dues_members_sheet.range(first_cell_names_str + ':' + last_cell_names_str)
        data2 = dues_members_sheet.range(first_cell_dues_str + ':' + last_cell_dues_str)

        people = []
        num_rows = last_row - first_row + 1
        for row in range(0, num_rows):
            name = data1[2*row].value
            ph_num = self._scrub_phone(data1[(2*row) + 1].value)
            dues = data2[row].value
            people.append(Person(name, ph_num, dues))

        return people

    def add_phone(self, person: Person):
        # enroll a user for alerts by adding their phone number
        # to the spreadsheet
        pass
