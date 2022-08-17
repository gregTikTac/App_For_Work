import openpyxl

class XmlLoader:
    data = {}
    headers = ['older_og', 'officer_og', 'pat', 'pm_for_group_og', 'ak_for_pat']

    def __init__(self, path):
        self.path = path

    def load(self):
        file_for_work = openpyxl.load_workbook('word_automation.xlsm')
        sheet = file_for_work.active
        index = 0
        for head in self.headers:
            self.data[head] = self._get_column(index, sheet)
            index += 1


    def _get_column(self, index, sheet):
        lst = []
        for row in sheet.rows:
            if row[index].value == None:
                return lst[1:]
            lst.append(row[index].value)
        return lst
