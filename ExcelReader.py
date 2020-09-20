import xlrd


class ExcelReader:
    def __init__(self, excel_path):
        self.excel_path = excel_path

    def read(self):
        collected_data = []

        opened_workbook = xlrd.open_workbook(self.excel_path)
        selected_sheet = opened_workbook.sheet_by_index(0)

        for row in range(1, selected_sheet.nrows):
            selected_data = {
                "targetEmail": selected_sheet.cell_value(row, 0),
                "targetPassword": selected_sheet.cell_value(row, 1),
                "studentEmail": selected_sheet.cell_value(row, 2),
                "from": selected_sheet.cell_value(row, 3),
                "subject": selected_sheet.cell_value(row, 4),
                "excludeDomain": selected_sheet.cell_value(row, 5)
            }
            collected_data.append(selected_data)

        return collected_data
