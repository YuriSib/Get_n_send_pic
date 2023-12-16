import urllib.request
import pickle
import os
import openpyxl


class ExcelWork:
    def __init__(self, table_path):
        self.table_path = table_path
        self.old_link_dict = {}
        self.new_link_dict = {}

    @staticmethod
    def find_col_num(col_name, work_sheet):
        col_num = -1
        for col in work_sheet.iter_cols(min_row=1, max_row=1, values_only=True):
            col_num += 1
            if col_name in col[0]:
                break
        return col_num

    def get_links(self):
        wb = openpyxl.load_workbook(self.table_path)
        ws = wb.active

        link_col_num = self.find_col_num('фото', ws)
        code_col_num = self.find_col_num('Код товара', ws)

        row_num = 1
        for row in ws.iter_rows(min_col=link_col_num + 1, max_col=link_col_num + 1, min_row=2, values_only=True):
            row_num += 1
            if row[0]:
                product_links = [link.strip() for link in row[0].split(',')]
            else:
                product_links = 0

            product_code = ws.cell(row_num, code_col_num + 1).value
            self.old_link_dict[product_code] = product_links
        return self.old_link_dict

    def links_rename(self):
        if self.old_link_dict:
            wb = openpyxl.load_workbook(self.table_path)
            ws = wb.active

            link_col_num = self.find_col_num('фото', ws)
            code_col_num = self.find_col_num('Код товара', ws)

            for key, values in self.old_link_dict.items():
                amount_link = len(values)
                link_list = [f"{key}-{num}" for num in range(1, amount_link + 1)]
                self.new_link_dict[key] = link_list



class PicSave:
    def __init__(self, url_, name):
        self.url_ = url_
        self.name = name

    def photo_saver(self):
        try:
            urllib.request.urlopen(self.url_)
        except urllib.error.URLError as e:
            print("Ошибка при скачивании фото:", e)
        else:
            try:
                with open(f'{self.name}.jpg', 'wb') as f:
                    f.write(urllib.request.urlopen(self.url_).read())
                    print("Фото успешно скачано!")
            except Exception as e:
                print(f"Ошибка при скачивании фото: {e}")


class WorkToPic:
    def __init__(self, server_position, path_table_dir):
        self.server_position = server_position
        self.path_table_dir = path_table_dir
        self.list_tables = []

    def table_choice(self):
        files = os.listdir(self.path_table_dir)
        self.list_tables = [file for file in files if ".xlsx"]

        return self.list_tables

    def save(self):
        for table in self.list_tables:
            ew = ExcelWork(table)
            link_dict = ew.get_links()


if __name__ == "__main__":
    table_path_ = 'Паяльники.xlsx'
    ew_ = ExcelWork(table_path_)
    link_dict_ = ew_.get_links()
    ew_.links_rename()

