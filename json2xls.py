#!/usr/bin/env python
#-*- coding:utf-8 -*-
import json
from xlwt import Workbook, XFStyle, Style, Font, Pattern, Borders

XLS_COLORS = [
    'gray_ega', 'dark_green', 'indigo',
    'gold', 'blue_grey', 'dark_green_ega',
    'lavender', 'yellow', 'purple_ega',
    'olive_ega', 'olive_green', 'light_yellow',
    'pale_blue', 'violet', 'dark_red_ega',
    'cyan_ega', 'dark_blue', 'blue_gray',
    'magenta_ega', 'lime', 'blue', 'grey40',
    'pink', 'grey25', 'rose', 'white', 'black',
    'silver_ega', 'gray50', 'periwinkle',
    'sea_green', 'orange', 'red', 'grey80',
    'dark_teal', 'brown', 'ivory', 'bright_green',
    'ocean_blue', 'dark_blue_ega', 'dark_yellow',
    'light_turquoise', 'light_blue', 'dark_purple',
    'ice_blue', 'light_orange', 'grey50', 'grey_ega',
    'tan', 'sky_blue', 'gray25', 'gray40', 'coral',
    'light_green', 'aqua', 'dark_red', 'gray80',
    'green', 'teal_ega', 'teal', 'plum', 'turquoise'
]


class Json2Xls(object):

    def __init__(self, json_data, sheet_name='sheet0', title_color='lime', font_name='Arial'):
        self.sheet_name = sheet_name
        self.json_data = json_data
        self.book = Workbook(encoding='utf-8')
        self.sheet = self.book.add_sheet(self.sheet_name)
        self.title_color = title_color
        self.title_start_col = 0
        self.title_start_row = 0
        self.title_merge_info = {}
        self.data_start_row = 0

        self.font = Font()
        self.font.name = font_name
        self.font.bold = True

        self.pattern = Pattern()
        self.pattern.pattern = Pattern.SOLID_PATTERN
        self.pattern.pattern_fore_colour = Style.colour_map[self.title_color]

        self.borders = Borders()
        self.borders.left = 1
        self.borders.right = 1
        self.borders.top = 1
        self.borders.bottom = 1

        self.title_style = XFStyle()
        self.title_style.font = self.font
        self.title_style.borders = self.borders
        self.title_style.pattern = self.pattern



    def __parse_merge_info(self, data, parent=None):
        for key, value in data.items():
            if isinstance(value, dict):
                value_len = len(value)
                if value_len > len(data[key]):
                    self.title_merge_info[key]['len'] = value_len
                else:
                    self.title_merge_info.update({key: len(value)})
                    self.__parse_merge_info(value, parent=key)
            else:
                if not parent:
                    self.title_merge_info.update({key: 1})

    def __genarate_title(self, data):
        self.__parse_merge_info(data)
        for key_col, key_name in enumerate(data.keys()):
            self.sheet.write_merge(
                self.title_start_row,
                self.title_start_row,
                self.title_start_col,
                self.title_start_col + self.title_merge_info[key_name] - 1,
                key_name,
                self.title_style
            )
            self.title_start_col += self.title_merge_info[key_name]

    def __fill_data(self, data):
        for index, key_name in enumerate(self.title_merge_info.keys()):
            if isinstance(data[key_name], dict):
                for index1, value in enumerate(data[key_name].values()):
                    self.sheet.row(self.data_start_row).write(index1 + index, value)
            else:
                self.sheet.row(self.data_start_row).write(index, data[key_name])
        self.data_start_row += 1

    def make(self):
        data = json.loads(self.json_data)
        if not isinstance(data, (dict, list)):
            raise Exception('bad json format')
        if isinstance(data, dict):
            data = [data]

        self.__genarate_title(data[0])
        self.data_start_row = len(self.title_merge_info) - 1
        for d in data:
            self.__fill_data(d)
        self.book.save("test.xls")

json_data = '[{"title": {"tag": "im xtitle tag", "ner": "im xtitle ner"}, "body":{"tag": "im xbody tag"}}, {"title": {"tag": "im title tag", "ner": "im title ner"}, "body":{"tag": "im body tag"}}]'
j = Json2Xls(json_data)
j.make()

