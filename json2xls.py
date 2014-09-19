#!/usr/bin/env python
#-*- coding:utf-8 -*-
import json
from xlwt import Workbook, XFStyle, Style, Font, Pattern

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

        self.font = Font()
        self.font.name = font_name
        self.font.bold = True
        self.pattern = Pattern()
        self.pattern.pattern = Pattern.SOLID_PATTERN
        self.pattern.pattern_fore_colour = Style.colour_map[self.title_color]
        self.title_style = XFStyle()
        self.title_style.font = self.font
        self.title_style.pattern = self.pattern



    def __parse_merge_info(self, data, parent=None):
        for key, value in data.items():
            if isinstance(value, dict):
                value_len = len(value)
                if parent and value_len > self.title_merge_info[parent]['len']:
                    self.title_merge_info[parent]['len'] = value_len
                else:
                    self.title_merge_info.update({key: {'len': len(value), 'parent':parent}})
                    self.__parse_merge_info(value, parent=key)

    def __genarate_title(self, data):
        for key_col, key_name in enumerate(data.keys()):
            self.sheet.write_merge(
                self.title_start_row,
                self.title_start_row,
                self.title_start_col,
                self.title_start_col + self.title_merge_len - 1,
                key_name,
                self.title_style
            )
            self.title_start_col += self.title_merge_len

    def json2xls(self):
        data = json.loads(self.json_data)
        if not isinstance(data, (dict, list)):
            raise Exception('bad json format')
        if isinstance(data, dict):
            data = [data]

        self.__genarate_title(data[0])
        self.book.save("test.xls")

#json_data = '{"title": {"tag": "im title tag", "ner": "im title ner"}, "body":{"tag": "im body tag", "ner": "im body ner"}}'
#j = Json2Xls(json_data)
#j.json2xls()

