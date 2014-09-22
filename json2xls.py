#!/usr/bin/env python
#-*- coding:utf-8 -*-
import json
import requests
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

    def __init__(self, url_or_json, method='get', params=None, data=None, headers=None, sheet_name='sheet0', title_color='lime', font_name='Arial'):
        self.sheet_name = sheet_name
        self.url_or_json = url_or_json
        self.method = method
        self.params = params
        self.data = data
        self.headers = headers

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

        self.borders = Borders()
        self.borders.left = 1
        self.borders.right = 1
        self.borders.top = 1
        self.borders.bottom = 1

        self.title_style = XFStyle()
        self.title_style.font = self.font
        self.title_style.borders = self.borders
        self.title_style.pattern = self.pattern

    def __parse_dict_depth(self, d, depth=0):
        if not isinstance(d, dict) or not d:
            return depth
        return max(self.__parse_dict_depth(v, depth+1) for k, v in d.iteritems())

    def __check_dict_deep(self, d):
        depth = self.__parse_dict_depth(d)
        if depth > 1:
            raise Exception("dict is too deep")

    def __get_json(self):
        data = None
        try:
            data = json.loads(self.url_or_json)
        except:
            try:
                if self.method.lower() == 'get':
                    resp = requests.get(self.url_or_json, params=self.params, headers=self.headers)
                    data = resp.json()
                else:
                    resp = requests.post(self.url_or_json, data=self.data, headers=self.headers)
                    data = resp.json()
            except Exception as e:
                print e
        return data

    def __fill_title(self, data):
        self.__check_dict_deep(data)
        for index, key in enumerate(data.keys()):
            self.sheet.row(self.title_start_row).write(index, key, self.title_style)
        self.title_start_row += 1

    def __fill_data(self, data):
        self.__check_dict_deep(data)
        for index, value in enumerate(data.values()):
            self.sheet.row(self.title_start_row).write(index, value)

        self.title_start_row += 1

    def make(self):
        data = self.__get_json()
        if not isinstance(data, (dict, list)):
            raise Exception('bad json format')
        if isinstance(data, dict):
            data = [data]

        self.__fill_title(data[0])
        for d in data:
            self.__fill_data(d)
        self.book.save("test.xls")

url_or_json = '''[{"name": "John", "age": 30,
  "sex": "male"},
 {"name": "Alice", "age": 18,
  "sex": "female"}
]'''
j = Json2Xls("http://142.4.209.29:9080/ask/%E6%88%91%E6%83%B3%E8%87%AA%E6%9D%80")
j.make()

