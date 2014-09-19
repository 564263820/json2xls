#!/usr/bin/env python
#-*- coding:utf-8 -*-
import json
from xlwt import Workbook, XFStyle, Style, Font, Pattern


class Json2Xls(object):

    def __init__(self, json_data, sheet_name='sheet0', title_color='periwinkle', font_name='Verdana'):
        self.sheet_name = sheet_name
        self.json_data = json_data
        self.book = Workbook(encoding='utf-8')
        self.sheet = self.book.add_sheet(self.sheet_name)
        self.title_color = title_color

        self.font = Font()
        self.font.name = font_name
        self.font.bold = True
        self.pattern = Pattern()
        self.pattern.pattern = Pattern.SOLID_PATTERN
        self.pattern.pattern_fore_colour = Style.colour_map[self.title_color]
        self.title_style = XFStyle()
        self.title_style.font = self.font
        self.title_style.pattern = self.pattern

    def __genarate_title(self, data):
        start_col = 0
        row = 0
        for key_name in data.keys():
            end_col = len(data[key_name])
            self.sheet.write_merge(row, row, start_col, start_col + end_col - 1, key_name, self.title_style)
            start_col += end_col
            if isinstance(data[key_name], dict):
                row += 1
                self.__genarate_title(data[key_name])

    def json2xls(self):
        data = json.loads(self.json_data)
        if not isinstance(data, (dict, list)):
            raise Exception('bad json format')
        if isinstance(data, dict):
            data = [data]



        self.__genarate_title(data[0])
        self.book.save("test.xls")

json_data = '{"title": {"tag": "im title tag", "ner": "im title ner"}, "body":{"tag": "im body tag", "ner": "im body ner"}}'
j = Json2Xls(json_data)
j.json2xls()
