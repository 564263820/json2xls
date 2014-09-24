json2xls:根据json数据生成excel表格
==================================

[![](https://badge.fury.io/py/json2xls.png)](http://badge.fury.io/py/json2xls)
[![](https://pypip.in/d/json2xls/badge.png)](https://pypi.python.org/pypi/json2xls)


**安装**

    pip install json2xls

**根据json数据生成excel**

code:

    :::python
    from json2xls import Json2Xls

    json_data = '{"name": "ashin", "age": 16, "sex": "male"}'
    Json2Xls('test.xls', json_data).make()

command:

    python json2xls.py test.xls '{"a":"a", "b":"b"}'
    python json2xls.py test.xls '[{"a":"a", "b":"b"},{"a":1, "b":2}]'

    # from file: json of text
    python json2xls.py test.xls "`cat data.json`"

    # from file: json of line
    python json2xls.py test.xls data2.json

excel:

    age | name | sex
    ----|------|----
    30  | John | male
    18  | Alice| female


**根据请求url返回的json生成excel**

默认请求为get，get请求参数为params={}, post请求参数为data={}

code:

    :::python
    from json2xls import Json2Xls

    url = 'http://api.bosonnlp.com/sentiment/analysis'
    Json2Xls('test.xlsx', url, method='post').make()

command:

    python json2xls.py test.xls http://api.map.baidu.com/telematics/v3/weather\?location\=%E4%B8%8A%E6%B5%B7\&output\=json\&ak\=640f3985a6437dad8135dae98d775a09

excel:

    status | message
    -------|--------
    403    | no token header

**自定义title和body的生成**

默认只支持一层json的excel生成，如果是多层套嵌的json，请自定义生成title和body，只需定义`title_callback`和`body_callback`方法，在调用`make`的时候传入即可。

    :::python
    def title_callback(obj, data):
        '''use one data record to generate excel title'''
        for index, key in enumerate(data.keys()):
            obj.sheet.row(obj.title_start_row).write(index,
                                                     'new:' + key.encode('utf-8'), obj.title_style)
        obj.title_start_row += 1

    j = Json2Xls('../test_data/title_callback.xls', "../test_data/sentiment_output.json")
    j.make(title_callback=title_callback)

