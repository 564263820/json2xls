Json2Xls:根据json数据生成excel表格
==================================

Json2Xls不支持多层套嵌的json数据，只可以根据一层json生成表格

**根据json数据生成excel**

code:

    :::python
    from json2xls import Json2Xls

    json_data = '{"name": "ashin", "age": 16, "sex": "male"}'
    Json2Xls('test.xls', json_data)

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
    Json2Xls('test.xlsx', url, method='post')

excel:

    status | message
    -------|--------
    403    | no token header


command:

    python json2xls.py test.xls http://api.map.baidu.com/telematics/v3/weather\?location\=%E4%B8%8A%E6%B5%B7\&output\=json\&ak\=640f3985a6437dad8135dae98d775a09
    python json2xls.py test3.xls '{"a":"a", "b":"b"}'
