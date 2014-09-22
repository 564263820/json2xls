Json2Xls:根据json数据生成excel表格
==================================

Json2Xls不支持多层套嵌的json数据，只可以根据一层json生成表格

* 根据json数据生成excel

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


* 根据请求url返回的json生成excel

code:

    :::python
    from json2xls import Json2Xls

    url = 'http://api.bosonnlp.com/sentiment/analysis'
    Json2Xls('test.xlsx', url, method='post')

excel:

    status | message
    -------|--------
    403    | no token header

