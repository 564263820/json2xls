#!/usr/bin/env python
# coding:utf-8

"""
json2xls
===========

根据json数据生成excel表格

安装
----

:py:mod:`json2xls` 代码托管在 `GitHub`_，并且已经发布到 `PyPI`_，可以直接通过 `pip` 安装::

    $ pip install json2xls

:py:mod:`json2xls` 以 MIT 协议发布。

.. _GitHub: https://github.com/axiaoxin/json2xls
.. _PyPI: https://pypi.python.org/pypi/json2xls

使用教程
--------

    >>> from json2xls import Json2Xls

    >>> json_data = '{"name": "ashin", "age": 16, "sex": "male"}'
    >>> Json2Xls('test.xls', json_data).make()

可以在 `json2xls`_ 文档站点阅读详细的 json2xls 文档。

.. _json2xls: http://json2xls.axiaoxin.com/
"""

__author__ = 'Axiaoxin'
__email__ = '254606826@qq.com'
__version__ = '0.1.1'

from json2xls import Json2Xls
