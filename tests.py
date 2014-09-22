
data = {
    "x": "vx",
    "a": {"a1": "va1"},
    "b": {"b1": "vb1", "b2": "vb2"},
    "c": {"c1": "vc1", "c2": "vc2", "c3": "vc3"},
}

d = {}

def parse_dict(data, parent=None):
    for key, value in data.items():
        if isinstance(value, dict):
            value_len = len(value)
            if value_len > len(data[key]):
                d[key]['len'] = value_len
            else:
                d.update({key: len(value)})
                parse_dict(value, parent=key)
        else:
            if not parent:
                d.update({key: 1})

def v(data):
    print data.items()

v(data)
