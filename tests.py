from treelib import Tree, Node


data = {
    "a": {"a1": "va1"},
    "b": {"b1": "vb1", "b2": "vb2"},
    "c": {"c1": "vc1", "c2": "vc2", 'c3': 'vc3'},
    "d": {"d1": {'dd1': 'vdd1', 'dd2': 'vdd2', 'dd3': 'vdd3'}, 'd2': 'vd2'},
}

d = {}

def parse_dict(data, parent=None):
    for key, value in data.items():
        if isinstance(value, dict):
            # {'a': {'parent': None, 'len': 1}, 'c': {'parent': None, 'len': 3}, 'b': {'parent': None, 'len': 2}, 'd': {'parent': None, 'len': 2}, 'd1': {'parent': 'd', 'len': 3}}
            value_len = len(value)
            print value_len
            if parent and value_len > d[parent]['len']:
                d[parent]['len'] = value_len
            else:
                d.update({key: {'len': len(value), 'parent':parent}})
                parse_dict(value, parent=key)

parse_dict(data)
print d
