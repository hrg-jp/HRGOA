import json
def json2dict(json_file):
    with open(json_file) as json_data:
        content = json.load(json_data)
        json_data.close()
    return content
