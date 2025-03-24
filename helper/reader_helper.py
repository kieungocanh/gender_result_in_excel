import json

def load_json(file_path):
    with open(file_path, encoding='utf-8') as f:
        data = json.load(f)
    return data
