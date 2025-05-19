import os
import json

def flatten_items(items):
    """nested-category 구조를 flat + __category__ 마커 방식으로 변환"""
    flat = []
    if items and isinstance(items[0], dict) and 'category' in items[0]:
        for cat in items:
            cat_name = cat.get('category', '')
            flat.append({'__category__': cat_name})
            for item in cat.get('items', []):
                flat.append(item)
    else:
        flat = items
    return flat

def convert_folder_jsons(folder_path, backup=True):
    for fname in os.listdir(folder_path):
        if not fname.endswith('.json'):
            continue
        fpath = os.path.join(folder_path, fname)
        with open(fpath, 'r', encoding='utf-8') as f:
            data = json.load(f)
        items = data.get('items', [])
        flat_items = flatten_items(items)
        data['items'] = flat_items
        # 백업본 저장
        if backup:
            os.rename(fpath, fpath + '.bak')
        # 덮어쓰기
        with open(fpath, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        print(f'변환 완료: {fname}')

if __name__ == '__main__':
    folder = '2025년 견적서_주식회사/Json_1'  # 변환할 폴더 경로
    convert_folder_jsons(folder) 