import re
import math
from decimal import Decimal, ROUND_HALF_UP

def read_dxf_data(file_path):
    text_entities = []
    try:
        with open(file_path, 'r', encoding='utf-8', errors='replace') as f:
            lines = [line.strip() for line in f]
    except Exception as e:
        print(f"ファイルの読み込みに失敗しました: {e}")
        return []

    i = 0
    while i < len(lines):
        if lines[i] == '0' and i + 1 < len(lines) and lines[i+1] in ('TEXT', 'MTEXT'):
            text_val, x, y = "", None, None
            i += 2
            while i < len(lines) and lines[i] != '0':
                code = lines[i]
                val = lines[i+1] if i + 1 < len(lines) else ""
                
                if code in ('1', '3'): text_val += val
                elif code == '10': 
                    try: x = float(val)
                    except ValueError: pass
                elif code == '20': 
                    try: y = float(val)
                    except ValueError: pass
                i += 2
            
            clean_text = text_val.strip().replace('\n', '').replace('\r', '')
            if clean_text:
                text_entities.append({'text': clean_text, 'x': x, 'y': y})
        else:
            i += 1
            
    return text_entities

def parse_target_value(text):
    text = text.strip()
    if '→' in text:
        target_part = text.split('→')[-1].strip()
    else:
        target_part = text
    
    match = re.search(r'\d+\.\d+', target_part)
    if match:
        raw_val = match.group()
        dec_val = Decimal(raw_val)
        return float(dec_val.quantize(Decimal('0.01'), rounding=ROUND_HALF_UP))
    return None

def calculate_proximate_texts(text_entities, max_distance, fallback_value):
    pattern = re.compile(r'[東西南北]+\s*[-－]\s*\d+')
    window_door_pattern = re.compile(r'([A-Za-z]+\d+[A-Za-z]*)|(\d+[A-Za-z]+)|(両開き)|(片開き)')
    
    results = []
    for i, target in enumerate(text_entities):
        if pattern.search(target['text']):
            closest_text = None
            min_dist = float('inf')
            for j, candidate in enumerate(text_entities):
                if i == j:
                    continue
                    
                # ★追加：RHを含む文字列はサッシ抽出候補から完全に除外する
                if 'RH' in candidate['text'].upper():
                    continue
                    
                if window_door_pattern.search(candidate['text']):
                    dist = math.hypot(target['x'] - candidate['x'], target['y'] - candidate['y'])
                    if dist < min_dist and dist <= max_distance:
                        min_dist = dist
                        closest_text = candidate['text']
            
            results.append({
                '項目名 (Target)': target['text'],
                '近接する値 (Value)': closest_text if closest_text else fallback_value,
                '距離': round(min_dist, 2) if closest_text else "（手動入力）",
                'X座標': round(target['x'], 2),
                'Y座標': round(target['y'], 2)
            })
    return results

def extract_foundation_data(text_entities, max_distance=1000):
    foundation_data = {
        '玄関土間面積': None,
        '日射当たる周長': None,
        '日射当たらない周長': None,
        '基礎壁面積': {}
    }
    
    directions_priority = ["北東", "南東", "南西", "北西", "北", "東", "南", "西"]
    
    for ent in text_entities:
        text = ent['text'].replace(' ', '').replace(' ', '')
        
        m_hit = re.search(r'日射の当たる基礎壁周長[:：]?([0-9.]+)', text)
        if m_hit:
            foundation_data['日射当たる周長'] = float(m_hit.group(1))
            continue
            
        m_nohit = re.search(r'日射の当たらない基礎壁周長[:：]?([0-9.]+)', text)
        if m_nohit:
            foundation_data['日射当たらない周長'] = float(m_nohit.group(1))
            continue
            
        for d in directions_priority:
            if text.startswith(d):
                m_wall = re.search(r'[=＝]\s*([0-9.]+)', text)
                if m_wall:
                    foundation_data['基礎壁面積'][d] = float(m_wall.group(1))
                break

    g_text = next((e for e in text_entities if '玄関土間面積表' in e['text']), None)
    if g_text:
        TABLE_WIDTH_TOLERANCE = 3000  
        TABLE_HEIGHT_TOLERANCE = 8000 
        candidates = []
        for e in text_entities:
            if '合計' in e['text']:
                dx = abs(e['x'] - g_text['x'])
                dy = g_text['y'] - e['y'] 
                if 0 < dy < TABLE_HEIGHT_TOLERANCE and dx < TABLE_WIDTH_TOLERANCE:
                    dist = math.hypot(dx, dy)
                    candidates.append((dist, e))
        candidates.sort(key=lambda x: x[0])
        
        if candidates:
            target_gokei = candidates[0][1]
            right_nums = []
            for e in text_entities:
                if e['x'] > target_gokei['x'] and abs(e['y'] - target_gokei['y']) <= 200:
                    if abs(e['x'] - target_gokei['x']) < 4000:
                        val = parse_target_value(e['text'])
                        if val is not None:
                            right_nums.append((e['x'] - target_gokei['x'], val))
            
            right_nums.sort(key=lambda x: x[0])
            if right_nums:
                foundation_data['玄関土間面積'] = right_nums[0][1]
                
    return foundation_data

def extract_roof_floor_data(text_entities):
    roof_floor_data = {'UB部分': None, 'その他床': None, '天井': None, '屋根': None, '外気床': None}
    target_tables = {'UB部分': 'UB部分', 'その他床': 'その他床', '天井': '天井', '屋根': '屋根', '外気床': '外気に接する床'}
    
    TABLE_WIDTH_TOLERANCE = 3000  
    TABLE_HEIGHT_TOLERANCE = 8000 
    NUM_X_TOLERANCE = 4000  
    NUM_Y_TOLERANCE = 200  

    for key, search_word in target_tables.items():
        t_texts = [e for e in text_entities if search_word in e['text']]
        if t_texts:
            t_text = next((e for e in t_texts if '面積表' in e['text']), t_texts[0])
            candidates = []
            for e in text_entities:
                if '合計' in e['text']:
                    dx = abs(e['x'] - t_text['x'])
                    dy = t_text['y'] - e['y'] 
                    if 0 < dy < TABLE_HEIGHT_TOLERANCE and dx < TABLE_WIDTH_TOLERANCE:
                        dist = math.hypot(dx, dy)
                        candidates.append((dist, e))
            candidates.sort(key=lambda x: x[0])
            
            if candidates:
                target_gokei = candidates[0][1]
                right_nums = []
                for e in text_entities:
                    if e['x'] > target_gokei['x'] and abs(e['y'] - target_gokei['y']) <= NUM_Y_TOLERANCE:
                        if abs(e['x'] - target_gokei['x']) < NUM_X_TOLERANCE:
                            val = parse_target_value(e['text'])
                            if val is not None:
                                dist = math.hypot(e['x'] - target_gokei['x'], e['y'] - target_gokei['y'])
                                right_nums.append((dist, val))
                
                right_nums.sort(key=lambda x: x[0])
                if right_nums:
                    roof_floor_data[key] = right_nums[0][1]
                    
    return roof_floor_data

def extract_wall_area_data(text_entities):
    directions = ["北東", "南東", "南西", "北西", "北", "東", "南", "西"]
    area_data = {}
    y_tolerance = 200.0 
    processed_texts = set()
    
    for d in directions:
        search_word = f"{d}壁面面積"
        target_text_info = None
        
        for item in text_entities:
            item_id = f"{item['text']}_{item['x']}_{item['y']}"
            if item_id not in processed_texts and search_word in item['text'].replace(' ', ''):
                target_text_info = item
                processed_texts.add(item_id)
                break
        
        if target_text_info:
            target_y = target_text_info['y']
            target_x = target_text_info['x']
            same_row_texts = [
                item for item in text_entities 
                if abs(item['y'] - target_y) <= y_tolerance and item['x'] > target_x
            ]
            same_row_texts = sorted(same_row_texts, key=lambda i: i['x'])
            
            for item in same_row_texts:
                val = parse_target_value(item['text'])
                if val is not None:
                    area_data[d] = val
                    break
                    
    return area_data

def save_to_excel(results, file_path):
    if results:
        import openpyxl
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "抽出結果"
        headers = list(results[0].keys())
        ws.append(headers)
        for row_data in results:
            ws.append([row_data[k] for k in headers])
        wb.save(file_path)
        print(f"\n✅ 窓・ドアの抽出データ(確認用)を {file_path} に保存しました。")
    else:
        print("\n⚠️ 「方角-数字」のパターンに一致するテキストが見つかりませんでした。")