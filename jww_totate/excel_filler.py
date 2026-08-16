import os
import glob
import re
import unicodedata
import xlwings as xw
from decimal import Decimal, ROUND_HALF_UP

def get_dynamic_u_values():
    """ U値シートから動的に「F-1」等に紐づく平均熱貫流率を抽出する """
    u_val_files = [f for f in glob.glob('*.xlsx') if 'U値' in f]
    if not u_val_files:
        print("⚠️ U値が記載されたExcelファイルが見つかりません。デフォルト値を使用します。")
        return {}
    
    import openpyxl
    u_values = {}
    try:
        wb = openpyxl.load_workbook(u_val_files[0], data_only=True)
        target_sheet = "木造用" if "木造用" in wb.sheetnames else wb.sheetnames[0]
        ws = wb[target_sheet]
        
        pattern = re.compile(r'^[A-Z]-\d+$')
        
        for r in range(1, ws.max_row + 1):
            for c in range(1, ws.max_column + 1):
                val = ws.cell(r, c).value
                if isinstance(val, str):
                    val_s = unicodedata.normalize('NFKC', val).strip().upper()
                    if pattern.match(val_s):
                        target_r = r + 14
                        if target_r <= ws.max_row:
                            label_val = ws.cell(target_r, c).value
                            if isinstance(label_val, str) and "平均熱貫流率" in label_val.replace(" ", "").replace(" ", ""):
                                for offset in range(1, 6):
                                    u_val_candidate = ws.cell(target_r, c + offset).value
                                    if isinstance(u_val_candidate, (int, float)):
                                        u_values[val_s] = float(u_val_candidate)
                                        break
                                    elif isinstance(u_val_candidate, str):
                                        try:
                                            u_values[val_s] = float(u_val_candidate.strip())
                                            break
                                        except ValueError:
                                            pass
        wb.close()
    except Exception as e:
        print(f"⚠️ U値シートの読み込みエラー: {e}")
        
    return u_values

def parse_window_string(val_str, sash_u=2.33, shutter_u=2.11):
    u_value = sash_u  # 基本は入力されたサッシU値を使用
    eta_value = 0.32
    width = 0.0
    height = 0.0
    match = re.search(r'([A-Za-z]+)(\d{5,6})[A-Za-z]*', val_str)
    
    if match:
        prefix = match.group(1).upper()
        numbers = match.group(2)
        if 'HS' in prefix:
            u_value = shutter_u # HSが入っていればシャッター用U値
            
        w_str = numbers[:3]
        h_str = numbers[3:]
        w_raw = float(w_str[0] + '.' + w_str[1:])
        h_raw = float(h_str[0] + '.' + h_str[1:])
        width = float(Decimal(str(w_raw)).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP))
        height = float(Decimal(str(h_raw)).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP))
        
    return u_value, eta_value, width, height

def transfer_to_foundation_sheet(wb, foundation_data, dynamic_u_values):
    """Ｃ（基礎）シートへの転記メソッド"""
    sheet_name = 'Ｃ（基礎）'
    sheet_names = [sheet.name for sheet in wb.sheets]
    if sheet_name not in sheet_names: return

    ws = wb.sheets[sheet_name]
    
    # 面積の転記
    if foundation_data.get('玄関土間面積') is not None: 
        ws.range('H7').value = foundation_data['玄関土間面積']
    if foundation_data.get('基礎断熱部分面積') is not None: 
        ws.range('H8').value = foundation_data['基礎断熱部分面積']

    # 玄関土間の周長
    if foundation_data.get('日射当たる周長（玄関土間）') is not None:
        ws.range('H22').value = foundation_data['日射当たる周長（玄関土間）']
        ws.range('K22').value = 0.99
    if foundation_data.get('日射当たらない周長（玄関土間）') is not None:
        ws.range('H23').value = foundation_data['日射当たらない周長（玄関土間）']
        ws.range('K23').value = 0.99
        ws.range('AG23').value = True 

    # 浴室土間（他土間）の周長
    if foundation_data.get('日射当たる周長（浴室土間）') is not None:
        ws.range('H24').value = foundation_data['日射当たる周長（浴室土間）']
        ws.range('K24').value = 0.99
    if foundation_data.get('日射当たらない周長（浴室土間）') is not None:
        ws.range('H25').value = foundation_data['日射当たらない周長（浴室土間）']
        ws.range('K25').value = 0.99
        ws.range('AG25').value = True

    opposites = {'北': '南', '南': '北', '東': '西', '西': '東', '北東': '南西', '南西': '北東', '北西': '南東', '南東': '北西'}
    wall_data = foundation_data.get('基礎壁面積', {})
    processed_walls = {}
    for dir_name, area in wall_data.items():
        processed_walls[dir_name] = {'area': area, 'check': True} 
        opp = opposites.get(dir_name)
        if opp and opp not in wall_data: processed_walls[opp] = {'area': area, 'check': None}

    has_8dir = any(d in processed_walls for d in ['北東', '南東', '南西', '北西'])
    for row in range(35, 45):
        cell_dir = ws.range(f'E{row}').value 
        if has_8dir and cell_dir in ['東', '南', '西', '北']:
            if cell_dir == '東': cell_dir = '北東'
            elif cell_dir == '南': cell_dir = '南東'
            elif cell_dir == '西': cell_dir = '南西'
            elif cell_dir == '北': cell_dir = '北西'
            ws.range(f'E{row}').value = cell_dir
            
        if cell_dir in processed_walls:
            info = processed_walls[cell_dir]
            ws.range(f'B{row}').value = 'L-1'         
            ws.range(f'H{row}').value = info['area']
            ws.range(f'K{row}').value = dynamic_u_values.get('L-1', 1.23)
            ws.range(f'AG{row}').value = True if info.get('check') is True else None

def transfer_to_roof_floor_sheet(wb, roof_floor_data, dynamic_u_values):
    sheet_name = 'Ｂ（屋根・床等）'
    if sheet_name not in [s.name for s in wb.sheets]: return

    ws = wb.sheets[sheet_name]
    
    configs = [
        {'key': '吹抜', 'shiyo': 'R-1', 'bui': '天井', 'u_val': dynamic_u_values.get('R-1', 1.23), 'temp_coef': 1.0},
        {'key': 'その他床', 'shiyo': 'F-1', 'bui': 'その他床', 'u_val': dynamic_u_values.get('F-1', 0.413), 'temp_coef': 0.7},
        {'key': '天井', 'shiyo': 'R-1', 'bui': '天井', 'u_val': dynamic_u_values.get('R-1', 0.288), 'temp_coef': 1.0},
        {'key': '屋根', 'shiyo': 'R-2', 'bui': '屋根', 'u_val': dynamic_u_values.get('R-2', 0.207), 'temp_coef': 1.0},
        {'key': '外気', 'shiyo': 'F-2', 'bui': '外気床', 'u_val': dynamic_u_values.get('F-2', 0.337), 'temp_coef': 1.0}
    ]
    
    row = 19
    for config in configs:
        area_val = roof_floor_data.get(config['key'])
        if area_val is not None and area_val > 0:
            ws.range(f'B{row}').value = config['shiyo']       
            ws.range(f'D{row}').value = config['bui']         
            ws.range(f'F{row}').value = area_val              
            ws.range(f'L{row}').value = config['u_val']       
            ws.range(f'N{row}').value = config['temp_coef']   
            row += 1

def transfer_to_wall_area(wb, wall_area_data):
    sheet_mapping = {
        '北': 'Ａ（北）', '北東': 'Ａ（北東）', '東': 'Ａ（東）', '南東': 'Ａ（南東）',
        '南': 'Ａ（南）', '南西': 'Ａ（南西）', '西': 'Ａ（西）', '北西': 'Ａ（北西）'
    }
    sheet_names = [s.name for s in wb.sheets]
    for direction, area_val in wall_area_data.items():
        sheet_name = sheet_mapping.get(direction)
        if sheet_name in sheet_names:
            ws = wb.sheets[sheet_name]
            ws.range('L35').value = area_val

def transfer_to_envelope_calc_sheet(extracted_data, foundation_data, roof_floor_data, wall_area_data, house_name, house_address, output_dir=".", sash_u=2.33, shutter_u=2.11):
    dynamic_u_values = get_dynamic_u_values()
    
    target_files = [f for f in glob.glob('*外皮計算書*.xlsx') if not os.path.basename(f).startswith('~$')]
    if not target_files:
        print("⚠️ エラー: 同フォルダ内に「外皮計算書」を含むExcelファイルが見つかりません。")
        return None
        
    template_path = os.path.abspath(target_files[0])
    safe_house_name = "".join(c for c in house_name if c not in r'\/:*?"<>|') 
    output_filename = os.path.abspath(os.path.join(output_dir, f"外皮計算書ver2.4（{safe_house_name}）.xlsx"))
    
    app = xw.App(visible=False)
    
    try:
        wb = app.books.open(template_path)
        
        sheet_common = '共通条件・結果'
        if sheet_common in [s.name for s in wb.sheets]:
            ws_common = wb.sheets[sheet_common]
            ws_common.range('K6').value = house_name
            ws_common.range('K7').value = house_address
            ws_common.range('AA7').value = "５地域" if "東広島市" in house_address else "６地域"

        sheet_mapping = {
            '北': 'Ａ（北）', '北東': 'Ａ（北東）', '東': 'Ａ（東）', '南東': 'Ａ（南東）',
            '南': 'Ａ（南）', '南西': 'Ａ（南西）', '西': 'Ａ（西）', '北西': 'Ａ（北西）'
        }
        
        grouped_data = {}
        for item in extracted_data:
            target_name = item['項目名 (Target)']
            val_str = str(item['近接する値 (Value)'])
            direction = re.split(r'[-－]', target_name)[0].strip()
            num = re.split(r'[-－]', target_name)[1].strip()
            if direction not in grouped_data: grouped_data[direction] = []
            grouped_data[direction].append({'num': int(num), 'val': val_str, 'raw_name': target_name})
        
        sheet_names = [s.name for s in wb.sheets]
        
        for direction, items in grouped_data.items():
            sheet_name = sheet_mapping.get(direction)
            if sheet_name not in sheet_names: continue
                
            ws = wb.sheets[sheet_name]
            
            for row in range(8, 20):
                for col in ['B', 'D', 'F', 'H', 'J', 'AG']: ws.range(f'{col}{row}').value = None 
            for row in range(26, 29):
                for col in ['J', 'N', 'P', 'R']: ws.range(f'{col}{row}').value = None

            window_row, door_row, total_exclude_area = 8, 26, 0.0
            sorted_items = sorted(items, key=lambda x: x['num'])
            
            for item in sorted_items:
                val = item['val']
                target_name = item['raw_name']
                
                if '両開き' in val or '片開き' in val:
                    if door_row > 28: continue
                    width = 0.98 if '両開き' in val else 0.965
                    height, u_value = 2.40, 2.91
                    ws.range(f'J{door_row}').value = target_name
                    ws.range(f'N{door_row}').value = width      
                    ws.range(f'P{door_row}').value = height     
                    ws.range(f'R{door_row}').value = u_value    
                    total_exclude_area += (width * height)
                    door_row += 1
                else:
                    if window_row > 19: continue
                    u_value, eta_value, width, height = parse_window_string(val, sash_u, shutter_u)
                    
                    ws.range(f'B{window_row}').value = target_name
                    ws.range(f'D{window_row}').value = width      
                    ws.range(f'F{window_row}').value = height     
                    ws.range(f'H{window_row}').value = u_value    
                    ws.range(f'J{window_row}').value = eta_value 
                    ws.range(f'AG{window_row}').value = True 
                    
                    total_exclude_area += (width * height)
                    window_row += 1
                    
            envelope_row = 35
            ws.range(f'J{envelope_row}').value = 'W-1'                    
            ws.range(f'R{envelope_row}').value = dynamic_u_values.get('W-1', 0.421)                    
            ws.range(f'N{envelope_row}').value = round(total_exclude_area, 3) 

        transfer_to_foundation_sheet(wb, foundation_data, dynamic_u_values)
        transfer_to_roof_floor_sheet(wb, roof_floor_data, dynamic_u_values)
        transfer_to_wall_area(wb, wall_area_data)

        if os.path.exists(output_filename): os.remove(output_filename)
        wb.save(output_filename)
        print(f"\n✅ 外皮計算書の転記が完了しました！ファイル【{os.path.basename(output_filename)}】を物件フォルダに保存しました。")
        return output_filename

    except Exception as e:
        print(f"\n⚠️ エラーが発生しました: {e}")
        return None
    finally:
        wb.close()
        app.quit()