import os
import ezdxf
from ezdxf import recover
from shapely.geometry import Polygon, LineString, Point

def fix_jwcad_dxf_bug(file_path):
    """
    Jw_cadが書き出したDXFのOBJECTSセクションにある致命的な構文バグを、
    ezdxfが読み込めるように一時的に修正（置換）した一時ファイルを生成する。
    """
    temp_path = file_path + ".fixed.dxf"
    try:
        with open(file_path, 'rb') as f:
            raw_data = f.read()
        
        # Jw_cad特有の無効なルート辞書宣言を、標準のDICTIONARYに書き換える
        fixed_data = raw_data.replace(b'ACDBDICTIONARYWDFLT', b'DICTIONARY')
        
        with open(temp_path, 'wb') as f:
            f.write(fixed_data)
            
        return temp_path
    except Exception as e:
        print(f"⚠️ DXF修復処理に失敗しました: {e}")
        return file_path # 失敗した場合は元のファイルを返す

def check_group0_entity(entity, found_north_flag):
    """
    エンティティがグループ0（枠組み・断熱文字＆白枠・方位1点）に該当するか判定する関数
    """
    layer_name = entity.dxf.layer
    dxftype = entity.dxftype()
    
    is_g0 = False
    is_frame_line = False
    updated_north_flag = found_north_flag

    insulation_keywords = ['屋根断熱', '下屋', '外気に接する床', '外気床']

    # A. レイヤー0系（枠組み）
    if layer_name in ['0', '00-通り芯', '00-枠']:
        is_g0 = True
        if dxftype == 'LINE':
            is_frame_line = True

    # B. テキスト判定（断熱範囲・方位）
    elif dxftype in ['TEXT', 'MTEXT']:
        text_val = entity.dxf.text if dxftype == 'TEXT' else entity.plain_text()
        
        # 断熱設置範囲テキスト
        if any(kw in text_val for kw in insulation_keywords):
            is_g0 = True
        # 方位（最初に見つかった1つだけを保持）
        elif ('真北' in text_val or '北' in text_val) and not found_north_flag:
            is_g0 = True
            updated_north_flag = True

    # C. 断熱設置範囲の白枠（Color 7 または ByLayer白）
    elif layer_name in ['02-仕上', '0'] and entity.dxf.color in [7, 256]:
        is_g0 = True
        if dxftype == 'LINE':
            is_frame_line = True

    return is_g0, is_frame_line, updated_north_flag

def create_boundary_polygon(frame_lines):
    """
    枠組み線群からバウンディングボックス（ポリゴン）を生成する関数
    """
    if not frame_lines:
        return None
        
    all_x = [pt[0] for line in frame_lines for pt in line]
    all_y = [pt[1] for line in frame_lines for pt in line]
    
    min_x, max_x = min(all_x), max(all_x)
    min_y, max_y = min(all_y), max(all_y)
    
    if abs(max_x - min_x) < 1.0 or abs(max_y - min_y) < 1.0:
        return None
    
    return Polygon([
        (min_x, min_y),
        (max_x, min_y),
        (max_x, max_y),
        (min_x, max_y)
    ])

def is_entity_inside_polygon(entity, boundary_polygon):
    """
    エンティティが枠組みポリゴンの内部/交差領域にあるか判定する関数
    """
    if boundary_polygon is None:
        return True

    dxftype = entity.dxftype()
    try:
        # ★修正：スライス（[:2]）を廃止し、.x と .y プロパティを明示的に取得する
        if dxftype == 'LINE':
            start_pt = (entity.dxf.start.x, entity.dxf.start.y)
            end_pt = (entity.dxf.end.x, entity.dxf.end.y)
            line_geom = LineString([start_pt, end_pt])
            return boundary_polygon.intersects(line_geom)
        elif dxftype in ['TEXT', 'MTEXT']:
            point_geom = Point(entity.dxf.insert.x, entity.dxf.insert.y)
            return boundary_polygon.contains(point_geom)
        else:
            if hasattr(entity.dxf, 'insert'):
                return boundary_polygon.contains(Point(entity.dxf.insert.x, entity.dxf.insert.y))
            return True
    except Exception:
        return True

def extract_dxf_groups(input_dxf_path, output_dxf_path):
    """
    平面図bfからグループ0およびグループ1(レイヤー5)を抽出し、新しいDXFファイルへ出力する関数
    """
    if not os.path.exists(input_dxf_path):
        print(f"⚠️ エラー: ファイルが見つかりません: {input_dxf_path}")
        return 0, 0

    # Jw_cad出力DXFの構文バグを事前に修正した一時ファイルを作成
    fixed_dxf_path = fix_jwcad_dxf_bug(input_dxf_path)

    try:
        doc_in = ezdxf.readfile(fixed_dxf_path)
    except ezdxf.DXFStructureError as e:
        print(f"⚠️ 深刻なDXF構造エラーが残っています: {e}")
        return 0, 0

    msp_in = doc_in.modelspace()

    doc_out = ezdxf.new(setup=True)
    msp_out = doc_out.modelspace()

    for layer in doc_in.layers:
        if layer.dxf.name not in doc_out.layers:
            doc_out.layers.new(name=layer.dxf.name, dxfattribs={'color': layer.dxf.color})

    group0_entities = []
    frame_lines = []
    found_north = False

    # 1. グループ0の抽出
    for entity in msp_in:
        is_g0, is_frame_line, found_north = check_group0_entity(entity, found_north)
        if is_g0:
            group0_entities.append(entity)
            if is_frame_line:
                # ★修正：スライス（[:2]）を廃止し、.x と .y を取得してタプルにする
                start = (entity.dxf.start.x, entity.dxf.start.y)
                end = (entity.dxf.end.x, entity.dxf.end.y)
                frame_lines.append((start, end))

    # 2. 枠組みポリゴンの形成
    boundary_polygon = create_boundary_polygon(frame_lines)

    # 3. グループ1（レイヤー5）の抽出
    group1_entities = []
    for entity in msp_in:
        layer_name = entity.dxf.layer
        if '5' in layer_name or '寸法' in layer_name:
            if is_entity_inside_polygon(entity, boundary_polygon):
                group1_entities.append(entity)

    # 4. 新規DXFへのエンティティ書き出し
    for ent in group0_entities:
        msp_out.add_entity(ent.copy())

    for ent in group1_entities:
        msp_out.add_entity(ent.copy())

    doc_out.saveas(output_dxf_path)
    
    # 使い終わった一時ファイルを削除
    if fixed_dxf_path != input_dxf_path and os.path.exists(fixed_dxf_path):
        os.remove(fixed_dxf_path)

    return len(group0_entities), len(group1_entities)