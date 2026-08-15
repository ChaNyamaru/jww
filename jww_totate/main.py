import datetime
import os
import glob
import tkinter as tk
import pprint
from tkinter import ttk

from dxf_parser import read_dxf_data, calculate_proximate_texts, extract_foundation_data, extract_roof_floor_data, extract_wall_area_data, save_to_excel
from excel_filler import transfer_to_envelope_calc_sheet
from web_automator import HouseAppAutomationMethods
from application_filler import fill_application_forms 

def get_user_input():
    input_data = {}

    def on_submit():
        input_data["house_name"] = entry_name.get().strip()
        input_data["house_address"] = entry_address.get().strip()
        input_data["house_area"] = entry_area.get().strip()
        input_data["start_date"] = entry_date.get().strip()
        input_data["user_input"] = combo_door.get().strip()
        input_data["hot_water_type"] = combo_water.get().strip()
        
        # ★追加: サッシのU値入力
        input_data["sash_u"] = entry_sash_u.get().strip()
        input_data["shutter_u"] = entry_shutter_u.get().strip()
        
        root.destroy()

    root = tk.Tk()
    root.title("エネルギー消費性能計算 - 入力フォーム")
    # 入力項目が増えたので縦幅をさらに拡張
    root.geometry("450x550")
    root.attributes('-topmost', True)

    tk.Label(root, text="住宅の名称", font=("", 10, "bold")).pack(pady=(15, 0))
    entry_name = tk.Entry(root, width=50)
    entry_name.pack()

    tk.Label(root, text="住宅の所在地", font=("", 10, "bold")).pack(pady=(10, 0))
    entry_address = tk.Entry(root, width=50)
    entry_address.pack()

    tk.Label(root, text="建築物の延べ面積 (㎡)", font=("", 10, "bold")).pack(pady=(10, 0))
    entry_area = tk.Entry(root, width=50)
    entry_area.pack()

    tk.Label(root, text="新築竣工時期/着工予定時期 (例: 2026/08/10)", font=("", 10, "bold")).pack(pady=(10, 0))
    entry_date = tk.Entry(root, width=50)
    entry_date.pack()

    tk.Label(root, text="玄関扉の種類", font=("", 10, "bold")).pack(pady=(10, 0))
    combo_door = ttk.Combobox(root, values=["片開き", "両開き"], state="readonly", width=47)
    combo_door.current(0) 
    combo_door.pack()

    tk.Label(root, text="給湯設備の種類", font=("", 10, "bold")).pack(pady=(10, 0))
    combo_water = ttk.Combobox(root, values=["ガス", "長方形エコキュート", "正方形エコキュート"], state="readonly", width=47)
    combo_water.current(0) 
    combo_water.pack()

    # ★追加: U値の入力フィールド
    tk.Label(root, text="サッシのU値 (W/m2K)", font=("", 10, "bold")).pack(pady=(10, 0))
    entry_sash_u = tk.Entry(root, width=50)
    entry_sash_u.insert(0, "2.33") # デフォルト値を入れておく
    entry_sash_u.pack()

    tk.Label(root, text="シャッター付きサッシのU値 (HS文字付) (W/m2K)", font=("", 10, "bold")).pack(pady=(10, 0))
    entry_shutter_u = tk.Entry(root, width=50)
    entry_shutter_u.insert(0, "2.11") # デフォルト値を入れておく
    entry_shutter_u.pack()

    tk.Button(root, text="自動計算を開始する", command=on_submit, bg="#add8e6", font=("", 11, "bold")).pack(pady=20)

    root.mainloop()
    return input_data

def main():
    dxf_files = [f for f in os.listdir('.') if f.lower().endswith('.dxf')]
    
    heimen_dxf = next((f for f in dxf_files if '平面図' in f), None)
    ritumen_dxf = next((f for f in dxf_files if '立面図' in f), None)
    menseki_dxf = next((f for f in dxf_files if '面積表' in f), None)
    
    if not heimen_dxf:
        print("⚠️ エラー: 「平面図」を含むDXFファイルが見つかりません。")
        return
    print(f"📄 平面図ファイル: {heimen_dxf} を検出しました。")

    if not ritumen_dxf:
        print("⚠️ エラー: 「立面図」を含むDXFファイルが見つかりません。")
        return
    print(f"📄 立面図ファイル: {ritumen_dxf} を検出しました。")

    if not menseki_dxf:
        print("⚠️ エラー: 「面積表」を含むDXFファイルが見つかりません。")
        return
    print(f"📄 面積表ファイル: {menseki_dxf} を検出しました。")

    print("\n入力フォームを立ち上げます...")
    user_data = get_user_input()
    
    if not user_data:
        print("❌ 処理がキャンセルされました。")
        return
        
    house_name = user_data.get("house_name", "（名称未入力）")
    house_address = user_data.get("house_address", "（所在地未入力）")
    house_area = user_data.get("house_area", "")
    start_date = user_data.get("start_date", "")
    user_input = user_data.get("user_input", "（手動入力なし）")
    hot_water_type = user_data.get("hot_water_type", "ガス")
    
    # ★追加: サッシU値をFloatに変換（空欄や文字列が入った場合はデフォルト値にフォールバック）
    try: sash_u = float(user_data.get("sash_u"))
    except: sash_u = 2.33
    try: shutter_u = float(user_data.get("shutter_u"))
    except: shutter_u = 2.11

    DISTANCE_LIMIT = 1000

    safe_house_name = "".join(c for c in house_name if c not in r'\/:*?"<>|')
    output_dir = os.path.abspath(safe_house_name)
    os.makedirs(output_dir, exist_ok=True)

    print(f"\n✅ 入力を受け付けました。名称: {house_name} / 所在地: {house_address}")
    print(f"📁 保存先フォルダを作成しました: {output_dir}")
    print("図面データを解析しています...")
    
    heimen_texts = read_dxf_data(heimen_dxf)
    extracted_data = calculate_proximate_texts(heimen_texts, DISTANCE_LIMIT, user_input)
    foundation_data = extract_foundation_data(heimen_texts, DISTANCE_LIMIT)
    roof_floor_data = extract_roof_floor_data(heimen_texts)
    
    wall_area_data = {}
    if ritumen_dxf:
        ritumen_texts = read_dxf_data(ritumen_dxf)
        wall_area_data = extract_wall_area_data(ritumen_texts)
    
    # ★変更: 引数に sash_u, shutter_u を追加
    envelope_excel_path = transfer_to_envelope_calc_sheet(
        extracted_data, foundation_data, roof_floor_data, wall_area_data, 
        house_name, house_address, output_dir, sash_u, shutter_u
    )

    print("\n床面積データを抽出しています...")
    floor_areas = HouseAppAutomationMethods.parse_dxf_floor_areas(menseki_dxf)
    
    if envelope_excel_path and os.path.exists(envelope_excel_path):
        common_data = HouseAppAutomationMethods.parse_excel_common_conditions(envelope_excel_path, house_address)
    else:
        print("⚠️ 警告: 外皮計算書のExcelファイルが見つかりませんでした。")
        common_data = {"region": "５地域" if "東広島市" in house_address else "６地域"}

    print("ブラウザを起動し、エネルギー消費性能計算プログラムへの入力を開始します...")
    calc_results = HouseAppAutomationMethods.run_automation_flow(
        housing_name=house_name,
        floor_areas=floor_areas,
        common_data=common_data,
        hot_water_type=hot_water_type,
        download_dir=output_dir
    )

    fill_application_forms(house_name, house_address, common_data, calc_results, envelope_excel_path, house_area, start_date, output_dir)
    
    print("\n✨ すべての自動化処理が完了しました！")

if __name__ == "__main__":
    main()