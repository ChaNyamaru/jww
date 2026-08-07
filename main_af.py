import datetime
import os
import glob
import tkinter as tk
import pprint
from tkinter import ttk

from dxf_parser import read_dxf_data, calculate_proximate_texts, extract_foundation_data, extract_roof_floor_data, extract_wall_area_data, save_to_excel
from excel_filler import transfer_to_envelope_calc_sheet
from web_automator import HouseAppAutomationMethods

def get_user_input():
    """GUIを表示し、ユーザーの入力を辞書形式で返す関数"""
    input_data = {}

    def on_submit():
        # 実行ボタンが押されたら、入力された値を取得して辞書に格納
        input_data["house_name"] = entry_name.get().strip()
        input_data["house_address"] = entry_address.get().strip()
        input_data["user_input"] = combo_door.get().strip()
        input_data["hot_water_type"] = combo_water.get().strip()
        
        # ウィンドウを閉じて次の処理へ進む
        root.destroy()

    # ウィンドウの作成
    root = tk.Tk()
    root.title("エネルギー消費性能計算 - 入力フォーム")
    root.geometry("450x350")
    
    # ウィンドウを常に最前面に表示（他のウィンドウに隠れないようにする）
    root.attributes('-topmost', True)

    # --- 住宅の名称 ---
    tk.Label(root, text="住宅の名称", font=("", 10, "bold")).pack(pady=(15, 0))
    entry_name = tk.Entry(root, width=50)
    entry_name.pack()

    # --- 住宅の所在地 ---
    tk.Label(root, text="住宅の所在地", font=("", 10, "bold")).pack(pady=(15, 0))
    entry_address = tk.Entry(root, width=50)
    entry_address.pack()

    # --- 玄関扉の種類（プルダウン） ---
    tk.Label(root, text="玄関扉の種類", font=("", 10, "bold")).pack(pady=(15, 0))
    combo_door = ttk.Combobox(root, values=["片開き", "両開き"], state="readonly", width=47)
    combo_door.current(0) # デフォルト設定
    combo_door.pack()

    # --- 給湯設備の種類（プルダウン） ---
    tk.Label(root, text="給湯設備の種類", font=("", 10, "bold")).pack(pady=(15, 0))
    combo_water = ttk.Combobox(root, values=["ガス", "長方形エコキュート", "正方形エコキュート"], state="readonly", width=47)
    combo_water.current(0) # デフォルト設定
    combo_water.pack()

    # --- 実行ボタン ---
    tk.Button(root, text="自動計算を開始する", command=on_submit, bg="#add8e6", font=("", 11, "bold")).pack(pady=25)

    # ウィンドウを表示して待機
    root.mainloop()

    return input_data

def main():
    # 1. 同フォルダ内からDXFファイルを自動検出
    dxf_files = [f for f in os.listdir('.') if f.lower().endswith('.dxf')]
    
    heimen_dxf = next((f for f in dxf_files if '平面図' in f), None)
    ritumen_dxf = next((f for f in dxf_files if '立面図' in f), None)
    menseki_dxf = next((f for f in dxf_files if '面積表' in f), None)
    
    if not heimen_dxf:
        print("$26A0$FE0F エラー: 「平面図」を含むDXFファイルが見つかりません。")
        return
        
    print(f"$D83D$DCC4 平面図ファイル: {heimen_dxf} を検出しました。")

    if not ritumen_dxf:
            print("$26A0$FE0F エラー: 「立面図」を含むDXFファイルが見つかりません。")
            return
            
    print(f"$D83D$DCC4 立面図ファイル: {ritumen_dxf} を検出しました。")

    if not menseki_dxf:
                print("$26A0$FE0F エラー: 「面積表」を含むDXFファイルが見つかりません。")
                return
                
    print(f"$D83D$DCC4 面積表ファイル: {menseki_dxf} を検出しました。")

    # 2. GUIを立ち上げてユーザー入力を取得
    print("\n入力フォームを立ち上げます...")
    user_data = get_user_input()
    
    # 途中でウィンドウの×ボタンが押された場合などのキャンセル処理
    if not user_data:
        print("$274C 処理がキャンセルされました。")
        return
        
    house_name = user_data.get("house_name")
    house_address = user_data.get("house_address")
    user_input = user_data.get("user_input")
    hot_water_type = user_data.get("hot_water_type")
    
    # 空欄だった場合のフォールバック処理
    if not house_name:
        house_name = "（名称未入力）"
    if not house_address:
        house_address = "（所在地未入力）"
    if not user_input:
        user_input = "（手動入力なし）"
    if not hot_water_type:
        hot_water_type = "ガス"

    DISTANCE_LIMIT = 1000
    now = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
    base_name = os.path.splitext(heimen_dxf)[0]
    EXCEL_FILE_PATH = f"{base_name}_抽出結果_{now}.xlsx"

    print(f"\n$2705 入力を受け付けました。名称: {house_name} / 所在地: {house_address}")
    print("図面データを解析しています...")
    
    # 3. 平面図からのデータ抽出
    heimen_texts = read_dxf_data(heimen_dxf)
    extracted_data = calculate_proximate_texts(heimen_texts, DISTANCE_LIMIT, user_input)
    foundation_data = extract_foundation_data(heimen_texts, DISTANCE_LIMIT)
    roof_floor_data = extract_roof_floor_data(heimen_texts)
    
    # 4. 立面図からのデータ抽出
    wall_area_data = {}
    if ritumen_dxf:
        ritumen_texts = read_dxf_data(ritumen_dxf)
        wall_area_data = extract_wall_area_data(ritumen_texts)
    
    # 5. 外皮計算書への転記
    transfer_to_envelope_calc_sheet(extracted_data, foundation_data, roof_floor_data, wall_area_data, house_name, house_address)

    # 6. エネルギー消費性能計算プログラムへのWeb自動化連携
    print("\n床面積データを抽出しています...")
    floor_areas = HouseAppAutomationMethods.parse_dxf_floor_areas(menseki_dxf)
    
    # 生成された外皮計算書ファイルから外皮データを再読み込み
    # （glob.globが重複していた部分を整理しました）
    excel_files = [f for f in glob.glob(f'*外皮計算書ver2.4（{house_name}）*.xlsx') if not os.path.basename(f).startswith('~$')]
    
    if excel_files:
        excel_path = os.path.abspath(excel_files[0])
        common_data = HouseAppAutomationMethods.parse_excel_common_conditions(excel_path, house_address)
    else:
        print("$26A0$FE0F 警告: 外皮計算書のExcelファイルが見つかりませんでした。")
        common_data = {"region": "５地域" if "東広島市" in house_address else "６地域"}

    # =====================================================================
    # 確認用（デバッグ）
    # =====================================================================
    #print("\n" + "="*50)
    #print("$D83D$DD0D 【デバッグ用】 自動入力前に抽出データを確認します")
    #print("="*50)
    #print("【DXF抽出結果 (床面積)】:")
    #pprint.pprint(floor_areas)
    #print("-" * 50)
    #print("【Excel抽出結果 (外皮・共通条件)】:")
    #pprint.pprint(common_data)
    #print("="*50 + "\n")
    # =====================================================================

    # もし「値がちゃんと取れているかだけ確認したい」場合は、
    # 以下の HouseAppAutomationMethods.run_automation_flow(...) の先頭に # を付けてコメントアウトしてください。

    print("ブラウザを起動し、エネルギー消費性能計算プログラムへの入力を開始します...")
    HouseAppAutomationMethods.run_automation_flow(
        housing_name=house_name,
        floor_areas=floor_areas,
        common_data=common_data,
        hot_water_type=hot_water_type,
        download_dir=os.getcwd()
    )

if __name__ == "__main__":
    main()