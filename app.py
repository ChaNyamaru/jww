import datetime
import os
import glob
import streamlit as st
import tempfile # アップロードされたファイルを一時保存するため

# 外部モジュール（同じフォルダにある前提）
from dxf_parser import read_dxf_data, calculate_proximate_texts, extract_foundation_data, extract_roof_floor_data, extract_wall_area_data, save_to_excel
from excel_filler import transfer_to_envelope_calc_sheet
from web_automator import HouseAppAutomationMethods

# --- Streamlitの画面設定 ---
st.set_page_config(page_title="エネルギー消費性能計算フォーム", layout="centered")

def main():
    st.title("エネルギー消費性能計算 - 入力フォーム")
    st.markdown("必要な情報とDXFファイルを入力・アップロードしてください。")

    # --- 1. ユーザー入力フォーム ---
    with st.form("input_form"):
        house_name = st.text_input("住宅の名称", placeholder="例：山田邸")
        house_address = st.text_input("住宅の所在地", placeholder="例：東京都...")
        
        user_input = st.selectbox("玄関扉の種類", ["片開き", "両開き"])
        hot_water_type = st.selectbox("給湯設備の種類", ["ガス", "長方形エコキュート", "正方形エコキュート"])
        
        # --- 2. ファイルのアップロード ---
        st.markdown("### 図面ファイル (DXF)")
        heimen_file = st.file_uploader("平面図のDXFファイルをアップロード", type=['dxf'])
        ritumen_file = st.file_uploader("立面図のDXFファイルをアップロード", type=['dxf'])
        menseki_file = st.file_uploader("面積表のDXFファイルをアップロード", type=['dxf'])

        # 実行ボタン
        submitted = st.form_submit_button("自動計算を開始する")

    # --- 3. 実行ボタンが押された後の処理 ---
    if submitted:
        # 必須入力のチェック
        if not heimen_file or not ritumen_file or not menseki_file:
            st.error("エラー: 平面図、立面図、面積表のすべてのDXFファイルをアップロードしてください。")
            return
            
        # 未入力項目のフォールバック
        house_name = house_name if house_name else "（名称未入力）"
        house_address = house_address if house_address else "（所在地未入力）"
        user_input = user_input if user_input else "（手動入力なし）"
        hot_water_type = hot_water_type if hot_water_type else "ガス"

        st.success(f"入力を受け付けました。名称: {house_name} / 所在地: {house_address}")
        
        # 進行状況を表示するスピナー
        with st.spinner('図面データを解析しています...'):
            
            # --- アップロードされたファイルを一時的に保存 ---
            # DXFパーサーがファイルパスを要求する作りのため、一時フォルダに保存する
            temp_dir = tempfile.mkdtemp()
            heimen_dxf = os.path.join(temp_dir, heimen_file.name)
            ritumen_dxf = os.path.join(temp_dir, ritumen_file.name)
            menseki_dxf = os.path.join(temp_dir, menseki_file.name)

            with open(heimen_dxf, "wb") as f:
                f.write(heimen_file.getbuffer())
            with open(ritumen_dxf, "wb") as f:
                f.write(ritumen_file.getbuffer())
            with open(menseki_dxf, "wb") as f:
                f.write(menseki_file.getbuffer())

            DISTANCE_LIMIT = 1000
            
            try:
                # 3. 平面図からのデータ抽出
                st.text("平面図を解析中...")
                heimen_texts = read_dxf_data(heimen_dxf)
                extracted_data = calculate_proximate_texts(heimen_texts, DISTANCE_LIMIT, user_input)
                foundation_data = extract_foundation_data(heimen_texts, DISTANCE_LIMIT)
                roof_floor_data = extract_roof_floor_data(heimen_texts)
                
                # 4. 立面図からのデータ抽出
                st.text("立面図を解析中...")
                ritumen_texts = read_dxf_data(ritumen_dxf)
                wall_area_data = extract_wall_area_data(ritumen_texts)
                
                # 5. 外皮計算書への転記
                st.text("外皮計算書（Excel）を作成中...")
                transfer_to_envelope_calc_sheet(extracted_data, foundation_data, roof_floor_data, wall_area_data, house_name, house_address)

                # 6. Web自動化連携のためのデータ準備
                st.text("床面積データを抽出中...")
                floor_areas = HouseAppAutomationMethods.parse_dxf_floor_areas(menseki_dxf)
                
                # 生成されたExcelファイルを探す
                excel_files = [f for f in glob.glob(f'*外皮計算書ver2.4（{house_name}）*.xlsx') if not os.path.basename(f).startswith('~$')]
                
                if excel_files:
                    excel_path = os.path.abspath(excel_files[0])
                    common_data = HouseAppAutomationMethods.parse_excel_common_conditions(excel_path, house_address)
                else:
                    st.warning("警告: 外皮計算書のExcelファイルが見つかりませんでした。デフォルト値を使用します。")
                    common_data = {"region": "５地域" if "東広島市" in house_address else "６地域"}

                st.info("データの解析が完了しました。ブラウザ自動操作を開始します...")

                # 7. ブラウザ自動操作の実行
                HouseAppAutomationMethods.run_automation_flow(
                    housing_name=house_name,
                    floor_areas=floor_areas,
                    common_data=common_data,
                    hot_water_type=hot_water_type,
                    download_dir=os.getcwd()
                )
                
                st.success("🎉 全ての処理が完了しました！")

            except Exception as e:
                st.error(f"処理中にエラーが発生しました: {e}")

if __name__ == "__main__":
    main()