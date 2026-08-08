import os
import datetime
from dxf_extractor import extract_dxf_groups

def main():
    # 1. 同フォルダ内から平面図bf（または平面図）DXFファイルを自動検出
    dxf_files = [f for f in os.listdir('.') if f.lower().endswith('.dxf')]
    
    heimen_bf_dxf = next((f for f in dxf_files if '平面図bf' in f or 'bf' in f.lower()), None)
    
    if not heimen_bf_dxf:
        # 見つからない場合は「平面図」が含まれるものをフォールバック検出
        heimen_bf_dxf = next((f for f in dxf_files if '平面図' in f), None)

    if not heimen_bf_dxf:
        print("⚠️ エラー: 「平面図bf」または「平面図」を含むDXFファイルが見つかりません。")
        return
        
    print(f"📄 平面図bfファイル: {heimen_bf_dxf} を検出しました。")

    # 出力ファイル名の設定
    now = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
    base_name = os.path.splitext(heimen_bf_dxf)[0]
    output_dxf_path = f"{base_name}_抽出結果_{now}.dxf"

    print("\n図面データから枠組みおよびレイヤー5の要素を抽出しています...")

    # 2. グループ抽出メソッドの実行
    g0_count, g1_count = extract_dxf_groups(heimen_bf_dxf, output_dxf_path)

    if g0_count > 0 or g1_count > 0:
        print("\n==================================================")
        print("✅ 枠組みおよびレイヤー5の抽出が完了しました！")
        print(f"・グループ0 (枠組み/断熱枠/方位1点): {g0_count} 件")
        print(f"・グループ1 (範囲内レイヤー5)      : {g1_count} 件")
        print(f"📁 出力DXFファイル: {output_dxf_path}")
        print("==================================================")
    else:
        print("\n⚠️ 該当する要素の抽出に失敗したか、データが見つかりませんでした。")

if __name__ == "__main__":
    main()