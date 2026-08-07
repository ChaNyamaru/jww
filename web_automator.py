import os
import re
import time
import math
import openpyxl
from decimal import Decimal, ROUND_HALF_UP
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from dxf_parser import read_dxf_data

class HouseAppAutomationMethods:

    @staticmethod
    def parse_target_value(text):
        """
        「＝」や「→」が含まれる場合は右側の数値を抽出する。
        「主1」のような漢字混じりの識別子は除外し、純粋な面積数値のみをパースする。
        """
        import re
        from decimal import Decimal, ROUND_HALF_UP

        text = text.strip()
        
        # 1. 「＝」や「=」「→」が含まれている場合、その後ろ（右側）だけを切り取る
        # （例: "主1 ＝ 25.92" が1つのテキストになっていた場合 -> " 25.92" になる）
        if '＝' in text or '=' in text:
            target_part = re.split(r'[＝=]', text)[-1].strip()
        elif '→' in text:
            target_part = text.split('→')[-1].strip()
        else:
            target_part = text
            
        # 2. 面積の文字列から、想定される単位や空白をすべて一時的に消す
        clean_text = target_part.replace(' ', '').replace(' ', '').replace(',', '')
        clean_text = clean_text.replace('㎡', '').replace('m2', '').replace('m²', '').replace('坪', '')
        
        # 3. ★ここがポイント★
        # 単位を消した後も「数字とピリオド」以外の文字（漢字やアルファベットなど）が残っているかチェック。
        # 残っている場合（例: target_part が "主1" の場合、clean_text も "主1" となる）、
        # それは面積ではなく識別子とみなして None を返し、スキップさせる。
        if re.search(r'[^\d\.]', clean_text):
            return None
            
        # 4. 確実に数値（面積）のみ残っている場合、数値を抽出して四捨五入
        match = re.search(r'\d+(?:\.\d+)?', target_part)
        if match:
            raw_val = match.group()
            dec_val = Decimal(raw_val)
            return float(dec_val.quantize(Decimal('0.01'), rounding=ROUND_HALF_UP))
            
        return None

    @classmethod
    def parse_dxf_floor_areas(cls, dxf_path: str) -> dict:
        """
        DXFの座標データを基に「主たる居室面積」「その他の居室面積」「合計面積」の
        文字列の右側にある数値を正確に最短距離で紐付けて抽出する
        """
        texts = read_dxf_data(dxf_path)
        
        target_tables = {
            '主たる居室面積': 'primary_area',
            'その他の居室面積': 'other_area',
            '合計': 'total_area'
        }
        
        results = {
            "primary_area": "0.00",
            "other_area": "0.00",
            "total_area": "0.00"
        }
        
        # ★改修: 右方向へ探す最大距離（6000mm）、Y方向（高さ）のズレ許容範囲（1000mm）に拡大
        NUM_X_TOLERANCE = 6000
        NUM_Y_TOLERANCE = 1000

        for search_word, key in target_tables.items():
            # 検索対象のラベルをすべて取得
            target_entities = [e for e in texts if search_word in e['text']]
            
            best_val = None
            min_dist = float('inf')
            
            # 複数の同じラベルがある場合も考慮し、最も数値が近くにあるものを正解とする
            for target_entity in target_entities:
                for e in texts:
                    dx = e['x'] - target_entity['x']
                    dy = e['y'] - target_entity['y']
                    
                    # 確実に右側にあり、Y方向のズレが許容範囲内のものを対象とする
                    if 0 < dx < NUM_X_TOLERANCE and abs(dy) < NUM_Y_TOLERANCE:
                        val = cls.parse_target_value(e['text'])
                        if val is not None:
                            dist = math.hypot(dx, dy) # 直線距離を計算
                            # 一番距離が近い（すぐ右隣にある）数値を採用
                            if dist < min_dist:
                                min_dist = dist
                                best_val = val
                                
            if best_val is not None:
                results[key] = str(best_val)

        return results

    @staticmethod
    def parse_excel_common_conditions(excel_path: str, house_address: str) -> dict:
        """
        Excelファイルの「共通条件・結果」シートから、指定されたセル番地（J11, J12, X11, X12）の値を抽出する。
        （プログラムで生成直後のExcelファイルは数式の計算結果が失われているため、Win32COMでExcelを起動して確実に取得する）
        """
        results = {}
        abs_excel_path = os.path.abspath(excel_path)
        
        def safe_float(val):
            try:
                if val is None:
                    return None
                if isinstance(val, (int, float)):
                    if val < 0: # 面積やUA値がマイナスになることはないため、COMのエラー値（-2146826281等）を弾く
                        return None
                    return float(val)
                match = re.search(r'\d+(?:\.\d+)?', str(val))
                if match:
                    return float(match.group())
            except Exception:
                pass
            return None

        # ★改修：1. COM経由でExcelを裏側で操作し、計算済みの値を確実に取得する
        try:
            import win32com.client
            import pythoncom
            pythoncom.CoInitialize() # スレッドエラー防止
            
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            try:
                # リンク更新なし、読み取り専用で開いて数式を強制計算させる
                wb = excel.Workbooks.Open(abs_excel_path, False, True)
                try:
                    ws = wb.Worksheets("共通条件・結果")
                except Exception:
                    ws = wb.Worksheets(1)
                
                # 指定セルからピンポイントで抽出
                val_area = safe_float(ws.Range("J11").Value)
                if val_area is not None: results["envelope_area_total"] = val_area
                    
                val_ua = safe_float(ws.Range("J12").Value)
                if val_ua is not None: results["avg_heat_transfer"] = val_ua
                    
                val_cool = safe_float(ws.Range("X11").Value)
                if val_cool is not None: results["cooling_solar_heat_gain"] = val_cool
                    
                val_heat = safe_float(ws.Range("X12").Value)
                if val_heat is not None: results["heating_solar_heat_gain"] = val_heat
                
                wb.Close(False)
            finally:
                excel.Quit()
        except Exception as e:
            print(f"⚠️ Excelの起動(COM操作)に失敗したため、openpyxlで代替抽出します: {e}")
            
            # 2. COMが使えない場合のフォールバック（保険）
            try:
                wb_ox = openpyxl.load_workbook(abs_excel_path, data_only=True)
                ws_ox = wb_ox["共通条件・結果"] if "共通条件・結果" in wb_ox.sheetnames else wb_ox.worksheets[0]
                
                if "envelope_area_total" not in results:
                    val = safe_float(ws_ox['J11'].value)
                    if val is not None: results["envelope_area_total"] = val
                if "avg_heat_transfer" not in results:
                    val = safe_float(ws_ox['J12'].value)
                    if val is not None: results["avg_heat_transfer"] = val
                if "cooling_solar_heat_gain" not in results:
                    val = safe_float(ws_ox['X11'].value)
                    if val is not None: results["cooling_solar_heat_gain"] = val
                if "heating_solar_heat_gain" not in results:
                    val = safe_float(ws_ox['X12'].value)
                    if val is not None: results["heating_solar_heat_gain"] = val
            except Exception as ox_e:
                print(f"⚠️ openpyxlでの抽出にも失敗しました: {ox_e}")

        # 地域区分の自動判定
        if "東広島市" in house_address:
            results["region"] = "５地域"
        else:
            results["region"] = "６地域"

        return results
    
    @classmethod
    def run_automation_flow(cls, housing_name: str, floor_areas: dict, common_data: dict, hot_water_type: str, download_dir: str):
        
        # =================================================================
        # 【変更点】新規起動ではなく、デバッグモードのChrome（ポート9222）に接続
        # =================================================================
        options = webdriver.ChromeOptions()
        options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
        
        try:
            driver = webdriver.Chrome(options=options)
            wait = WebDriverWait(driver, 10)
            print("🔗 すでに開いているChromeに接続しました！")
        except Exception as e:
            print("⚠️ Chromeに接続できませんでした。『自動化用Chrome起動.bat』でChromeが開かれているか確認してください。")
            raise e

        def set_input_value_by_js(element, value):
            driver.execute_script("""
                let el = arguments[0];
                let val = arguments[1];
                
                // ★ここが超重要★
                // もし取得した要素(appbox)が <input> ではなく枠組み(div)だった場合、
                // その中から実際の <input> を探し出してターゲットを切り替える！
                if (el.tagName.toLowerCase() !== 'input') {
                    let innerInput = el.querySelector('input');
                    if (innerInput) {
                        el = innerInput;
                    } else {
                        return;
                    }
                }
                
                el.removeAttribute('readonly');
                
                let nativeInputValueSetter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, "value").set;
                if (nativeInputValueSetter) {
                    nativeInputValueSetter.call(el, val);
                } else {
                    el.value = val;
                }
                
                el.dispatchEvent(new Event('input', { bubbles: true }));
                el.dispatchEvent(new Event('change', { bubbles: true }));
            """, element, value)

        try:
            # =================================================================
            # 0. 最初に必ず「基本情報」タブに戻る（前回のテストの続きから始まるのを防ぐ）
            # =================================================================
            print("🏠 基本情報タブへ移動（初期化）します...")
            driver.execute_script("""
                let tabs = Array.from(document.querySelectorAll('a, button, .v-tab'));
                let basicTab = tabs.find(t => t.textContent.includes('基本情報'));
                if(basicTab) basicTab.click();
            """)
            time.sleep(2) # 画面が切り替わるのを待つ

            # =================================================================
            # 1. 基本情報の入力
            # =================================================================
            print("📝 基本情報を入力中...")
            basic_fields = {
                "appbox-住宅タイプの名称": housing_name,
                "appbox-入力責任者": "新見尚己",
                "apptext-主たる居室": floor_areas.get("primary_area", ""),
                "apptext-その他の居室": floor_areas.get("other_area", ""),
                "apptext-非居室": floor_areas.get("non_residential_area", ""),
                "apptext-合計": floor_areas.get("total_area", "")
            }
            
            for testid, fval in basic_fields.items():
                try:
                    if fval and fval != "0.00":
                        elem = driver.find_element(By.CSS_SELECTOR, f'[data-testid="{testid}"]')
                        set_input_value_by_js(elem, fval)
                        print(f"✔️ {testid} に {fval} を入力しました。")
                except Exception:
                    print(f"⚠️ {testid} の入力欄が見つからないか、スキップしました。")

            # =================================================================
            # 2. 外皮項目の編集 (★ appbox- に戻しました)
            # =================================================================
            print("➡️ 外皮タブへ遷移します...")
            driver.execute_script("""
                let tabs = Array.from(document.querySelectorAll('a, button, .v-tab'));
                let envTab = tabs.find(t => t.textContent.includes('外皮'));
                if(envTab) envTab.click();
            """)
            time.sleep(2)
            
            env_fields = {
                "appbox-外皮面積の合計": common_data.get("envelope_area_total", ""),
                "appbox-冷房期の平均日射熱取得率（η<sub>AC</sub>）": common_data.get("cooling_solar_heat_gain", ""),
                "appbox-外皮平均熱貫流率（U<sub>A</sub>）": common_data.get("avg_heat_transfer", ""),
                "appbox-暖房期の平均日射熱取得率（η<sub>AH</sub>）": common_data.get("heating_solar_heat_gain", "")
            }
            
            for testid, fval in env_fields.items():
                if fval: 
                    try:
                        elem = driver.find_element(By.CSS_SELECTOR, f'[data-testid="{testid}"]')
                        set_input_value_by_js(elem, str(fval))
                        print(f"✔️ {testid} に {fval} を入力しました。")
                    except Exception:
                        print(f"⚠️ {testid} の入力欄が見つからないか、スキップしました。")

            # =================================================================
            # 3. 給湯項目の編集
            # =================================================================
            print("➡️ 給湯タブへ遷移します...")
            driver.execute_script("""
                let tabs = Array.from(document.querySelectorAll('a, button, .v-tab'));
                let hwTab = tabs.find(t => t.textContent.includes('給湯'));
                if(hwTab) hwTab.click();
            """)
            time.sleep(2)

            try:
                if hot_water_type == "ガス":
                    driver.execute_script("""
                        let labels = Array.from(document.querySelectorAll('label, .v-label'));
                        let target = labels.find(l => l.textContent.includes('ガス潜熱回収型') || l.textContent.includes('ガス'));
                        if(target) target.click();
                    """)
                elif hot_water_type in ["長方形エコキュート", "正方形エコキュート"]:
                    # 1. まず大元の「電気ヒートポンプ給湯機」を選択する
                    driver.execute_script("""
                        let elements = Array.from(document.querySelectorAll('label, .v-label, .v-radio'));
                        let target = elements.find(el => el.textContent.includes('電気ヒートポンプ給湯機') || el.textContent.includes('エコキュート'));
                        if(target) {
                            let input = target.querySelector('input[type="radio"]');
                            if(input) input.click();
                            else target.click();
                        }
                    """)
                    
                    # 項目が展開されるのを少し待つ
                    time.sleep(1) 
                    
                    # 2. ★追加★ 画像にあった「JIS効率を入力する」のラジオボタンをクリックする
                    driver.execute_script("""
                        let subElements = Array.from(document.querySelectorAll('label, .v-label, .v-radio'));
                        let subTarget = subElements.find(el => el.textContent.includes('JIS効率を入力する'));
                        if(subTarget) {
                            let input = subTarget.querySelector('input[type="radio"]');
                            if(input) input.click();
                            else subTarget.click();
                        }
                    """)
                    
                    # 3. JIS効率の入力欄がアニメーションで出現するのを待つ
                    time.sleep(1.5) 
                    
                    # 4. JIS効率の入力
                    jis_val = "3.0" if hot_water_type == "長方形エコキュート" else "3.6"
                    
                    jis_testid = "appbox-JIS効率" 
                    elem = driver.find_element(By.CSS_SELECTOR, f'[data-testid="{jis_testid}"]')
                    set_input_value_by_js(elem, jis_val)
                    print(f"✔️ {jis_testid} に {jis_val} を入力しました。")
                    
            except Exception as e:
                print(f"⚠️ 給湯設定の操作に失敗しました: {e}")

                # =================================================================
            # 4. 保存処理
            # =================================================================
            print("💾 保存処理を実行します...")
            
            try:
                # 1. 画面上部にある1回目の保存ボタンをクリック
                # (.title_command の中にある class="save" のボタンを指定)
                first_save_btn = driver.find_element(By.CSS_SELECTOR, '.title_command input.save')
                driver.execute_script("arguments[0].click();", first_save_btn)
                print("✔️ 1回目の「保存」をクリックしました！")
                
                # モーダル（ポップアップ）がアニメーションで表示されるのを少し待つ
                time.sleep(1.5) 
                
                # 2. ポップアップ画面の2回目の保存ボタンをクリック
                # (.modal-footer の中にある value="保存" のボタンを指定)
                modal_save_btn = driver.find_element(By.CSS_SELECTOR, '.modal-footer input[value="保存"]')
                driver.execute_script("arguments[0].click();", modal_save_btn)
                print("✔️ 完了ポップアップの「保存」をクリックしました！データの保存完了です。")
                
            except Exception as e:
                print(f"⚠️ 保存ボタンが見つからないか、クリックに失敗しました: {e}")

            # =================================================================
            # 5. 計算とPDF出力
            # =================================================================
            print("🧮 計算処理を実行します...")
            
            try:
                # 保存の直後なので、念のため画面が落ち着くのを少し待つ
                time.sleep(2)

                # 1. 「計算」ボタンをクリック（変わりやすい親クラスを無視して、直接「計算」ボタンを探す）
                driver.execute_script("""
                    let buttons = Array.from(document.querySelectorAll('input[type="button"], button'));
                    let calcBtn = buttons.find(b => b.value === '計算' || b.classList.contains('keisan'));
                    if(calcBtn) calcBtn.click();
                """)
                print("✔️ 「計算」をクリックしました。処理完了を待機します...")
                
                # 計算には少し時間がかかることがあるため長めに待機
                time.sleep(5) 
                
                # 2. 「PDFを出力する」ボタンをクリック（こちらも柔軟な探し方に変更）
                driver.execute_script("""
                    let buttons = Array.from(document.querySelectorAll('input[type="button"], button'));
                    let pdfBtn = buttons.find(b => b.value && b.value.includes('PDF') || b.classList.contains('uq_btnPdf'));
                    if(pdfBtn) pdfBtn.click();
                """)
                print("✔️ 「PDFを出力する」をクリックしました！")
                
                # 新しいタブが開く、または印刷ダイアログが出るのを待つ
                time.sleep(3) 

                # =============================================================
                # 3. 印刷ダイアログ（ChromeのUI）の突破
                # =============================================================
                if len(driver.window_handles) > 1:
                    driver.switch_to.window(driver.window_handles[-1])
                    print("📄 PDFのタブに切り替えました。")
                    time.sleep(1)
                
                from selenium.webdriver.common.action_chains import ActionChains
                from selenium.webdriver.common.keys import Keys
                
                # エンターキーを送信して青い「保存」ボタンを押す
                ActionChains(driver).send_keys(Keys.ENTER).perform()
                print("✔️ 印刷ダイアログで「保存（Enter）」を実行しました！")

            except Exception as e:
                print(f"⚠️ 計算またはPDF出力処理に失敗しました: {e}")

        except Exception as e:
            print(f"\n⚠️ 自動化処理中にエラーが発生しました: {e}")
            raise e
        finally:
            print("Pythonの処理を終了します。（Chromeは開いたまま維持します）")
            # 【変更点】デバッグモードのChromeが勝手に閉じないようにコメントアウト
            # driver.quit()
