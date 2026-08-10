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
        if '＝' in text or '=' in text:
            target_part = re.split(r'[＝=]', text)[-1].strip()
        elif '→' in text:
            target_part = text.split('→')[-1].strip()
        else:
            target_part = text
            
        # 2. 面積の文字列から、想定される単位や空白をすべて一時的に消す
        clean_text = target_part.replace(' ', '').replace(' ', '').replace(',', '')
        clean_text = clean_text.replace('㎡', '').replace('m2', '').replace('m²', '').replace('坪', '')
        
        # 3. 単位を消した後も「数字とピリオド」以外の文字（漢字やアルファベットなど）が残っているかチェック
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
        
        NUM_X_TOLERANCE = 6000
        NUM_Y_TOLERANCE = 1000

        for search_word, key in target_tables.items():
            target_entities = [e for e in texts if search_word in e['text']]
            
            best_val = None
            min_dist = float('inf')
            
            for target_entity in target_entities:
                for e in texts:
                    dx = e['x'] - target_entity['x']
                    dy = e['y'] - target_entity['y']
                    
                    if 0 < dx < NUM_X_TOLERANCE and abs(dy) < NUM_Y_TOLERANCE:
                        val = cls.parse_target_value(e['text'])
                        if val is not None:
                            dist = math.hypot(dx, dy) 
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
        """
        results = {}
        abs_excel_path = os.path.abspath(excel_path)
        
        def safe_float(val):
            try:
                if val is None:
                    return None
                if isinstance(val, (int, float)):
                    if val < 0:
                        return None
                    return float(val)
                match = re.search(r'\d+(?:\.\d+)?', str(val))
                if match:
                    return float(match.group())
            except Exception:
                pass
            return None

        try:
            import win32com.client
            import pythoncom
            pythoncom.CoInitialize()
            
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            try:
                wb = excel.Workbooks.Open(abs_excel_path, False, True)
                try:
                    ws = wb.Worksheets("共通条件・結果")
                except Exception:
                    ws = wb.Worksheets(1)
                
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

        if "東広島市" in house_address:
            results["region"] = "５地域"
        else:
            results["region"] = "６地域"

        return results
    
    @classmethod
    def run_automation_flow(cls, housing_name: str, floor_areas: dict, common_data: dict, hot_water_type: str, download_dir: str):
        
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
            # 0. 基本情報タブへ移動
            print("🏠 基本情報タブへ移動（初期化）します...")
            driver.execute_script("""
                let tabs = Array.from(document.querySelectorAll('a, button, .v-tab'));
                let basicTab = tabs.find(t => t.textContent.includes('基本情報'));
                if(basicTab) basicTab.click();
            """)
            time.sleep(2)

            # 1. 基本情報の入力
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
            # 2. 外皮項目の編集
            # =================================================================
            print("➡️ 外皮タブへ遷移します...")
            driver.execute_script("""
                let tabs = Array.from(document.querySelectorAll('a, button, .v-tab'));
                let envTab = tabs.find(t => t.textContent.includes('外皮'));
                if(envTab) envTab.click();
            """)
            time.sleep(2)
            
            # --- ★追加：暖房期日射熱取得率の小数点第2位以下切り捨て処理 ---
            from decimal import Decimal, ROUND_DOWN
            heating_val = common_data.get("heating_solar_heat_gain", "")
            cooling_val = common_data.get("cooling_solar_heat_gain", "")
            avgheat_val = common_data.get("avg_heat_transfer", "")
            if heating_val != "":
                try:
                    # '0.1' で小数点第1位まで残し、第2位以降を切り捨てる（ROUND_DOWN）
                    # （例: 1.39 → 1.3）
                    # ※もし「第2位まで残す（例: 1.395 → 1.39）」という意味であれば、 Decimal('0.01') に書き換えてください。
                    heating_val = float(Decimal(str(heating_val)).quantize(Decimal('0.1'), rounding=ROUND_DOWN))
                except Exception:
                    pass

            if cooling_val != "":
                try:
                    cooling_val = float(Decimal(str(cooling_val)).quantize(Decimal('0.1'), rounding=ROUND_DOWN))
                except Exception:
                    pass

            if avgheat_val != "":
                try:
                    avgheat_val = float(Decimal(str(avgheat_val)).quantize(Decimal('0.01'), rounding=ROUND_DOWN))   
                except Exception:
                    pass
                

            # -----------------------------------------------------------
            
            env_fields = {
                "appbox-外皮面積の合計": common_data.get("envelope_area_total", ""),
                "appbox-冷房期の平均日射熱取得率（η<sub>AC</sub>）": cooling_val,
                "appbox-外皮平均熱貫流率（U<sub>A</sub>）": avgheat_val,
                "appbox-暖房期の平均日射熱取得率（η<sub>AH</sub>）": heating_val
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
            # 3. 暖房・冷房タブの編集
            # =================================================================
            print("➡️ 暖房タブへ遷移します...")
            driver.execute_script("""
                let tabs = Array.from(document.querySelectorAll('a, button, .v-tab'));
                let heatTab = tabs.find(t => t.textContent.includes('暖房'));
                if(heatTab) heatTab.click();
            """)
            time.sleep(1.5)
            # 「設置しない」を選択
            driver.execute_script("""
                let labels = Array.from(document.querySelectorAll('label, .v-label'));
                let target = labels.find(l => l.textContent.includes('設置しない'));
                if(target) target.click();
            """)
            print("✔️ 暖房設備を「設置しない」に設定しました。")

            print("➡️ 冷房タブへ遷移します...")
            driver.execute_script("""
                let tabs = Array.from(document.querySelectorAll('a, button, .v-tab'));
                let coolTab = tabs.find(t => t.textContent.includes('冷房'));
                if(coolTab) coolTab.click();
            """)
            time.sleep(1.5)
            # 「設置しない」を選択
            driver.execute_script("""
                let labels = Array.from(document.querySelectorAll('label, .v-label'));
                let target = labels.find(l => l.textContent.includes('設置しない'));
                if(target) target.click();
            """)
            print("✔️ 冷房設備を「設置しない」に設定しました。")


            # =================================================================
            # 4. 換気タブの編集
            # =================================================================
            print("➡️ 換気タブへ遷移します...")
            driver.execute_script("""
                let tabs = Array.from(document.querySelectorAll('a, button, .v-tab'));
                let ventTab = tabs.find(t => t.textContent.includes('換気'));
                if(ventTab) ventTab.click();
            """)
            time.sleep(1.5)

            # 換気1：ダクト式第二種 または ダクト式第三種 を選択
            driver.execute_script("""
                let labels = Array.from(document.querySelectorAll('label, .v-label'));
                let v1 = labels.find(l => l.textContent.includes('壁付け式第二種換気設備') || l.textContent.includes('壁付け式第三種換気設備'));
                if(v1) v1.click();
            """)
            time.sleep(1)

            # 換気2：比消費電力「入力する」を選択
            driver.execute_script("""
                let labels = Array.from(document.querySelectorAll('label, .v-label'));
                let v2 = labels.find(l => (l.textContent.trim() === '入力する' && !l.textContent.includes('効率')) || l.textContent.includes('比消費電力を入力'));
                if(v2) v2.click();
            """)
            time.sleep(1.5)

            # 比消費電力 0.07 を入力 (表示された最初のテキストボックス)
            driver.execute_script("""
                let inputs = Array.from(document.querySelectorAll('input[type="text"], input[type="number"]'));
                let visibleInputs = inputs.filter(i => i.offsetParent !== null && !i.readOnly && !i.disabled);
                if (visibleInputs.length > 0) {
                    let targetInput = visibleInputs[0];
                    targetInput.removeAttribute('readonly');
                    targetInput.value = '0.07';
                    targetInput.dispatchEvent(new Event('input', { bubbles: true }));
                    targetInput.dispatchEvent(new Event('change', { bubbles: true }));
                }
            """)
            print("✔️ 比消費電力を「0.07」に設定しました。")
            
            # 換気3：換気回数 0.5回/h を選択
            driver.execute_script("""
                let labels = Array.from(document.querySelectorAll('label, .v-label'));
                let v3 = labels.find(l => l.textContent.includes('0.5回/h') || l.textContent.includes('0.5回/ｈ'));
                if(v3) v3.click();
            """)
            print("✔️ 換気回数を「0.5回/h」に設定しました。")
            time.sleep(1)


            # =================================================================
            # 5. 給湯タブの編集
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
                        let elements = Array.from(document.querySelectorAll('label, .v-label, .v-radio'));
                        // 「ガス」だけだと従来型を誤爆するので、固有の文字列を指定する
                        let target = elements.find(el => el.textContent.includes('ガス潜熱回収型'));
                        if(target) {
                            let input = target.querySelector('input[type="radio"]');
                            if(input) input.click();
                            else target.click(); // 🌟エコキュートと同じく、ラベルのクリックを発動させる
                        }
                    """)
                    time.sleep(1)
                    
                    # 効率（エネルギー消費効率）を入力を選択
                    driver.execute_script("""
                        let labels = Array.from(document.querySelectorAll('label, .v-label'));
                        let eff_input = labels.find(l => l.textContent.includes('効率（エネルギー消費効率）を入力') || l.textContent.includes('効率を入力'));
                        if(!eff_input) eff_input = labels.find(l => l.textContent.trim() === '入力する');
                        if(eff_input) eff_input.click();
                    """)
                    time.sleep(1.5)
                    
                    # エネルギー消費効率 91.5 を入力
                    driver.execute_script("""
                        let inputs = Array.from(document.querySelectorAll('input[type="text"], input[type="number"]'));
                        let visibleInputs = inputs.filter(i => i.offsetParent !== null && !i.readOnly && !i.disabled);
                        if (visibleInputs.length > 0) {
                            let targetInput = visibleInputs[0];
                            targetInput.removeAttribute('readonly');
                            targetInput.value = '91.5';
                            targetInput.dispatchEvent(new Event('input', { bubbles: true }));
                            targetInput.dispatchEvent(new Event('change', { bubbles: true }));
                        }
                    """)
                    print("✔️ ガスのエネルギー消費効率を「91.5」に設定しました。")
                    
                elif hot_water_type in ["長方形エコキュート", "正方形エコキュート"]:
                    driver.execute_script("""
                        let elements = Array.from(document.querySelectorAll('label, .v-label, .v-radio'));
                        let target = elements.find(el => el.textContent.includes('電気ヒートポンプ給湯機') || el.textContent.includes('エコキュート'));
                        if(target) {
                            let input = target.querySelector('input[type="radio"]');
                            if(input) input.click();
                            else target.click();
                        }
                    """)
                    time.sleep(1) 
                    
                    driver.execute_script("""
                        let subElements = Array.from(document.querySelectorAll('label, .v-label, .v-radio'));
                        let subTarget = subElements.find(el => el.textContent.includes('JIS効率を入力する'));
                        if(subTarget) {
                            let input = subTarget.querySelector('input[type="radio"]');
                            if(input) input.click();
                            else subTarget.click();
                        }
                    """)
                    time.sleep(1.5) 
                    
                    jis_val = "3.0" if hot_water_type == "長方形エコキュート" else "3.6"
                    jis_testid = "appbox-JIS効率" 
                    try:
                        elem = driver.find_element(By.CSS_SELECTOR, f'[data-testid="{jis_testid}"]')
                        set_input_value_by_js(elem, jis_val)
                        print(f"✔️ エコキュートのJIS効率に「{jis_val}」を設定しました。")
                    except Exception:
                        pass
                
                time.sleep(1)

                # NEW: 給湯5 (配管方式と配管径) と 給湯7 (浴槽の保温措置)
                driver.execute_script("""
                    let labels = Array.from(document.querySelectorAll('label, .v-label'));
                    
                    // 配管方式: ヘッダー方式
                    let header = labels.find(l => l.textContent.includes('ヘッダー方式'));
                    if(header) header.click();
                """)
                time.sleep(0.5)
                
                driver.execute_script("""
                    let labels = Array.from(document.querySelectorAll('label, .v-label'));
                    
                    // 配管径: すべての配管径が13A以下
                    let pipe = labels.find(l => l.textContent.includes('すべての配管径が13A以下') || l.textContent.includes('13A以下'));
                    if(pipe) pipe.click();
                    
                    // 浴槽の保温措置: 高断熱浴槽を使用する
                    let bath = labels.find(l => l.textContent.includes('高断熱浴槽を使用する'));
                    if(bath) bath.click();
                """)
                print("✔️ ヘッダー方式（13A以下）と高断熱浴槽の設定を行いました。")
                
            except Exception as e:
                print(f"⚠️ 給湯設定の操作に失敗しました: {e}")


            # =================================================================
            # 6. 保存処理
            # =================================================================
            print("💾 保存処理を実行します...")
            try:
                first_save_btn = driver.find_element(By.CSS_SELECTOR, '.title_command input.save')
                driver.execute_script("arguments[0].click();", first_save_btn)
                print("✔️ 1回目の「保存」をクリックしました！")
                
                time.sleep(1.5) 
                
                modal_save_btn = driver.find_element(By.CSS_SELECTOR, '.modal-footer input[value="保存"]')
                driver.execute_script("arguments[0].click();", modal_save_btn)
                print("✔️ 完了ポップアップの「保存」をクリックしました！データの保存完了です。")
                
            except Exception as e:
                print(f"⚠️ 保存ボタンが見つからないか、クリックに失敗しました: {e}")

            time.sleep(3) 

           # =================================================================
            # 7. 計算と結果抽出、そしてPDF出力
            # =================================================================
            print("🧮 計算処理を実行します...")
            try:
                time.sleep(2)
                driver.execute_script("""
                    let buttons = Array.from(document.querySelectorAll('input[type="button"], button'));
                    let calcBtn = buttons.find(b => b.value === '計算' || b.classList.contains('keisan'));
                    if(calcBtn) calcBtn.click();
                """)
                print("✔️ 「計算」をクリックしました。処理完了を待機します...")
                
                # 計算完了までしっかり待つ
                time.sleep(5) 
                
                # ★修正：PDF出力ボタンを押す「前」に、安全にBEIなどを抽出しておく
                print("📊 計算結果画面からBEI・エネルギー消費量を抽出しています...")
                calc_results = None
                try:
                    calc_results = driver.execute_script("""
                        let results = { design_energy: "", standard_energy: "", bei: "" };
                        
                        // 設計一次エネルギー消費量（最後の行の合計値を取得）
                        let designEls = document.querySelectorAll('td.result_calclated');
                        if(designEls.length > 0) {
                            let el = designEls[designEls.length - 1].cloneNode(true);
                            let span = el.querySelector('span');
                            if(span) el.removeChild(span); // "GJ"の文字を消す
                            results.design_energy = el.textContent.trim();
                        }
                        
                        // 基準一次エネルギー消費量（最後の行の合計値を取得）
                        let standardEls = document.querySelectorAll('td.standard');
                        if(standardEls.length > 0) {
                            let el = standardEls[standardEls.length - 1].cloneNode(true);
                            let span = el.querySelector('span');
                            if(span) el.removeChild(span); // "GJ"の文字を消す
                            results.standard_energy = el.textContent.trim();
                        }
                        
                        // BEI
                        let beiEl = document.querySelector('td.bei_value');
                        if(beiEl) results.bei = beiEl.textContent.trim();
                        
                        return results;
                    """)
                    print(f"✔️ 抽出結果: {calc_results}")
                except Exception as e:
                    print(f"⚠️ 計算結果の抽出に失敗しました: {e}")

            ## PDF出力処理    

            except Exception as e:
                print(f"⚠️ 計算またはPDF出力処理に失敗しました: {e}")

            #input("\n✅ PDFの保存が完了したら、この黒い画面（ターミナル）をクリックして Enterキー を押してください（ブラウザが閉じます）...")
            
            # 抽出した結果をメイン処理（Excel転記）に返す
            return calc_results

        except Exception as e:
            print(f"\n⚠️ 自動化処理中にエラーが発生しました: {e}")
            raise e
        finally:
            print("Pythonの処理を終了します。（Chromeは開いたまま維持します）")