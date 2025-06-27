import os
import re
import time
import logging
import pandas as pd
import argparse
import sys 

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, StaleElementReferenceException
from webdriver_manager.chrome import ChromeDriverManager

logging.basicConfig(
    level=logging.INFO,     
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler('judicial_scraper_true_final.log', encoding='utf-8')
    ]
)
logger = logging.getLogger(__name__)

class JudicialScraper:
    def __init__(self, target_name, keyword, max_records=100):
        self.driver = None
        self.data = []
        self.target_name = target_name
        self.keyword = keyword
        self.max_records = max_records
        self.main_window_handle = None
 
    def setup_driver(self):
        logger.info("⚙️  設定 WebDriver...")
        options = Options()
        options.add_experimental_option("detach", True)
        options.add_argument("--disable-blink-features=AutomationControlled")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("start-maximized")
        # Add user agent to avoid detection
        options.add_argument("--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")
        service = Service(ChromeDriverManager().install())
        self.driver = webdriver.Chrome(service=service, options=options)
        logger.info("✅  WebDriver 設定完成。")
 
    def auto_input_and_search(self):
        self.driver.get("https://judgment.judicial.gov.tw/FJUD/default.aspx")
        logger.info(f"🔍  正在使用關鍵字 '{self.keyword}' 進行查詢...")
        try:
            input_box = WebDriverWait(self.driver, 15).until(EC.presence_of_element_located((By.ID, "txtKW")))
            input_box.clear()
            input_box.send_keys(self.keyword)
            search_button = self.driver.find_element(By.ID, "btnSimpleQry")
            search_button.click()
                         
            logger.info("⏳  等待並取得查詢 QID...")
            WebDriverWait(self.driver, 15).until(EC.presence_of_element_located((By.ID, "hidQID")))
            qid = self.driver.find_element(By.ID, "hidQID").get_attribute("value")
                         
            results_url = f"https://judgment.judicial.gov.tw/FJUD/qryresultlst.aspx?ty=JUDBOOK&q={qid}"
            logger.info(f"🚀  取得 QID，直接跳轉至結果頁...")
            self.driver.get(results_url)
 
            # 等待搜尋結果表格載入
            table_selector = (By.CLASS_NAME, "jub-table")
            logger.info("⏳  等待搜尋結果表格 (class='jub-table') 載入...")
            WebDriverWait(self.driver, 20).until(EC.presence_of_element_located(table_selector))
                         
            logger.info("✅  結果頁表格已成功載入！")
            self.main_window_handle = self.driver.current_window_handle
        except Exception as e:
            logger.error(f"❌ 搜尋或跳轉過程出錯：{e}", exc_info=True)
            raise

    def parse_onclick_params(self, onclick_str):
        """解析 onclick 參數並構建判決書 URL"""
        try:
            # 從 onclick="cookieId('TCDV%2c114%2c%e6%8a%97%2c141%2c20250430%2c1','0','abc177...','','','DS','1')" 
            # 提取參數
            match = re.search(r"cookieId\('([^']+)','([^']+)','([^']+)','([^']*)','([^']*)','([^']*)','([^']*)'\)", onclick_str)
            if match:
                param1, param2, qid, param4, param5, param6, param7 = match.groups()
                
                # 構建判決書詳細頁面 URL
                # 根據司法院網站的 URL 結構
                detail_url = f"https://judgment.judicial.gov.tw/FJUD/data.aspx?ty=JD&id={param1}&q={qid}"
                return detail_url
            else:
                logger.warning(f"無法解析 onclick 參數: {onclick_str}")
                return None
        except Exception as e:
            logger.error(f"解析 onclick 參數時出錯: {e}")
            return None

    def extract_judgment_info_from_row(self, row):
        """從表格行中提取判決資訊"""
        try:
            tds = row.find_elements(By.TAG_NAME, "td")
            if len(tds) < 4:
                return None
            
            # 第二個 td 包含判決字號和連結
            judgment_td = tds[1]
            
            # 尋找連結
            links = judgment_td.find_elements(By.TAG_NAME, "a")
            if not links:
                return None
            
            link = links[0]
            onclick_attr = link.get_attribute('onclick')
            judgment_number = link.text.strip()
            
            # 第三個 td 包含日期
            date_text = tds[2].text.strip()
            
            # 第四個 td 包含案由
            case_reason = tds[3].text.strip()
            
            # 解析 onclick 參數構建 URL
            detail_url = self.parse_onclick_params(onclick_attr)
            
            if detail_url:
                return {
                    "url": detail_url,
                    "text": judgment_number,
                    "date": date_text,
                    "case_reason": case_reason
                }
            else:
                return None
                
        except Exception as e:
            logger.error(f"提取判決資訊時出錯: {e}")
            return None

    def extract_court_info(self, judgment_number):
        """從判決字號中提取法院資訊"""
        try:
            # 判決字號格式通常為：法院簡稱+年份+案件類別+案件號碼
            # 例如：臺北地方法院112年度執字第12345號
            court_patterns = [
                r'(.*?地方法院)',
                r'(.*?高等法院)', 
                r'(.*?法院)',
                r'(最高法院)',
                r'(司法院)'
            ]
            
            for pattern in court_patterns:
                match = re.search(pattern, judgment_number)
                if match:
                    return match.group(1)
            
            # 如果找不到完整法院名稱，嘗試從簡稱推斷
            court_abbreviations = {
                '北院': '臺北地方法院',
                '新院': '新北地方法院', 
                '桃院': '桃園地方法院',
                '中院': '臺中地方法院',
                '南院': '臺南地方法院',
                '高院': '高等法院',
                '最高院': '最高法院'
            }
            
            for abbr, full_name in court_abbreviations.items():
                if abbr in judgment_number:
                    return full_name
                    
            return "未知法院"
        except Exception as e:
            logger.error(f"提取法院資訊時出錯: {e}")
            return "未知法院"

    def extract_case_type_and_year(self, judgment_number):
        """從判決字號中提取案件類型和年度"""
        try:
            # 提取年度
            year_match = re.search(r'(\d{2,3})年', judgment_number)
            year = f"民國{year_match.group(1)}年" if year_match else "未知年度"
            
            # 提取案件類型
            case_type_match = re.search(r'年度?(\w+)字', judgment_number)
            case_type = case_type_match.group(1) if case_type_match else "未知案件類型"
            
            return year, case_type
        except Exception as e:
            logger.error(f"提取案件類型和年度時出錯: {e}")
            return "未知年度", "未知案件類型"
 
    def parse_result_cards(self):
        logger.info(f"🔍  開始擷取完整內文與分析資料（上限 {self.max_records} 筆）...")
        processed_count = 0
        page_count = 1
                 
        while processed_count < self.max_records:
            logger.info(f"📄  處理第 {page_count} 頁...")
            
            # ========== DEBUG: 檢查頁面內容 ==========
            print("DEBUG: 當前頁面 URL:", self.driver.current_url)
            
            try:
                # 等待表格載入
                WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, "jub-table")))
                
                # 找到主要的判決表格
                table = self.driver.find_element(By.CLASS_NAME, "jub-table")
                rows = table.find_elements(By.TAG_NAME, "tr")
                
                print(f"DEBUG: 找到表格，共 {len(rows)} 行")
                
                links_to_process = []
                
                # 跳過表頭，從第二行開始處理
                for i, row in enumerate(rows[1:], 1):
                    try:
                        judgment_info = self.extract_judgment_info_from_row(row)
                        if judgment_info:
                            links_to_process.append(judgment_info)
                            print(f"DEBUG: 第 {i} 行 - 判決字號: {judgment_info['text']}, 日期: {judgment_info['date']}")
                    except Exception as e:
                        print(f"DEBUG: 處理第 {i} 行時出錯: {e}")
                        continue
                
            except TimeoutException:
                logger.warning("⚠️  當前頁面找不到判決表格。")
                break
            except Exception as e:
                logger.error(f"❌ 尋找判決表格時出錯: {e}")
                break
            
            print(f"DEBUG: 成功解析 {len(links_to_process)} 個連結")
            if links_to_process:
                print(f"DEBUG: 第一個連結範例: {links_to_process[0]}")
            # ========== DEBUG 結束 ==========
                         
            for link_info in links_to_process:
                if processed_count >= self.max_records:
                    logger.info(f"已達到最大擷取筆數 {self.max_records}，停止處理。")
                    break
                                 
                try:
                    logger.info(f"📖  處理第 {processed_count + 1} 筆: {link_info['text']}")
                    
                    # 開啟新分頁並導航到判決詳細頁面
                    self.driver.execute_script("window.open(arguments[0]);", link_info['url'])
                    self.driver.switch_to.window(self.driver.window_handles[-1])
                    
                    # 等待頁面載入
                    time.sleep(3)
                                         
                    try:
                        # 修正的內容擷取方式
                        full_text = ""
                        
                        # 嘗試多種方式獲取判決內容
                        content_selectors = [
                            (By.ID, "jud_content"),  # 判決內容
                            (By.CLASS_NAME, "jud-content"),  # 判決內容 class
                            (By.ID, "content"),  # 通用內容ID
                            (By.CLASS_NAME, "content"),  # 通用內容class
                            (By.CSS_SELECTOR, ".mainContent"),  # 主要內容區域
                            (By.CSS_SELECTOR, "#mainContent"),  # 主要內容區域ID
                            (By.CSS_SELECTOR, "[id*='content']"),  # 包含 content 的 ID
                            (By.CSS_SELECTOR, "[class*='content']"),  # 包含 content 的 class
                            (By.TAG_NAME, "body")  # 最後備用選項
                        ]
                        
                        for selector in content_selectors:
                            try:
                                WebDriverWait(self.driver, 5).until(EC.presence_of_element_located(selector))
                                content_element = self.driver.find_element(*selector)
                                full_text = content_element.text
                                if full_text and len(full_text) > 100:  # 確保獲取到足夠的內容
                                    logger.info(f"✅  成功使用選擇器 {selector} 獲取內容 ({len(full_text)} 字元)")
                                    break
                            except:
                                continue
                        
                        if not full_text:
                            logger.warning("⚠️  無法獲取判決內容，嘗試獲取頁面所有文字")
                            full_text = self.driver.find_element(By.TAG_NAME, "body").text
                        
                        # 使用從表格中獲取的案由，如果沒有則從內容中提取
                        case_reason = link_info.get('case_reason', '')
                        if not case_reason:
                            reason_match = re.search(r'裁判案由[：:\s]*([^\n\r]+)', full_text)
                            if reason_match:
                                case_reason = reason_match.group(1).strip()
                            else:
                                case_reason = "未知案由"
                        
                        # 提取法院和案件資訊
                        court_name = self.extract_court_info(link_info['text'])
                        year, case_type = self.extract_case_type_and_year(link_info['text'])
                        
                        name_matches = self.extract_names_and_roles(full_text)
                        target_role = self.determine_target_role(name_matches, self.target_name)
                        
                        # 轉換日期格式
                        formatted_date = self.convert_date_format(link_info['date'])
                        
                        # 建立完整的資料記錄
                        record = {
                            "序號": processed_count + 1,
                            "搜尋關鍵字": self.keyword,
                            "目標人物": self.target_name,
                            "判決字號": link_info['text'],
                            "法院名稱": court_name,
                            "裁判年度": year,
                            "案件類型": case_type,
                            "裁判日期": formatted_date,
                            "裁判案由": case_reason,
                            "目標人物身份": target_role,
                            "所有當事人": ", ".join(list(dict.fromkeys([name for _, name in name_matches]))),
                            "當事人角色分配": "; ".join([f"{role}:{name}" for role, name in name_matches]),
                            "判決書連結": link_info['url'],
                            "內容長度": len(full_text),
                            "擷取時間": time.strftime("%Y-%m-%d %H:%M:%S")
                        }
                        
                        self.data.append(record)
                        processed_count += 1
                        logger.info(f"✅  成功處理，獲取 {len(full_text)} 字元內容。")
                        
                    except Exception as e:
                        logger.error(f"❌  讀取或分析內文時出錯: {e}")
                        # 即使出錯也記錄基本資訊
                        court_name = self.extract_court_info(link_info['text'])
                        year, case_type = self.extract_case_type_and_year(link_info['text'])
                        formatted_date = self.convert_date_format(link_info['date'])
                        
                        record = {
                            "序號": processed_count + 1,
                            "搜尋關鍵字": self.keyword,
                            "目標人物": self.target_name,
                            "判決字號": link_info['text'],
                            "法院名稱": court_name,
                            "裁判年度": year,
                            "案件類型": case_type,
                            "裁判日期": formatted_date,
                            "裁判案由": link_info.get('case_reason', '擷取失敗'),
                            "目標人物身份": "擷取失敗",
                            "所有當事人": "",
                            "當事人角色分配": "",
                            "判決書連結": link_info['url'],
                            "內容長度": 0,
                            "擷取時間": time.strftime("%Y-%m-%d %H:%M:%S")
                        }
                        self.data.append(record)
                        processed_count += 1
                    finally:
                        self.driver.close()
                        self.driver.switch_to.window(self.main_window_handle)
                        time.sleep(1)  # 短暫停頓避免過於頻繁的請求
                except Exception as e:
                    logger.warning(f"⚠️  處理單筆連結時發生未知錯誤: {e}")
 
            if processed_count >= self.max_records: break
 
            try:
                next_btn = WebDriverWait(self.driver, 5).until(EC.element_to_be_clickable((By.ID, "hlNext")))
                logger.info("🔄  點擊下一頁...")
                next_btn.click()
                time.sleep(2)  # 等待頁面載入
                page_count += 1
            except TimeoutException:
                logger.info("ℹ️  已達最後一頁。")
                break
                         
        logger.info(f"✅  全部擷取完成，共處理 {len(self.data)} 筆有效記錄")

    def convert_date_format(self, date_str):
        """轉換民國年日期為西元年格式"""
        try:
            date_match = re.search(r'(\d{2,3})\.(\d{1,2})\.(\d{1,2})', date_str)
            if date_match:
                year = int(date_match.group(1)) + 1911
                month = int(date_match.group(2))
                day = int(date_match.group(3))
                return f"{year}-{month:02d}-{day:02d}"
            else:
                return date_str
        except Exception as e:
            logger.error(f"日期轉換錯誤: {e}")
            return date_str
 
    def export_to_excel(self, records, filename="judicial_result.xlsx"):
        """匯出資料到 Excel 檔案"""
        if not records:
            logger.warning("⚠️  查無結果可匯出")
            return
        
        try:
            # 建立 DataFrame
            df = pd.DataFrame(records)
            
            # 按日期排序（最新的在前）
            df = df.sort_values(by=["裁判日期"], ascending=False)
            
            # 設定輸出路徑
            output_path = os.path.join(os.path.expanduser("~"), "Desktop", filename)
            
            # 匯出到 Excel
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                # 主要資料表
                df.to_excel(writer, sheet_name='判決資料', index=False)
                
                # 統計資料表
                summary_data = self.create_summary_data(df)
                summary_df = pd.DataFrame(summary_data)
                summary_df.to_excel(writer, sheet_name='統計摘要', index=False)
                
                # 調整欄位寬度
                for sheet_name in writer.sheets:
                    worksheet = writer.sheets[sheet_name]
                    for column in worksheet.columns:
                        max_length = 0
                        column_letter = column[0].column_letter
                        
                        for cell in column:
                            try:
                                if len(str(cell.value)) > max_length:
                                    max_length = len(str(cell.value))
                            except:
                                pass
                        
                        adjusted_width = min(max_length + 3, 50)
                        worksheet.column_dimensions[column_letter].width = adjusted_width
            
            logger.info(f"📁  已成功匯出 {len(df)} 筆資料到 {output_path}")
            
        except Exception as e:
            logger.error(f"❌  匯出 Excel 時發生錯誤: {e}")
            # 嘗試匯出到 CSV 作為備用方案
            try:
                csv_path = output_path.replace('.xlsx', '.csv')
                df = pd.DataFrame(records)
                df.to_csv(csv_path, index=False, encoding='utf-8-sig')
                logger.info(f"📁  已備用匯出 CSV 格式到 {csv_path}")
            except Exception as csv_error:
                logger.error(f"❌  CSV 匯出也失敗: {csv_error}")

    def create_summary_data(self, df):
        """建立統計摘要資料"""
        try:
            summary = []
            
            # 基本統計
            summary.append({"項目": "總判決書數量", "數值": len(df), "說明": "本次搜尋共找到的判決書數量"})
            summary.append({"項目": "搜尋關鍵字", "數值": self.keyword, "說明": "使用的搜尋關鍵字"})
            summary.append({"項目": "目標人物", "數值": self.target_name, "說明": "分析的目標人物姓名"})
            
            # 法院統計
            court_counts = df['法院名稱'].value_counts()
            summary.append({"項目": "最常見法院", "數值": court_counts.index[0] if len(court_counts) > 0 else "無", 
                          "說明": f"出現 {court_counts.iloc[0]} 次" if len(court_counts) > 0 else ""})
            
            # 身份統計
            role_counts = df['目標人物身份'].value_counts()
            summary.append({"項目": "最常見身份", "數值": role_counts.index[0] if len(role_counts) > 0 else "無",
                          "說明": f"出現 {role_counts.iloc[0]} 次" if len(role_counts) > 0 else ""})
            
            # 案件類型統計
            case_type_counts = df['案件類型'].value_counts()
            summary.append({"項目": "最常見案件類型", "數值": case_type_counts.index[0] if len(case_type_counts) > 0 else "無",
                          "說明": f"出現 {case_type_counts.iloc[0]} 次" if len(case_type_counts) > 0 else ""})
            
            # 年度分布
            year_counts = df['裁判年度'].value_counts()
            summary.append({"項目": "案件年度分布", "數值": f"{len(year_counts)} 個年度", 
                          "說明": f"從 {year_counts.index[-1]} 到 {year_counts.index[0]}" if len(year_counts) > 0 else ""})
            
            return summary
            
        except Exception as e:
            logger.error(f"建立統計摘要時出錯: {e}")
            return [{"項目": "統計錯誤", "數值": "無法產生統計", "說明": str(e)}]
 
    def extract_names_and_roles(self, full_text):
        patterns = [
            r'(原告|被告|抗告人|相對人|上訴人|被上訴人|聲請人|債務人|債權人|第三人)[：:\s]*([^\s，。；、\n\r\t]{2,20})',
            r'([^\s，。；、\n\r\t]{2,20})\s*(?:為|係|即)\s*(原告|被告|抗告人|相對人|上訴人|被上訴人|聲請人|債務人|債權人|第三人)'
        ]
        all_matches = []
        matches1 = re.findall(patterns[0], full_text)
        all_matches.extend(matches1)
        matches2 = re.findall(patterns[1], full_text)
        all_matches.extend([(match[1], match[0]) for match in matches2])
        cleaned_matches = []
        seen_combinations = set()
        for role, name in all_matches:
            name = re.sub(r'[^\u4e00-\u9fff\u0030-\u0039A-Za-z]', '', name).strip()
            if 2 <= len(name) <= 10 and (role, name) not in seen_combinations:
                cleaned_matches.append((role, name))
                seen_combinations.add((role, name))
        return cleaned_matches
 
    def determine_target_role(self, name_matches, target_name):
        target_roles = []
        for role, name in name_matches:
            if target_name in name or name in target_name:
                target_roles.append(role)
        if target_roles:
            priority_order = ['債務人', '被告', '抗告人', '上訴人', '被上訴人', '債權人', '原告', '聲請人', '相對人', '第三人']
            for priority_role in priority_order:
                if priority_role in target_roles:
                    return priority_role
            return target_roles[0]
        return "其他"
 
    def run(self):
        try:
            self.setup_driver()
            self.auto_input_and_search()
            self.parse_result_cards()
            self.export_to_excel(self.data)
        except Exception as e:
            logger.error(f"❌ 程式執行期間發生頂層錯誤: {e}", exc_info=True)
        finally:
            if self.driver:
                logger.info("🛑 程式執行完畢。")
                self.driver.quit()

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description="司法院裁判書系統爬蟲工具 (報表修正版)")
    parser.add_argument("target_name", type=str, help="目標人物姓名, e.g., \"許家瑋\"")
    parser.add_argument("keyword", type=str, help="搜尋關鍵字, e.g., \"許家瑋 本票裁定\"")
    parser.add_argument("max_records", type=int, help="最大擷取筆數, e.g., 5")
         
    try:
        args = parser.parse_args()
        print(f"參數設定: 目標人物={args.target_name}, 關鍵字={args.keyword}, 最大筆數={args.max_records}")
        scraper = JudicialScraper(
            target_name=args.target_name, 
            keyword=args.keyword, 
            max_records=args.max_records
        )
        scraper.run()
    except SystemExit as e:
        if e.code != 0:  # 只有在真正錯誤時才顯示幫助
            print("\n使用方式:")
            print("python fixed02.py <目標人物姓名> <搜尋關鍵字> <最大筆數>")
            print("範例: python fixed02.py \"許家瑋\" \"許家瑋 本票裁定\" 100")
        pass