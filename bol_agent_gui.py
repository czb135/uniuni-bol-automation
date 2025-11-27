import tkinter as tk
from tkinter import scrolledtext, messagebox
import threading
import time
from datetime import datetime

# Selenium 库
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys

# ================= 1. 地址大字典 (根据你的历史数据整理) =================
# 这里的 Key 是你在微信里写的简写，Value 是表单下拉框里要求的完整格式
ADDRESS_MAP = {
    # --- 常用大仓 (带星号格式) ---
    "LAX": "*LAX162*: 16288 Boyle Ave, Fontana CA 92337",
    "LAX162": "*LAX162*: 16288 Boyle Ave, Fontana CA 92337",
    "EWR": "*EWR600*: 600 Federal Blvd, Carteret NJ 07008",
    "EWR600": "*EWR600*: 600 Federal Blvd, Carteret NJ 07008",
    "NJ600": "*EWR600*: 600 Federal Blvd, Carteret NJ 07008",
    "JFK": "*JFK175*: 175-14 147th Ave, Jamaica NY 11434",
    "JFK175": "*JFK175*: 175-14 147th Ave, Jamaica NY 11434",
    "ORD": "*ORD121*: 1211 Tower Road, Schaumburg IL 60173",
    "ORD121": "*ORD121*: 1211 Tower Road, Schaumburg IL 60173",
    "DFW": "*DFW445*: 4450 W Walnut Hill Lane, Unit 100, Irving TX 75038",
    "DFW445": "*DFW445*: 4450 W Walnut Hill Lane, Unit 100, Irving TX 75038",
    "ATL": "*ATL441*: 4411 Bibb Boulevard, Tucker GA 30084",
    "ATL441": "*ATL441*: 4411 Bibb Boulevard, Tucker GA 30084",
    "MIA": "*MIA307*: 3075 NW 107th Ave, Doral FL 33172",
    "MIA307": "*MIA307*: 3075 NW 107th Ave, Doral FL 33172",

    # --- 卫星仓/其他站点 (不带星号格式) ---
    "BOS": "BOS001: 1 Wesley St, Malden, MA 02148",
    "BOS001": "BOS001: 1 Wesley St, Malden, MA 02148",
    "PHL": "PHL160: 1601 Boulevard Ave, Pennsauken NJ 08110",
    "PHL160": "PHL160: 1601 Boulevard Ave, Pennsauken NJ 08110",
    "DCA": "DCA522: 5225 Kilmer Place, Hyattsville MD 20781",
    "DCA522": "DCA522: 5225 Kilmer Place, Hyattsville MD 20781",
    "RDU": "RDU550: 5504 Caterpillar Dr, Apex NC 27539",
    "HFD": "HFD045: 45 Gracey Ave, Meriden CT 06451",
    "BDL": "HFD045: 45 Gracey Ave, Meriden CT 06451",
    "ORF": "ORF271: 271 Benton Road, Suffolk VA 23434",
    "DOV": "DOV011: 11 S Dupont Blvd, Milford DE 19963",
    "PVD": "PVD031: 31 Graystone St, Warwick RI 02886",
    "NJ25": "EWR025: 25 Amor Ave, Carlstadt NJ 07072",
    "EWR025": "EWR025: 25 Amor Ave, Carlstadt NJ 07072",
    "ORD102": "ORD102: 10216 Werch Dr, Woodridge IL 60517",
    "ATL760": "ATL760: 7600 Wood Rd, Douglasville GA 30134",
    "RIC": "RIC100: 10097 Patterson Park Rd, Suite 101, Ashland VA 23005",
}

# ================= 2. 业务规则逻辑 =================

def get_carrier(destination_key):
    dest = destination_key.upper()
    # 规则 1: Han Express (EWR/JFK)
    if "EWR" in dest or "JFK" in dest: return "Han Express"
    # 规则 2: NYQZ (ATL/MIA)
    if "ATL" in dest or "MIA" in dest: return "NYQZ"
    # 规则 3: 80s Express (中西部/东部卫星仓)
    if any(k in dest for k in ["ORD", "DFW", "BOS", "PHL", "DCA", "RDU", "HFD", "ORF", "DOV", "PVD", "WHS","RIC"]):
        return "80s Express"
    # 默认
    return "Spot Freight"

def get_pallet_count(destination_key):
    dest = destination_key.upper()
    # 规则: 短途12板，长途26板
    short_haul = ["EWR", "JFK", "NJ25", "PHL", "DCA", "BOS", "HFD", "PVD", "DOV"]
    if any(k in dest for k in short_haul):
        return 12
    return 26

# ================= 3. GUI 主程序 =================

class BOLAgentApp:
    def __init__(self, root):
        self.root = root
        self.root.title("UniUni BOL 自动开单机器人 (最终稳定版)")
        self.root.geometry("650x750")

        # 配置区
        config_frame = tk.LabelFrame(root, text="基础配置", padx=10, pady=10)
        config_frame.pack(fill="x", padx=10, pady=5)

        tk.Label(config_frame, text="Batch Number (批次号):").grid(row=0, column=0, sticky="w")
        self.entry_batch = tk.Entry(config_frame, width=35, fg="gray")
        self.entry_batch.grid(row=0, column=1, padx=5)
        self.entry_batch.insert(0, "请输入当日批次号")
        self.entry_batch.bind("<FocusIn>", lambda e: self._on_entry_focus_in(self.entry_batch, "请输入当日批次号"))
        self.entry_batch.bind("<FocusOut>", lambda e: self._on_entry_focus_out(self.entry_batch, "请输入当日批次号"))

        tk.Label(config_frame, text="Email (接收邮箱):").grid(row=1, column=0, sticky="w")
        self.entry_email = tk.Entry(config_frame, width=35, fg="gray")
        self.entry_email.grid(row=1, column=1, padx=5)
        self.entry_email.insert(0, "请输入邮箱")
        self.entry_email.bind("<FocusIn>", lambda e: self._on_entry_focus_in(self.entry_email, "请输入邮箱"))
        self.entry_email.bind("<FocusOut>", lambda e: self._on_entry_focus_out(self.entry_email, "请输入邮箱"))

        # 输入区
        input_frame = tk.LabelFrame(root, text="粘贴开单指令 (格式: 起点-终点 *数量)", padx=10, pady=10)
        input_frame.pack(fill="both", expand=True, padx=10, pady=5)
        self.txt_input = scrolledtext.ScrolledText(input_frame, height=10)
        self.txt_input.pack(fill="both", expand=True)
        self.txt_input.insert(tk.END, "EWR936-EWR600 *2\nEWR936-JFK *1\nNJ936-RIC-ORF *2")

        # 按钮
        btn_frame = tk.Frame(root, pady=10)
        btn_frame.pack()
        self.btn_start = tk.Button(btn_frame, text="开始自动开单", bg="#007bff", fg="white", font=("Arial", 12, "bold"), command=self.start_thread, height=2, width=20)
        self.btn_start.pack()

        # 日志区
        log_frame = tk.LabelFrame(root, text="执行日志", padx=10, pady=10)
        log_frame.pack(fill="both", expand=True, padx=10, pady=5)
        self.txt_log = scrolledtext.ScrolledText(log_frame, height=12, state='disabled', bg="#f4f4f4")
        self.txt_log.pack(fill="both", expand=True)

    def _on_entry_focus_in(self, entry, placeholder):
        """当输入框获得焦点时，如果是占位符文本则清空"""
        if entry.get() == placeholder:
            entry.delete(0, tk.END)
            entry.config(fg="black")

    def _on_entry_focus_out(self, entry, placeholder):
        """当输入框失去焦点时，如果为空则显示占位符"""
        if entry.get().strip() == "":
            entry.insert(0, placeholder)
            entry.config(fg="gray")

    def log(self, msg):
        self.txt_log.config(state='normal')
        self.txt_log.insert(tk.END, f"[{datetime.now().strftime('%H:%M:%S')}] {msg}\n")
        self.txt_log.see(tk.END)
        self.txt_log.config(state='disabled')

    def start_thread(self):
        batch_no = self.entry_batch.get().strip()
        email = self.entry_email.get().strip()
        raw_commands = self.txt_input.get("1.0", tk.END).strip()
        
        # 检查是否是占位符文本，如果是则视为空
        if batch_no == "请输入当日批次号":
            batch_no = ""
        if email == "请输入邮箱":
            email = ""
        
        if not raw_commands:
            messagebox.showwarning("提示", "请先输入指令")
            return

        self.btn_start.config(state="disabled", text="运行中...")
        threading.Thread(target=self.run_automation, args=(batch_no, email, raw_commands), daemon=True).start()

    def run_automation(self, batch_no, email, raw_commands):
        driver = None
        try:
            self.log("🚀 正在启动浏览器...")
            options = webdriver.ChromeOptions()
            driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
            
            # 1. 解析任务
            tasks = []
            lines = raw_commands.split('\n')
            for line in lines:
                if not line.strip(): continue
                try:
                    # 处理数量 *N
                    count = 1
                    if "*" in line:
                        parts = line.split("*")
                        line = parts[0]
                        count = int(parts[1].strip())
                    
                    # 处理 EWR936-LAX 或 NJ936-RIC-ORF
                    if "-" in line:
                        route_parts = line.split("-")
                        origin = route_parts[0].strip()
                        
                        # 标准化别名 (origin normalization)
                        origin_aliases = {"NJ936": "EWR936", "NJ600": "EWR600"}
                        origin = origin_aliases.get(origin.upper(), origin)
                        
                        # 判断是两站还是三站路线
                        if len(route_parts) == 2:
                            # 两站路线: Origin -> Final Stop
                            dest_key = route_parts[1].strip()
                            
                            # 标准化目的地别名 (destination normalization)
                            dest_aliases = {"NJ936": "EWR936", "NJ600": "EWR600"}
                            dest_key = dest_aliases.get(dest_key.upper(), dest_key)
                            
                            # 映射地址
                            full_address = ADDRESS_MAP.get(dest_key, dest_key) # 找不到就用原值
                            carrier = get_carrier(dest_key)
                            pallets = get_pallet_count(dest_key)
                            
                            for _ in range(count):
                                tasks.append({
                                    "bol_type": "two_stop",
                                    "origin": origin,
                                    "final_stop": full_address,
                                    "carrier": carrier,
                                    "pallets": str(pallets)
                                })
                        
                        elif len(route_parts) == 3:
                            # 三站路线: Origin -> Stop 1 -> Final Stop
                            stop1_key = route_parts[1].strip()
                            dest_key = route_parts[2].strip()
                            
                            # 标准化别名
                            dest_aliases = {"NJ936": "EWR936", "NJ600": "EWR600"}
                            stop1_key = dest_aliases.get(stop1_key.upper(), stop1_key)
                            dest_key = dest_aliases.get(dest_key.upper(), dest_key)
                            
                            # 映射地址
                            stop1_address = ADDRESS_MAP.get(stop1_key, stop1_key)
                            final_stop_address = ADDRESS_MAP.get(dest_key, dest_key)
                            
                            # 使用最终目的地的carrier
                            carrier = get_carrier(dest_key)
                            
                            for _ in range(count):
                                tasks.append({
                                    "bol_type": "three_stop",
                                    "origin": origin,
                                    "stop1": stop1_address,
                                    "final_stop": final_stop_address,
                                    "carrier": carrier,
                                    "stop1_pallets": "12",
                                    "stop1_pieces": "0",
                                    "stop1_volume": "10000",
                                    "final_pallets": "12",
                                    "final_pieces": "0",
                                    "final_volume": "10000"
                                })
                        else:
                            self.log(f"⚠️ 不支持的路由格式: {line} (需要2或3个站点)")
                            continue
                    else:
                        self.log(f"⚠️ 跳过无效行: {line}")
                except Exception as e:
                    self.log(f"❌ 解析错误: {line} ({e})")

            total = len(tasks)
            self.log(f"✅ 解析完成，共 {total} 张单据待生成。")

            # 2. 执行循环
            for i, task in enumerate(tasks, 1):
                if task.get('bol_type') == 'three_stop':
                    route_desc = f"{task['origin']} -> {task['stop1'][:15]} -> {task['final_stop'][:15]}"
                else:
                    route_desc = f"{task['origin']} -> {task['final_stop'][:15]}"
                self.log(f"正在填写第 {i}/{total} 张: {route_desc}...")
                self.fill_smartsheet(driver, task, batch_no, email)
                self.log(f"🎉 第 {i} 张提交成功！")
                time.sleep(2) # 稍微等待，避免过快

            self.log("🏁 所有任务执行完毕！")
            messagebox.showinfo("完成", "所有 BOL 已生成完毕！")

        except Exception as e:
            self.log(f"❌ 严重错误: {e}")
            messagebox.showerror("错误", str(e))
        finally:
            if driver: driver.quit()
            self.root.after(0, lambda: self.btn_start.config(state="normal", text="开始自动开单"))

    # ================= 4. 核心填表逻辑 (增强修复版) =================
    def fill_smartsheet(self, driver, data, batch_no, email):
        url = "https://app.smartsheet.com/b/form/a2a520ba7d614e88a00d211941d13364"
        driver.get(url)
        wait = WebDriverWait(driver, 20)

        # 定义一个超级填空函数 (同时支持 input 和 textarea)
        def set_field(label_keyword, value, is_dropdown=False, is_date=False):
            target_elem = None
            
            # --- 阶段 1: 定位元素 ---
            
            # 策略A: 优先尝试 aria-label (查找 input 或 textarea)
            try:
                # 这里的 XPath 意思是：查找所有 aria-label 包含关键字的 input 或 textarea 元素
                xpath_a = f"//*[(self::input or self::textarea) and contains(@aria-label, '{label_keyword}')]"
                target_elem = driver.find_element(By.XPATH, xpath_a)
            except:
                pass

            # 策略B: 如果A失败，通过可视 Label 查找紧邻的输入框
            if not target_elem:
                try:
                    # 查找 Label -> 找它后面紧跟着的 input 或 textarea
                    xpath_b = f"//label[contains(., '{label_keyword}')]/following::*[self::input or self::textarea][1]"
                    target_elem = driver.find_element(By.XPATH, xpath_b)
                except:
                    pass

            if not target_elem:
                print(f"❌ 无法定位字段: {label_keyword}")
                # 不抛出致命错误，而是尝试继续，防止一张单卡死整个程序
                return 

            # --- 阶段 2: 交互操作 ---
            
            # 滚动到可见位置
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", target_elem)
            time.sleep(0.3)

            try:
                if is_dropdown:
                    # 下拉框处理
                    driver.execute_script("arguments[0].click();", target_elem)
                    target_elem.send_keys(Keys.CONTROL + "a")
                    target_elem.send_keys(Keys.DELETE)
                    target_elem.send_keys(str(value))
                    time.sleep(1.0) 
                    target_elem.send_keys(Keys.ENTER)
                    target_elem.send_keys(Keys.TAB)
                
                elif is_date:
                    # 日期处理
                    driver.execute_script("arguments[0].click();", target_elem)
                    target_elem.send_keys(Keys.CONTROL + "a") 
                    target_elem.send_keys(Keys.DELETE)
                    target_elem.send_keys(str(value))
                    target_elem.send_keys(Keys.TAB)
                    
                else:
                    # 普通文本 / 多行文本 (Batch#s)
                    try:
                        target_elem.click()
                        target_elem.clear()
                        target_elem.send_keys(str(value))
                    except Exception as click_err:
                        # 如果点击报错 (invalid element state)，直接用 JS 强制赋值
                        # 这是解决 Batch#s 报错的终极方案
                        print(f"⚠️ 字段 {label_keyword} 无法点击，尝试 JS 强制写入...")
                        driver.execute_script("arguments[0].value = arguments[1];", target_elem, str(value))
                        # 触发一下 change 事件，确保系统识别到值变了
                        driver.execute_script("arguments[0].dispatchEvent(new Event('change'));", target_elem)

            except Exception as e:
                print(f"⚠️ 填写 {label_keyword} 失败: {e}")

        # --- 开始按顺序填写 ---
        print("正在填表...")
        
        # 1. Mode
        set_field("Mode", "Ground", is_dropdown=True)
        
        # 2. BOL Type (根据路由类型选择)
        bol_type = data.get('bol_type', 'two_stop')
        if bol_type == 'three_stop':
            set_field("BOL Type", "Origin - Stop1 - Final Stop", is_dropdown=True)
        else:
            set_field("BOL Type", "Origin -> Final Stop", is_dropdown=True)
        
        # 3. Ship Date (美式格式)
        today_date = datetime.now().strftime("%m/%d/%Y")
        set_field("Ship Date", today_date, is_date=True)
        
        # 4. Email
        set_field("Email address", email)
        
        # 5. Origin
        set_field("Origin", data['origin'], is_dropdown=True)
        
        # 6-9. 根据路由类型填写不同字段
        if bol_type == 'three_stop':
            # 三站路线: Stop 1 和 Final Stop
            # 6. Stop1
            set_field("Stop1", data['stop1'], is_dropdown=True)
            time.sleep(0.5)  # 等待下拉框选择完成
            
            # 7. Stop1 PALLET Count
            set_field("Stop1 PALLET Count", data['stop1_pallets'])
            
            # 8. Stop1 PIECE Count
            set_field("Stop1 PIECE Count", data['stop1_pieces'])
            
            # 9. Stop1 Volume Weight
            set_field("Stop1 Volume Weight", data['stop1_volume'])
            
            # 10. Final Stop
            set_field("Final Stop", data['final_stop'], is_dropdown=True)
            time.sleep(0.5)  # 等待下拉框选择完成
            
            # 11. Final Stop Total PALLET Count
            set_field("Final Stop Total PALLET Count", data['final_pallets'])
            
            # 12. Final Stop Total PIECE Count
            set_field("Final Stop Total PIECE Count", data['final_pieces'])
            
            # 13. Final Stop Volume Weight (必须找到 Final Stop 的 Volume Weight，不能是 Stop1 的)
            # 策略：先找到 "Final Stop Total PIECE Count" 字段，然后找它后面的 Volume Weight 字段
            volume_set = False
            try:
                # 先定位到 "Final Stop Total PIECE Count" 字段
                piece_count_elem = None
                try:
                    xpath_piece = f"//*[(self::input or self::textarea) and contains(@aria-label, 'Final Stop Total PIECE Count')]"
                    piece_count_elem = driver.find_element(By.XPATH, xpath_piece)
                except:
                    try:
                        xpath_piece = f"//label[contains(., 'Final Stop Total PIECE Count')]/following::*[self::input or self::textarea][1]"
                        piece_count_elem = driver.find_element(By.XPATH, xpath_piece)
                    except:
                        pass
                
                if piece_count_elem:
                    # 在 "Final Stop Total PIECE Count" 之后查找 Volume Weight 字段
                    # 排除包含 "Stop1" 的字段
                    xpath_volume = f"./following::*[(self::input or self::textarea) and (contains(@aria-label, 'Volume Weight') or contains(@aria-label, 'Volume')) and not(contains(@aria-label, 'Stop1'))][1]"
                    try:
                        target_elem = piece_count_elem.find_element(By.XPATH, xpath_volume)
                        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", target_elem)
                        time.sleep(0.3)
                        try:
                            target_elem.click()
                            target_elem.clear()
                            target_elem.send_keys(data['final_volume'])
                        except:
                            driver.execute_script("arguments[0].value = arguments[1];", target_elem, data['final_volume'])
                            driver.execute_script("arguments[0].dispatchEvent(new Event('change'));", target_elem)
                        volume_set = True
                        print(f"✅ 成功填写 Final Stop Volume Weight 字段")
                    except:
                        # 如果上面的方法失败，尝试通过 label 查找
                        try:
                            xpath_volume_label = f"./following::label[contains(., 'Volume Weight') and not(contains(., 'Stop1'))]/following::*[self::input or self::textarea][1]"
                            target_elem = piece_count_elem.find_element(By.XPATH, xpath_volume_label)
                            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", target_elem)
                            time.sleep(0.3)
                            try:
                                target_elem.click()
                                target_elem.clear()
                                target_elem.send_keys(data['final_volume'])
                            except:
                                driver.execute_script("arguments[0].value = arguments[1];", target_elem, data['final_volume'])
                                driver.execute_script("arguments[0].dispatchEvent(new Event('change'));", target_elem)
                            volume_set = True
                            print(f"✅ 成功填写 Final Stop Volume Weight 字段 (通过 label)")
                        except:
                            pass
            except Exception as e:
                print(f"⚠️ 查找 Final Stop Volume Weight 时出错: {e}")
            
            # 如果上面的方法都失败，尝试直接查找（但排除 Stop1）
            if not volume_set:
                try:
                    # 查找所有 Volume Weight 字段，但排除 Stop1 的
                    xpath_all = f"//*[(self::input or self::textarea) and (contains(@aria-label, 'Volume Weight') or contains(@aria-label, 'Volume')) and not(contains(@aria-label, 'Stop1'))]"
                    all_volume_fields = driver.find_elements(By.XPATH, xpath_all)
                    # 找到最后一个（应该是 Final Stop 的）
                    if all_volume_fields:
                        target_elem = all_volume_fields[-1]
                        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", target_elem)
                        time.sleep(0.3)
                        try:
                            target_elem.click()
                            target_elem.clear()
                            target_elem.send_keys(data['final_volume'])
                        except:
                            driver.execute_script("arguments[0].value = arguments[1];", target_elem, data['final_volume'])
                            driver.execute_script("arguments[0].dispatchEvent(new Event('change'));", target_elem)
                        volume_set = True
                        print(f"✅ 成功填写 Final Stop Volume Weight 字段 (使用最后一个匹配项)")
                except:
                    pass
            
            if not volume_set:
                print(f"⚠️ 无法定位 Final Stop Volume Weight 字段")
            
            # 14. Carrier
            set_field("Carrier", data['carrier'], is_dropdown=True)
        else:
            # 两站路线: 原有逻辑
            # 6. Final Stop
            set_field("Final Stop", data['final_stop'], is_dropdown=True)
            
            # 7. Pallets
            set_field("PALLET", data['pallets'])
            
            # 8. Pieces
            set_field("PIECE", "0")
            
            # 9. Volume
            set_field("Volume", "10000")
            
            # 10. Carrier
            set_field("Carrier", data['carrier'], is_dropdown=True)
        
        # 11. Batch (专门修复：支持 Textarea)
        set_field("Batch", batch_no)
        
        # 12. Cross-dock
        try:
            set_field("cross-dock", "No", is_dropdown=True)
        except:
            # 备用方案：暴力点击 No
            try:
                no_span = driver.find_element(By.XPATH, "//span[text()='No']")
                driver.execute_script("arguments[0].click();", no_span)
            except:
                pass

        # --- 提交 ---
        print("正在提交...")
        try:
            submit_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@data-client-id='form_submit_btn']")))
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", submit_btn)
            time.sleep(0.5)
            submit_btn.click()
            
            # 等待成功
            wait.until(EC.presence_of_element_located((By.XPATH, "//*[contains(text(), 'Thank you') or contains(text(), 'Success')]")))
        except Exception as e:
            print(f"提交步骤出错: {e}")
            # 如果提交出错，不要关闭浏览器，让用户能手动点一下
            pass


if __name__ == "__main__":
    root = tk.Tk()
    app = BOLAgentApp(root)
    root.mainloop()
