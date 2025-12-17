import tkinter as tk
from tkinter import scrolledtext, messagebox
import threading
import time
from datetime import datetime, timedelta
import math
from concurrent.futures import ThreadPoolExecutor

# Selenium 库
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys

# ================= 1. 地址大字典 =================
ADDRESS_MAP = {
    # --- 常用大仓 (带星号格式) ---
    "LAX": "*LAX162*: 16288 Boyle Ave, Fontana CA 92337",
    "LAX162": "*LAX162*: 16288 Boyle Ave, Fontana CA 92337",
    "EWR": "*EWR600*: 600 Federal Blvd, Carteret NJ 07008",
    "EWR600": "*EWR600*: 600 Federal Blvd, Carteret NJ 07008",
    "NJ600": "*EWR600*: 600 Federal Blvd, Carteret NJ 07008",
    "JFK": "*JFK175*: 175-14 147th Ave, Jamaica NY 11434",
    "JFK175": "*JFK175*: 175-14 147th Ave, Jamaica NY 11434",
    "ORD126": "ORD126: 12690 S Rte 59, Plainfield IL 60585",
    "ORD121": "*ORD121*: 1211 Tower Road, Schaumburg IL 60173",
    "DFW140": "DFW140: 1401 Dunn Dr, Carrolton TX 75006",
    "DFW445": "*DFW445*: 4450 W Walnut Hill Lane, Unit 100, Irving TX 75038",
    "ATL": "*ATL441*: 4411 Bibb Boulevard, Tucker GA 30084",
    "ATL441": "*ATL441*: 4411 Bibb Boulevard, Tucker GA 30084",
    "MIA": "*MIA307*: 3075 NW 107th Ave, Doral FL 33172",
    "MIA307": "*MIA307*: 3075 NW 107th Ave, Doral FL 33172",
    "IAH": "IAH879: 8790 Wallisville Rd, Houston TX 77029",

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
    "PIT": "PIT017: 17 Herron Ave, Emsworth PA 15202"
}

# ================= 2. 业务规则逻辑 =================

def get_carrier(destination_key):
    dest = destination_key.upper()
    if "EWR" in dest or "JFK" in dest: return "Han Express"
    if "ATL" in dest or "MIA" in dest: return "NYQZ"
    if any(k in dest for k in ["ORD", "DFW", "BOS", "PHL", "DCA", "RDU", "HFD", "ORF", "DOV", "PVD", "WHS", "RIC", "IAH"]):
        return "80s Express"
    return "Spot Freight"

def get_pallet_count(destination_key):
    dest = destination_key.upper()
    short_haul = ["EWR", "JFK", "NJ25", "PHL", "DCA", "BOS", "HFD", "PVD", "DOV"]
    if any(k in dest for k in short_haul):
        return 12
    return 26

# ================= 3. GUI 主程序 =================

class BOLAgentApp:
    def __init__(self, root):
        self.root = root
        self.root.title("UniUni BOL 自动开单机器人 (通用版)")
        self.root.geometry("750x850")

        # 配置区
        config_frame = tk.LabelFrame(root, text="基础配置", padx=10, pady=10)
        config_frame.pack(fill="x", padx=10, pady=5)

        # 1. Batch Number (带占位符)
        tk.Label(config_frame, text="Batch Number:").grid(row=0, column=0, sticky="w")
        self.entry_batch = tk.Entry(config_frame, width=35, fg="gray")
        self.entry_batch.grid(row=0, column=1, padx=5, pady=2, sticky="w")
        self.entry_batch.insert(0, "请输入当日批次号")
        self.entry_batch.bind("<FocusIn>", lambda e: self._on_entry_focus_in(self.entry_batch, "请输入当日批次号"))
        self.entry_batch.bind("<FocusOut>", lambda e: self._on_entry_focus_out(self.entry_batch, "请输入当日批次号"))

        # 2. Email (带占位符)
        tk.Label(config_frame, text="Email:").grid(row=1, column=0, sticky="w")
        self.entry_email = tk.Entry(config_frame, width=35, fg="gray")
        self.entry_email.grid(row=1, column=1, padx=5, pady=2, sticky="w")
        self.entry_email.insert(0, "请输入邮箱")
        self.entry_email.bind("<FocusIn>", lambda e: self._on_entry_focus_in(self.entry_email, "请输入邮箱"))
        self.entry_email.bind("<FocusOut>", lambda e: self._on_entry_focus_out(self.entry_email, "请输入邮箱"))

        # 3. Ship Date (日期选择)
        tk.Label(config_frame, text="Ship Date:").grid(row=2, column=0, sticky="w")
        date_frame = tk.Frame(config_frame)
        date_frame.grid(row=2, column=1, sticky="w", padx=5, pady=2)
        
        self.entry_date = tk.Entry(date_frame, width=15)
        self.entry_date.pack(side="left")
        self.entry_date.insert(0, datetime.now().strftime("%m/%d/%Y")) # 默认今天

        tk.Button(date_frame, text="昨天", command=lambda: self.set_date(-1), font=("Arial", 9)).pack(side="left", padx=2)
        tk.Button(date_frame, text="今天", command=lambda: self.set_date(0), font=("Arial", 9)).pack(side="left", padx=2)
        tk.Button(date_frame, text="明天", command=lambda: self.set_date(1), font=("Arial", 9)).pack(side="left", padx=2)

        # 4. 并发设置
        tk.Label(config_frame, text="并发窗口数:").grid(row=3, column=0, sticky="w")
        self.entry_workers = tk.Entry(config_frame, width=10)
        self.entry_workers.grid(row=3, column=1, sticky="w", padx=5, pady=2)
        self.entry_workers.insert(0, "3")
        tk.Label(config_frame, text="(注意：多开太卡，建议3-5个)", fg="gray").grid(row=3, column=1, sticky="e", padx=50)

        # 输入区
        input_frame = tk.LabelFrame(root, text="指令区 (格式: 起点-终点 *数量)", padx=10, pady=10)
        input_frame.pack(fill="both", expand=True, padx=10, pady=5)
        self.txt_input = scrolledtext.ScrolledText(input_frame, height=10)
        self.txt_input.pack(fill="both", expand=True)
        self.txt_input.insert(tk.END, "EWR936-EWR600 *2\nEWR936-JFK *1")

        # 按钮
        btn_frame = tk.Frame(root, pady=10)
        btn_frame.pack()
        self.btn_start = tk.Button(btn_frame, text="🚀 启动多窗口并发开单", bg="#007bff", fg="white", font=("Arial", 14, "bold"), command=self.start_thread, height=2, width=25)
        self.btn_start.pack()

        # 日志区
        log_frame = tk.LabelFrame(root, text="执行日志", padx=10, pady=10)
        log_frame.pack(fill="both", expand=True, padx=10, pady=5)
        self.txt_log = scrolledtext.ScrolledText(log_frame, height=12, state='disabled', bg="#f4f4f4")
        self.txt_log.pack(fill="both", expand=True)

    # --- 占位符处理逻辑 ---
    def _on_entry_focus_in(self, entry, placeholder):
        if entry.get() == placeholder:
            entry.delete(0, tk.END)
            entry.config(fg="black")

    def _on_entry_focus_out(self, entry, placeholder):
        if entry.get().strip() == "":
            entry.insert(0, placeholder)
            entry.config(fg="gray")

    def set_date(self, delta_days):
        target_date = datetime.now() + timedelta(days=delta_days)
        formatted = target_date.strftime("%m/%d/%Y")
        self.entry_date.delete(0, tk.END)
        self.entry_date.insert(0, formatted)

    def log(self, msg):
        self.txt_log.config(state='normal')
        self.txt_log.insert(tk.END, f"[{datetime.now().strftime('%H:%M:%S')}] {msg}\n")
        self.txt_log.see(tk.END)
        self.txt_log.config(state='disabled')

    def start_thread(self):
        batch_no = self.entry_batch.get().strip()
        email = self.entry_email.get().strip()
        ship_date = self.entry_date.get().strip()
        raw_commands = self.txt_input.get("1.0", tk.END).strip()
        try:
            max_workers = int(self.entry_workers.get().strip())
        except:
            max_workers = 3

        # 校验输入
        if batch_no == "请输入当日批次号" or not batch_no:
            messagebox.showwarning("提示", "请填写 Batch Number")
            return
        if email == "请输入邮箱" or not email:
            messagebox.showwarning("提示", "请填写 Email")
            return
        if not raw_commands:
            messagebox.showwarning("提示", "请先输入指令")
            return
        if not ship_date:
            messagebox.showwarning("提示", "日期不能为空")
            return

        self.btn_start.config(state="disabled", text="运行中...")
        threading.Thread(target=self.run_scheduler, args=(batch_no, email, ship_date, raw_commands, max_workers), daemon=True).start()

    # --- 核心调度 ---
    def run_scheduler(self, batch_no, email, ship_date, raw_commands, max_workers):
        try:
            all_tasks = self.parse_commands(raw_commands)
            total_tasks = len(all_tasks)
            
            if total_tasks == 0:
                self.log("⚠️ 没有生成有效任务，请检查输入。")
                return

            self.log(f"✅ 解析完成，共 {total_tasks} 张单据。")
            self.log(f"📅 发货日期: {ship_date}, 批次: {batch_no}, 邮箱: {email}")
            self.log(f"🔥 启动 {max_workers} 个窗口并发处理...")

            chunk_size = math.ceil(total_tasks / max_workers)
            task_chunks = [all_tasks[i:i + chunk_size] for i in range(0, total_tasks, chunk_size)]

            with ThreadPoolExecutor(max_workers=max_workers) as executor:
                futures = []
                for i, chunk in enumerate(task_chunks):
                    if not chunk: continue
                    worker_id = i + 1
                    futures.append(executor.submit(self.process_batch, chunk, batch_no, email, ship_date, worker_id))
            
            self.log("🏁🏁🏁 所有任务执行完毕！ 🏁🏁🏁")
            messagebox.showinfo("完成", "所有并发任务已完成！")

        except Exception as e:
            self.log(f"❌ 调度器严重错误: {e}")
        finally:
            self.root.after(0, lambda: self.btn_start.config(state="normal", text="🚀 启动多窗口并发开单"))

    def parse_commands(self, raw_commands):
        tasks = []
        lines = raw_commands.split('\n')
        for line in lines:
            if not line.strip(): continue
            try:
                count = 1
                if "*" in line:
                    parts = line.split("*")
                    line = parts[0]
                    count = int(parts[1].strip())
                
                if "-" in line:
                    route_parts = line.split("-")
                    origin = route_parts[0].strip()
                    origin_aliases = {"NJ936": "EWR936", "NJ600": "EWR600"}
                    origin = origin_aliases.get(origin.upper(), origin)
                    
                    if len(route_parts) == 2:
                        dest_key = route_parts[1].strip()
                        dest_aliases = {"NJ936": "EWR936", "NJ600": "EWR600"}
                        dest_key = dest_aliases.get(dest_key.upper(), dest_key)
                        full_address = ADDRESS_MAP.get(dest_key, dest_key)
                        carrier = get_carrier(dest_key)
                        pallets = get_pallet_count(dest_key)
                        
                        for _ in range(count):
                            tasks.append({
                                "bol_type": "two_stop", "origin": origin, "final_stop": full_address,
                                "carrier": carrier, "pallets": str(pallets)
                            })
                    
                    elif len(route_parts) == 3:
                        stop1_key = route_parts[1].strip()
                        dest_key = route_parts[2].strip()
                        dest_aliases = {"NJ936": "EWR936", "NJ600": "EWR600"}
                        stop1_key = dest_aliases.get(stop1_key.upper(), stop1_key)
                        dest_key = dest_aliases.get(dest_key.upper(), dest_key)
                        stop1_address = ADDRESS_MAP.get(stop1_key, stop1_key)
                        final_stop_address = ADDRESS_MAP.get(dest_key, dest_key)
                        carrier = get_carrier(dest_key)
                        
                        for _ in range(count):
                            tasks.append({
                                "bol_type": "three_stop", "origin": origin, "stop1": stop1_address,
                                "final_stop": final_stop_address, "carrier": carrier,
                                "stop1_pallets": "12", "stop1_pieces": "0", "stop1_volume": "10000",
                                "final_pallets": "12", "final_pieces": "0", "final_volume": "10000"
                            })
            except Exception as e:
                self.log(f"解析忽略: {line} ({e})")
        return tasks

    # --- Worker ---
    def process_batch(self, task_list, batch_no, email, ship_date, worker_id):
        self.log(f"🔧 工人 #{worker_id} 启动...")
        driver = None
        try:
            options = webdriver.ChromeOptions()
            driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
            driver.set_window_size(1200, 900)

            for i, task in enumerate(task_list, 1):
                desc = task['final_stop'][:10]
                self.log(f"▶️ 工人 #{worker_id}: 第 {i}/{len(task_list)} 张 ({desc})...")
                try:
                    self.fill_smartsheet(driver, task, batch_no, email, ship_date)
                    self.log(f"✅ 工人 #{worker_id}: 第 {i} 张完成！")
                    time.sleep(1.5)
                except Exception as e:
                    self.log(f"❌ 工人 #{worker_id}: 第 {i} 张失败: {e}")
                    continue 
            
            self.log(f"😴 工人 #{worker_id} 下班。")

        except Exception as e:
            self.log(f"💥 工人 #{worker_id} 崩溃: {e}")
        finally:
            if driver: driver.quit()

    # --- 填表动作 ---
    def fill_smartsheet(self, driver, data, batch_no, email, ship_date):
        url = "https://app.smartsheet.com/b/form/a2a520ba7d614e88a00d211941d13364"
        driver.get(url)
        wait = WebDriverWait(driver, 30)

        def set_field(label_keyword, value, is_dropdown=False, is_date=False):
            target_elem = None
            try:
                xpath_a = f"//*[(self::input or self::textarea) and contains(@aria-label, '{label_keyword}')]"
                target_elem = driver.find_element(By.XPATH, xpath_a)
            except: pass

            if not target_elem:
                try:
                    xpath_b = f"//label[contains(., '{label_keyword}')]/following::*[self::input or self::textarea][1]"
                    target_elem = driver.find_element(By.XPATH, xpath_b)
                except: pass

            if not target_elem:
                if "Volume" in label_keyword: return 
                print(f"找不到 {label_keyword}")
                return 

            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", target_elem)
            time.sleep(0.3)

            try:
                if is_dropdown:
                    driver.execute_script("arguments[0].click();", target_elem)
                    target_elem.send_keys(Keys.CONTROL + "a")
                    target_elem.send_keys(Keys.DELETE)
                    target_elem.send_keys(str(value))
                    time.sleep(1.2)
                    target_elem.send_keys(Keys.ENTER)
                    target_elem.send_keys(Keys.TAB)
                elif is_date:
                    driver.execute_script("arguments[0].click();", target_elem)
                    target_elem.send_keys(Keys.CONTROL + "a") 
                    target_elem.send_keys(Keys.DELETE)
                    target_elem.send_keys(str(value))
                    target_elem.send_keys(Keys.TAB)
                else:
                    try:
                        target_elem.click()
                        target_elem.clear()
                        target_elem.send_keys(str(value))
                    except:
                        driver.execute_script("arguments[0].value = arguments[1];", target_elem, str(value))
                        driver.execute_script("arguments[0].dispatchEvent(new Event('change'));", target_elem)
            except Exception as e:
                print(f"填值 {label_keyword} 出错: {e}")

        # 填表流程
        set_field("Mode", "Ground", is_dropdown=True)
        bol_type = data.get('bol_type', 'two_stop')
        if bol_type == 'three_stop':
            set_field("BOL Type", "Origin -> Stop1 -> Final Stop", is_dropdown=True)
        else:
            set_field("BOL Type", "Origin -> Final Stop", is_dropdown=True)
        
        # 使用传入的 ship_date
        set_field("Ship Date", ship_date, is_date=True)
        set_field("Email address", email)
        set_field("Origin", data['origin'], is_dropdown=True)

        if bol_type == 'three_stop':
            set_field("Stop1", data['stop1'], is_dropdown=True)
            time.sleep(0.5)
            set_field("Stop1 PALLET Count", data['stop1_pallets'])
            set_field("Stop1 PIECE Count", data['stop1_pieces'])
            set_field("Stop1 Volume Weight", data['stop1_volume'])
            
            set_field("Final Stop", data['final_stop'], is_dropdown=True)
            time.sleep(0.5)
            set_field("Final Stop Total PALLET Count", data['final_pallets'])
            set_field("Final Stop Total PIECE Count", data['final_pieces'])
            set_field("Volume Weight", data.get('final_volume', '10000'))
            
            try:
                all_vol_inputs = driver.find_elements(By.XPATH, "//*[contains(@aria-label, 'Volume') or contains(@placeholder, 'Volume')]")
                if len(all_vol_inputs) >= 2:
                    target_vol_element = all_vol_inputs[-1]
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", target_vol_element)
                    target_vol_element.clear()
                    target_vol_element.send_keys(data.get('final_volume', '10000'))

                    self.log(f"   [Worker#{driver.name}] 成功填写 Final Volume。")
                else:
                    self.log(f"   [Worker#{driver.name}] ⚠️ Volume Weight 元素定位失败，跳过。")
            except Exception as e:
                self.log(f"   [Worker#{driver.name}] ❌ 填写 Final Volume 兜底失败: {e}")
            
            set_field("Carrier", data['carrier'], is_dropdown=True)
        else:
            set_field("Final Stop", data['final_stop'], is_dropdown=True)
            set_field("PALLET", data['pallets'])
            set_field("PIECE", "0")
            set_field("Volume", "10000")
            set_field("Carrier", data['carrier'], is_dropdown=True)

        set_field("Batch", batch_no)
        try: set_field("cross-dock", "No", is_dropdown=True)
        except: pass

        try:
            submit_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@data-client-id='form_submit_btn']")))
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", submit_btn)
            time.sleep(0.5)
            submit_btn.click()
            wait.until(EC.presence_of_element_located((By.XPATH, "//*[contains(text(), 'Thank you') or contains(text(), 'Success')]")))
        except: pass

if __name__ == "__main__":
    root = tk.Tk()
    app = BOLAgentApp(root)
    root.mainloop()