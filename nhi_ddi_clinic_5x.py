import ctypes
from ctypes import wintypes
import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
import json
import os
import re
import struct
import sys
import threading
import time
import tkinter as tk
from tkinter import messagebox, scrolledtext, ttk
import requests
import urllib3
import winsound

try:
    import pythoncom
    import win32com.client

    HAS_WIN32COM = True
except ImportError:
    HAS_WIN32COM = False

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# =============================================================================
# 健保 5.1.5.7 元件路徑、本機連動埠與常數設定
# =============================================================================
CSHIS_LIB_DIR = r"C:\NHI\LIB"
CSHIS_DLL_PATH = r"C:\NHI\LIB\CSHIS.dll"
CSHISX_PROG_ID = "CsHisX.nhicshisx.1"

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.path.join(BASE_DIR, "config_clinic_drugs.json")
NHI_CLOUD_API_URL = "https://medcloudws2.nhi.gov.tw/api/imie5000/GetMedPrtData"

LOCAL_BRIDGE_HOST = "127.0.0.1"
LOCAL_BRIDGE_PORT = 18765
AUTO_DEDUP_SECONDS = 300  # 同病人 5 分鐘內不重複自動觸發

# Windows WinSCard 常數
SCARD_SCOPE_USER = 0
SCARD_S_SUCCESS = 0
SCARD_E_TIMEOUT = 0x8010000A
SCARD_E_NO_READERS_AVAILABLE = 0x8010002E
SCARD_STATE_UNAWARE = 0x00000000
SCARD_STATE_CHANGED = 0x00000002
SCARD_STATE_PRESENT = 0x00000020
SCARD_STATE_MUTE = 0x00000200

# Windows Global Hotkey 常數 (Alt + F9，不搶展望 F1~F12)
MOD_ALT = 0x0001
MOD_NOREPEAT = 0x4000
VK_F9 = 0x78
WM_HOTKEY = 0x0312
WM_QUIT = 0x0012

# 多組口服 NSAID 腎功能探測碼備援清單 (自動輪替記憶當期有效碼)
NSAID_PROBE_CANDIDATES = [
    "BC23177100",  # Celebrex 200mg
    "AB45725100",  # Lacoxa SR 400mg
    "AB336991G0",  # Voren 50mg
    "AC35651151",  # Etodolac SR
]

DEFAULT_CONFIG = {
    "hosp_id_fallback": "3xxxxxxxxx",
    "hca_id_fallback": "A000000000",
    "com_port_index": 0,
    "auto_slot_sense": True,
    "working_nsaid_probe": "BC23177100",
    "drugs": [
        {
            "name": "Sporanox 100mg 適撲諾 (Itraconazole)",
            "code": "BC22816100",
            "category": "08",
            "checked": True,
        },
        {
            "name": "Diflucan 150mg 泰復肯 (Fluconazole)",
            "code": "BC21676100",
            "category": "08",
            "checked": True,
        },
        {
            "name": "Lamisil 250mg 療黴舒 (Terbinafine)",
            "code": "B019320100",
            "category": "08",
            "checked": True,
        },
        {
            "name": "Celebrex 200mg 希樂葆 (Celecoxib)",
            "code": "BC23177100",
            "category": "08",
            "checked": True,
        },
        {
            "name": "Omeprazole (官方手冊範例碼)",
            "code": "BC26000245",
            "category": "08",
            "checked": False,
        },
    ],
}

BANNER_THEMES = {
    "IDLE": ("#ECEFF1", "#37474F", "#B0BEC5"),
    "CRITICAL": ("#C62828", "#FFFFFF", "#8E0000"),
    "HIGH": ("#E65100", "#FFFFFF", "#BF360C"),
    "MODERATE": ("#FBC02D", "#212121", "#F57F17"),
    "UNKNOWN": ("#FFF3CD", "#664D03", "#FFECB5"),
    "NORMAL": ("#2E7D32", "#FFFFFF", "#1B5E20"),
    "INFO": ("#00695C", "#FFFFFF", "#004D40"),
}


class SCARD_READERSTATEW(ctypes.Structure):
    _fields_ = [
        ("szReader", wintypes.LPCWSTR),
        ("pvUserData", ctypes.c_void_p),
        ("dwCurrentState", wintypes.DWORD),
        ("dwEventState", wintypes.DWORD),
        ("cbAtr", wintypes.DWORD),
        ("rgbAtr", wintypes.c_ubyte * 36),
    ]


def ci_get(d, key, default=""):
    if not isinstance(d, dict):
        return default
    if key in d and d[key] is not None:
        return d[key]
    kl = key.lower()
    for k, v in d.items():
        if k.lower() == kl and v is not None:
            return v
    return default


def format_roc_bday_and_age(roc_bday_str):
    s = (roc_bday_str or "").strip()
    if len(s) != 7 or not s.isdigit():
        return s or "未知"
    try:
        roc_y, m, d = int(s[0:3]), int(s[3:5]), int(s[5:7])
        west_y = roc_y + 1911
        today = datetime.date.today()
        age = today.year - west_y - ((today.month, today.day) < (m, d))
        return f"民{roc_y}/{m:02d}/{d:02d} ({age}歲)"
    except Exception:
        return s


def is_compatible_masked_text(full_txt, masked_txt):
    """比對健保雲端網頁隱碼字串（如 A********* 或 王○明）是否與實體卡一致。"""
    a = (full_txt or "").strip().upper()
    b = (masked_txt or "").strip().upper()
    if not a or not b:
        return True
    if len(a) != len(b):
        return True
    masks = {"*", "○", "Ｏ", "O", "X", "?", "•", "＊"}
    for ca, cb in zip(a, b):
        if cb in masks or ca in masks:
            continue
        if ca != cb:
            return False
    return True


def wake_existing_instance_if_running():
    """若程式已在背景執行，喚醒既有視窗並結束重複開啟的程序。"""
    try:
        r = requests.post(
            f"http://{LOCAL_BRIDGE_HOST}:{LOCAL_BRIDGE_PORT}/show_window",
            json={"action": "show"},
            timeout=0.8,
        )
        if r.status_code == 200:
            return True
    except Exception:
        pass
    return False


class ClinicNHIDDIApp:

    def __init__(self, root):
        self.root = root
        self.py_bits = struct.calcsize("P") * 8
        self.root.title(
            f"診所通用版｜健保雲端【肝腎警示・過敏史・DDI 交互作用】5.x 插槽感應版 (Python {self.py_bits}-bit)"
        )
        self.root.geometry("1040x840")
        self._cfg_lock = threading.Lock()
        self.config = self.load_config()

        self.drug_vars = []
        self.editing_index = None
        self.is_querying = False
        self._dll_dir_handle = None
        self._stop_event = threading.Event()
        self._auto_after_id = None
        self._topmost_after_id = None
        self._hotkey_tid = None
        self.http_server = None

        # 當前病人狀態與防重複紀錄
        self.last_own_read_ts = 0.0
        self.last_auto_pat_id = ""
        self.last_auto_pat_ts = 0.0
        self.current_pat_id = ""
        self.current_pat_name = ""
        self.current_hcv_msgs = []
        self.current_renal_level = "IDLE"
        self.current_renal_title = ""
        self.current_renal_detail = ""
        self.current_web_labs = None
        self.banner_levels = {"renal": "IDLE", "liver": "IDLE", "allergy": "IDLE"}

        if hasattr(os, "add_dll_directory") and os.path.isdir(CSHIS_LIB_DIR):
            self._dll_dir_handle = os.add_dll_directory(CSHIS_LIB_DIR)

        self.create_widgets()
        self.refresh_drug_checkboxes()

        # 視窗內快捷鍵與關閉事件
        self.root.bind("<F9>", lambda e: self.on_manual_hotkey())
        self.root.bind("<Escape>", lambda e: self.root.iconify())
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

        # 啟動背景服務：1. 本機網頁外掛接收埠  2. WinSCard 插槽感應  3. Alt+F9 全域熱鍵
        self.start_local_bridge_server()
        self.start_scard_monitor_thread()
        self.start_global_hotkey_thread()

    # =========================================================================
    # 設定檔讀寫與舊版自動升級
    # =========================================================================
    def load_config(self):
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    cfg = json.load(f)
                for k, v in DEFAULT_CONFIG.items():
                    cfg.setdefault(k, v)
                drugs = cfg.get("drugs", [])
                if drugs and all("checked" not in d for d in drugs):
                    for d in drugs:
                        is_sample = ("範例" in d.get("name", "")) or ("測試" in d.get("name", ""))
                        d["checked"] = not is_sample
                    self.save_config(cfg)
                return cfg
            except Exception:
                pass
        self.save_config(DEFAULT_CONFIG)
        return json.loads(json.dumps(DEFAULT_CONFIG))

    def save_config(self, cfg=None):
        with self._cfg_lock:
            if cfg is None:
                cfg = self.config
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(cfg, f, ensure_ascii=False, indent=4)

    # =========================================================================
    # GUI 介面建置（含自適應換行與工具按鈕）
    # =========================================================================
    def create_widgets(self):
        frm_header = tk.Frame(self.root, padx=10, pady=8, bg="#1C2833")
        frm_header.pack(fill="x")

        frm_top_bar = tk.Frame(frm_header, bg="#1C2833")
        frm_top_bar.pack(fill="x", pady=(0, 6))

        bit_hint = (
            "🟢 32-bit 就緒｜插槽感應啟動中..."
            if self.py_bits == 32
            else "🟡 目前為 64-bit（僅支援測試模式；正式讀卡請用 py -3.13-32 啟動）"
        )
        self.lbl_card_bar = tk.Label(
            frm_top_bar,
            text=f"💳 狀態：{bit_hint}",
            font=("Microsoft JhengHei", 10, "bold"),
            bg="#1C2833",
            fg="#F8F9F9",
            anchor="w",
        )
        self.lbl_card_bar.pack(side="left", fill="x", expand=True)

        self.lbl_slot_badge = tk.Label(
            frm_top_bar,
            text="📡 讀卡槽偵測中",
            font=("Microsoft JhengHei", 9),
            bg="#2E4053",
            fg="#AED6F1",
            padx=8,
            pady=2,
        )
        self.lbl_slot_badge.pack(side="right", padx=(6, 0))

        self.frm_banners = tk.Frame(frm_header, bg="#1C2833")
        self.frm_banners.pack(fill="x")
        self.frm_banners.columnconfigure(0, weight=5)
        self.frm_banners.columnconfigure(1, weight=4)
        self.frm_banners.columnconfigure(2, weight=4)

        # 1. 腎功能色塊
        self.frm_renal = tk.Frame(
            self.frm_banners,
            bg="#ECEFF1",
            highlightbackground="#B0BEC5",
            highlightthickness=3,
            padx=10,
            pady=8,
        )
        self.frm_renal.grid(row=0, column=0, sticky="nsew", padx=(0, 4))
        self.lbl_renal_title = tk.Label(
            self.frm_renal,
            text="🫘 腎功能 (eGFR / CKD)：待查詢",
            font=("Microsoft JhengHei", 12, "bold"),
            bg="#ECEFF1",
            fg="#37474F",
            anchor="w",
            justify="left",
        )
        self.lbl_renal_title.pack(fill="x")
        self.lbl_renal_detail = tk.Label(
            self.frm_renal,
            text="推入健保卡自動比對近 12 個月雲端 eGFR、CKD 期別與洗腎註記",
            font=("Microsoft JhengHei", 9),
            bg="#ECEFF1",
            fg="#37474F",
            anchor="w",
            justify="left",
            wraplength=340,
        )
        self.lbl_renal_detail.pack(fill="x", pady=(4, 0))

        # 2. 肝功能與 C 肝色塊（支援雲端網頁 GOT/GPT 同步亮燈）
        self.frm_liver = tk.Frame(
            self.frm_banners,
            bg="#ECEFF1",
            highlightbackground="#B0BEC5",
            highlightthickness=3,
            padx=10,
            pady=8,
        )
        self.frm_liver.grid(row=0, column=1, sticky="nsew", padx=4)
        self.lbl_liver_title = tk.Label(
            self.frm_liver,
            text="🧪 肝功能與 C 肝：待查詢",
            font=("Microsoft JhengHei", 12, "bold"),
            bg="#ECEFF1",
            fg="#37474F",
            anchor="w",
            justify="left",
        )
        self.lbl_liver_title.pack(fill="x")
        self.lbl_liver_detail = tk.Label(
            self.frm_liver,
            text="自動查 C 肝陽性；點閱雲端網頁時自動同步 GOT / GPT 真實數值",
            font=("Microsoft JhengHei", 9),
            bg="#ECEFF1",
            fg="#37474F",
            anchor="w",
            justify="left",
            wraplength=280,
        )
        self.lbl_liver_detail.pack(fill="x", pady=(4, 0))

        # 3. 雲端藥物過敏色塊
        self.frm_allergy = tk.Frame(
            self.frm_banners,
            bg="#ECEFF1",
            highlightbackground="#B0BEC5",
            highlightthickness=3,
            padx=10,
            pady=8,
        )
        self.frm_allergy.grid(row=0, column=2, sticky="nsew", padx=(4, 0))
        self.lbl_allergy_title = tk.Label(
            self.frm_allergy,
            text="🚨 雲端藥物過敏：待查詢",
            font=("Microsoft JhengHei", 12, "bold"),
            bg="#ECEFF1",
            fg="#37474F",
            anchor="w",
            justify="left",
        )
        self.lbl_allergy_title.pack(fill="x")
        self.lbl_allergy_detail = tk.Label(
            self.frm_allergy,
            text="自動抓取雲端與跨院 IC 卡上傳之有效過敏藥物紀錄",
            font=("Microsoft JhengHei", 9),
            bg="#ECEFF1",
            fg="#37474F",
            anchor="w",
            justify="left",
            wraplength=280,
        )
        self.lbl_allergy_detail.pack(fill="x", pady=(4, 0))

        self.frm_banners.bind("<Configure>", self._on_banners_resize)

        # 1. 模式切換與工具列
        frm_mode = ttk.LabelFrame(
            self.root, text="1. 自動插槽感應與院所設定（全域熱鍵 Alt+F9 喚出重查｜Esc 縮小至工作列）", padding=8
        )
        frm_mode.pack(fill="x", padx=10, pady=4)

        frm_mode_row1 = ttk.Frame(frm_mode)
        frm_mode_row1.pack(fill="x")

        self.auto_sense_var = tk.BooleanVar(value=bool(self.config.get("auto_slot_sense", True)))
        ttk.Checkbutton(
            frm_mode_row1,
            text="啟用讀卡槽自動感應（健保卡推入 2 秒後背景檢查；全綠靜默、有異常自動置頂彈出）",
            variable=self.auto_sense_var,
            command=self._on_toggle_auto_sense,
        ).pack(side="left", padx=(2, 12))

        ttk.Button(
            frm_mode_row1, text="📌 建立桌面無黑框捷徑", command=self.create_desktop_shortcut
        ).pack(side="right", padx=2)
        ttk.Button(
            frm_mode_row1, text="📂 開啟藥品 JSON 設定", command=self.open_config_in_notepad
        ).pack(side="right", padx=4)

        frm_mode_row2 = ttk.Frame(frm_mode)
        frm_mode_row2.pack(fill="x", pady=(6, 0))

        self.test_mode_var = tk.BooleanVar(value=(self.py_bits != 32))
        ttk.Checkbutton(
            frm_mode_row2,
            text="官方測試模式 (測試個案 Z299999992，勾選時暫停插卡感應)",
            variable=self.test_mode_var,
        ).pack(side="left", padx=2)

        frm_fb = ttk.Frame(frm_mode_row2)
        frm_fb.pack(side="right", padx=2)
        ttk.Label(frm_fb, text="機構碼:").pack(side="left")
        self.ent_hosp = ttk.Entry(frm_fb, width=12)
        self.ent_hosp.insert(0, self.config.get("hosp_id_fallback", "3xxxxxxxxx"))
        self.ent_hosp.pack(side="left", padx=3)

        ttk.Label(frm_fb, text="醫師身分證號(測試必填):").pack(side="left")
        self.ent_hca = ttk.Entry(frm_fb, width=12)
        self.ent_hca.insert(0, self.config.get("hca_id_fallback", ""))
        self.ent_hca.pack(side="left", padx=3)

        # 2. 常用藥品選單區
        frm_drugs = ttk.LabelFrame(
            self.root,
            text="2. 診間常用處方藥品選單（勾選狀態自動記憶，插卡時自動比對雲端餘藥交互作用）",
            padding=8,
        )
        frm_drugs.pack(fill="x", padx=10, pady=4)

        frm_quick = ttk.Frame(frm_drugs)
        frm_quick.pack(fill="x", pady=(0, 4))
        ttk.Button(
            frm_quick, text="全選", width=6, command=lambda: self.set_all_checks(True)
        ).pack(side="left", padx=(0, 4))
        ttk.Button(
            frm_quick, text="清除勾選", width=9, command=lambda: self.set_all_checks(False)
        ).pack(side="left")
        ttk.Label(
            frm_quick,
            text="💡 建議平常保持勾選診所最常開立的 4～6 種口服抗黴菌藥／抗生素／止痛藥，插卡即自動完成背景把關。",
            foreground="#566573",
        ).pack(side="left", padx=10)

        self.frm_chk_list = ttk.Frame(frm_drugs)
        self.frm_chk_list.pack(fill="x", pady=2)

        frm_edit = ttk.Frame(frm_drugs)
        frm_edit.pack(fill="x", pady=(6, 0))
        ttk.Label(frm_edit, text="藥品名稱:").pack(side="left")
        self.ent_new_name = ttk.Entry(frm_edit, width=26)
        self.ent_new_name.pack(side="left", padx=4)
        ttk.Label(frm_edit, text="健保碼(10碼):").pack(side="left")
        self.ent_new_code = ttk.Entry(frm_edit, width=13)
        self.ent_new_code.pack(side="left", padx=4)
        ttk.Label(frm_edit, text="類別:").pack(side="left")
        self.cmb_cat = ttk.Combobox(
            frm_edit, values=["08:西藥", "09:中藥"], width=8, state="readonly"
        )
        self.cmb_cat.current(0)
        self.cmb_cat.pack(side="left", padx=4)

        self.btn_save_drug = ttk.Button(
            frm_edit, text="＋新增藥品", command=self.save_or_add_drug
        )
        self.btn_save_drug.pack(side="left", padx=3)
        ttk.Button(
            frm_edit, text="－刪除勾選藥品", command=self.delete_selected_drugs
        ).pack(side="left", padx=3)

        # 3. 手動立即查詢按鈕
        frm_action = ttk.Frame(self.root, padding=4)
        frm_action.pack(fill="x", padx=10)
        self.btn_check = tk.Button(
            frm_action,
            text="🔍 手動立即讀卡檢查（F9 / Alt+F9）➔ 檢查【肝腎・過敏・DDI 交互作用】",
            bg="#1565C0",
            fg="white",
            font=("Microsoft JhengHei", 12, "bold"),
            height=2,
            command=lambda: self.start_check_thread(auto_triggered=False),
        )
        self.btn_check.pack(fill="x")

        # 4. 結果回饋區
        frm_res = ttk.LabelFrame(
            self.root, text="3. 藥品交互作用 (DDI)、過敏史與網頁檢驗同步明細", padding=8
        )
        frm_res.pack(fill="both", expand=True, padx=10, pady=(2, 8))
        self.txt_result = scrolledtext.ScrolledText(
            frm_res, font=("Microsoft JhengHei", 10), wrap=tk.WORD
        )
        self.txt_result.pack(fill="both", expand=True)

    def _on_banners_resize(self, event):
        total_w = max(event.width, 600)
        self.lbl_renal_detail.configure(wraplength=max(int(total_w * 0.36) - 28, 180))
        self.lbl_liver_detail.configure(wraplength=max(int(total_w * 0.30) - 28, 150))
        self.lbl_allergy_detail.configure(wraplength=max(int(total_w * 0.30) - 28, 150))

    def _on_toggle_auto_sense(self):
        self.config["auto_slot_sense"] = bool(self.auto_sense_var.get())
        self.save_config()

    def set_banner(self, target, level, title_text, detail_text):
        self.banner_levels[target] = level
        bg, fg, border = BANNER_THEMES.get(level, BANNER_THEMES["IDLE"])
        mapping = {
            "renal": (self.frm_renal, self.lbl_renal_title, self.lbl_renal_detail),
            "liver": (self.frm_liver, self.lbl_liver_title, self.lbl_liver_detail),
            "allergy": (
                self.frm_allergy,
                self.lbl_allergy_title,
                self.lbl_allergy_detail,
            ),
        }
        frm, lbl_t, lbl_d = mapping[target]
        frm.configure(bg=bg, highlightbackground=border)
        lbl_t.configure(bg=bg, fg=fg, text=title_text)
        lbl_d.configure(bg=bg, fg=fg, text=detail_text)

    # =========================================================================
    # 視窗置頂彈出與工作列閃爍輔助
    # =========================================================================
    def popup_alert_window(self, play_sound=True):
        """偵測到異常時：發出溫和提示音並自動彈出置頂 4 秒。"""
        if play_sound:
            try:
                winsound.MessageBeep(winsound.MB_ICONEXCLAMATION)
            except Exception:
                pass
        try:
            self.root.deiconify()
            self.root.state("normal")
            self.root.lift()
            self.root.attributes("-topmost", True)
            if self._topmost_after_id:
                self.root.after_cancel(self._topmost_after_id)
            self._topmost_after_id = self.root.after(
                4000, lambda: self.root.attributes("-topmost", False)
            )
        except Exception:
            pass

    def flash_taskbar_icon(self):
        """在醫師瀏覽雲端網頁同步異常檢驗值時，閃爍工作列圖示提醒而不遮擋瀏覽器。"""
        try:
            frame_hex = self.root.wm_frame()
            hwnd = int(frame_hex, 16) if frame_hex else self.root.winfo_id()
            ctypes.windll.user32.FlashWindow(hwnd, True)
        except Exception:
            pass

    # =========================================================================
    # 一鍵建立桌面無黑框捷徑 & 開啟 JSON
    # =========================================================================
    def create_desktop_shortcut(self):
        if not HAS_WIN32COM:
            messagebox.showwarning("提示", "需要 pywin32 才能自動建立捷徑。")
            return
        try:
            py_dir = os.path.dirname(sys.executable)
            pythonw_path = os.path.join(py_dir, "pythonw.exe")
            target_exe = pythonw_path if os.path.exists(pythonw_path) else sys.executable
            script_path = os.path.abspath(__file__)

            shell = win32com.client.Dispatch("WScript.Shell")
            desktop = shell.SpecialFolders("Desktop")
            lnk_path = os.path.join(desktop, "健保雲端DDI與肝腎警示.lnk")
            shortcut = shell.CreateShortcut(lnk_path)
            shortcut.TargetPath = target_exe
            shortcut.Arguments = f'"{script_path}"'
            shortcut.WorkingDirectory = BASE_DIR
            shortcut.Description = "健保雲端肝腎警示・過敏史・DDI插槽感應助手"
            shortcut.Save()
            messagebox.showinfo(
                "建立成功",
                f"✅ 已在桌面建立無黑框捷徑：\n{lnk_path}\n\n"
                "早上開診雙擊一次即可縮小在背景自動把關！",
            )
        except Exception as e:
            messagebox.showerror("建立捷徑失敗", str(e))

    def open_config_in_notepad(self):
        self.save_config()
        try:
            os.startfile(CONFIG_FILE)
        except Exception as e:
            messagebox.showerror("開啟失敗", str(e))

    # =========================================================================
    # 藥品清單與勾選記憶管理
    # =========================================================================
    def refresh_drug_checkboxes(self):
        for w in self.frm_chk_list.winfo_children():
            w.destroy()
        self.drug_vars.clear()
        for idx, drug in enumerate(self.config["drugs"]):
            var = tk.BooleanVar(value=bool(drug.get("checked", True)))
            cat_tag = "中" if drug.get("category") == "09" else "西"
            row_frm = ttk.Frame(self.frm_chk_list)
            row_frm.grid(row=idx // 2, column=idx % 2, sticky="w", padx=8, pady=2)
            ttk.Checkbutton(
                row_frm,
                text=f"[{cat_tag}] {drug['name']} ({drug['code']})",
                variable=var,
                command=self._persist_drug_check_states,
            ).pack(side="left")
            ttk.Button(
                row_frm,
                text="編輯",
                width=4,
                command=lambda i=idx: self.load_drug_for_edit(i),
            ).pack(side="left", padx=4)
            self.drug_vars.append((var, drug))

    def _persist_drug_check_states(self):
        for var, drug in self.drug_vars:
            drug["checked"] = bool(var.get())
        self.last_auto_pat_id = ""
        self.save_config()

    def set_all_checks(self, state):
        for var, _ in self.drug_vars:
            var.set(state)
        self._persist_drug_check_states()

    def load_drug_for_edit(self, idx):
        d = self.config["drugs"][idx]
        self.editing_index = idx
        self.ent_new_name.delete(0, tk.END)
        self.ent_new_name.insert(0, d["name"])
        self.ent_new_code.delete(0, tk.END)
        self.ent_new_code.insert(0, d["code"])
        self.cmb_cat.current(1 if d.get("category") == "09" else 0)
        self.btn_save_drug.configure(text="💾 儲存修改")

    def save_or_add_drug(self):
        name = self.ent_new_name.get().strip()
        code = self.ent_new_code.get().strip().upper()
        cat = "09" if self.cmb_cat.get().startswith("09") else "08"
        if not name or not code:
            messagebox.showwarning("提示", "請輸入藥品名稱與 10 碼健保醫令代碼！")
            return
        if self.editing_index is not None:
            prev_chk = self.config["drugs"][self.editing_index].get("checked", True)
            self.config["drugs"][self.editing_index] = {
                "name": name,
                "code": code,
                "category": cat,
                "checked": prev_chk,
            }
            self.editing_index = None
            self.btn_save_drug.configure(text="＋新增藥品")
        else:
            self.config["drugs"].append(
                {"name": name, "code": code, "category": cat, "checked": True}
            )
        self.last_auto_pat_id = ""
        self.save_config()
        self.refresh_drug_checkboxes()
        self.ent_new_name.delete(0, tk.END)
        self.ent_new_code.delete(0, tk.END)

    def delete_selected_drugs(self):
        rem = [d for v, d in self.drug_vars if not v.get()]
        if len(rem) == len(self.config["drugs"]):
            messagebox.showinfo("提示", "請先打勾要刪除的藥品項目。")
            return
        self.editing_index = None
        self.btn_save_drug.configure(text="＋新增藥品")
        self.config["drugs"] = rem
        self.save_config()
        self.refresh_drug_checkboxes()

    # =========================================================================
    # 背景服務 1：Windows WinSCard 讀卡機插槽感應（0% 佔用 COM1）
    # =========================================================================
    def start_scard_monitor_thread(self):
        if self.py_bits != 32:
            self.lbl_slot_badge.configure(text="⚪ 64-bit 未啟用插槽讀卡")
            return
        threading.Thread(target=self._scard_monitor_loop, daemon=True).start()

    def _list_pcsc_readers(self, winscard, h_ctx):
        cch = wintypes.DWORD(0)
        ret = winscard.SCardListReadersW(h_ctx, None, None, ctypes.byref(cch)) & 0xFFFFFFFF
        if ret == SCARD_E_NO_READERS_AVAILABLE or ret != SCARD_S_SUCCESS or cch.value == 0:
            return []
        buf = ctypes.create_unicode_buffer(cch.value)
        ret = (
            winscard.SCardListReadersW(
                h_ctx, None, ctypes.cast(buf, ctypes.c_void_p), ctypes.byref(cch)
            )
            & 0xFFFFFFFF
        )
        if ret != SCARD_S_SUCCESS:
            return []
        return [r for r in buf[: cch.value].split("\x00") if r]

    def _scard_monitor_loop(self):
        try:
            winscard = ctypes.WinDLL("winscard.dll")
            winscard.SCardEstablishContext.argtypes = [
                wintypes.DWORD,
                ctypes.c_void_p,
                ctypes.c_void_p,
                ctypes.POINTER(ctypes.c_size_t),
            ]
            winscard.SCardEstablishContext.restype = ctypes.c_long
            winscard.SCardReleaseContext.argtypes = [ctypes.c_size_t]
            winscard.SCardReleaseContext.restype = ctypes.c_long
            winscard.SCardListReadersW.argtypes = [
                ctypes.c_size_t,
                wintypes.LPCWSTR,
                ctypes.c_void_p,
                ctypes.POINTER(wintypes.DWORD),
            ]
            winscard.SCardListReadersW.restype = ctypes.c_long
            winscard.SCardGetStatusChangeW.argtypes = [
                ctypes.c_size_t,
                wintypes.DWORD,
                ctypes.POINTER(SCARD_READERSTATEW),
                wintypes.DWORD,
            ]
            winscard.SCardGetStatusChangeW.restype = ctypes.c_long
        except Exception:
            self.root.after(
                0, lambda: self.lbl_slot_badge.configure(text="⚠️ 無法載入 WinSCard")
            )
            return

        while not self._stop_event.is_set():
            h_ctx = ctypes.c_size_t(0)
            ret = (
                winscard.SCardEstablishContext(
                    SCARD_SCOPE_USER, None, None, ctypes.byref(h_ctx)
                )
                & 0xFFFFFFFF
            )
            if ret != SCARD_S_SUCCESS:
                time.sleep(3.0)
                continue

            try:
                readers = self._list_pcsc_readers(winscard, h_ctx.value)
                if not readers:
                    self.root.after(
                        0, lambda: self.lbl_slot_badge.configure(text="⚪ 未偵測到讀卡機")
                    )
                    time.sleep(3.0)
                    continue

                n = len(readers)
                self.root.after(
                    0,
                    lambda cnt=n: self.lbl_slot_badge.configure(
                        text=f"📡 插槽感應待命 ({cnt} 槽｜0% 佔線)"
                    ),
                )

                states = (SCARD_READERSTATEW * n)()
                for i, rname in enumerate(readers):
                    states[i].szReader = rname
                    states[i].dwCurrentState = SCARD_STATE_UNAWARE

                ret = (
                    winscard.SCardGetStatusChangeW(h_ctx.value, 500, states, n)
                    & 0xFFFFFFFF
                )
                if ret not in (SCARD_S_SUCCESS, SCARD_E_TIMEOUT):
                    time.sleep(2.0)
                    continue

                usable = []
                for i in range(n):
                    ev = states[i].dwEventState
                    usable.append(
                        bool(ev & SCARD_STATE_PRESENT) and not bool(ev & SCARD_STATE_MUTE)
                    )
                    states[i].dwCurrentState = ev & ~SCARD_STATE_CHANGED

                idle_ticks = 0
                while not self._stop_event.is_set():
                    ret = (
                        winscard.SCardGetStatusChangeW(h_ctx.value, 1000, states, n)
                        & 0xFFFFFFFF
                    )
                    if ret == SCARD_E_TIMEOUT:
                        idle_ticks += 1
                        if idle_ticks >= 5:
                            idle_ticks = 0
                            cur_readers = self._list_pcsc_readers(winscard, h_ctx.value)
                            if cur_readers != readers:
                                break
                        continue

                    if ret != SCARD_S_SUCCESS:
                        break

                    for i in range(n):
                        ev = states[i].dwEventState
                        now_usable = bool(ev & SCARD_STATE_PRESENT) and not bool(
                            ev & SCARD_STATE_MUTE
                        )
                        was_usable = usable[i]
                        usable[i] = now_usable
                        states[i].dwCurrentState = ev & ~SCARD_STATE_CHANGED

                        if (not was_usable) and now_usable:
                            slot_info = f"插槽 #{i + 1}"
                            self.root.after(
                                0, lambda s=slot_info: self.on_card_inserted_event(s)
                            )
            finally:
                try:
                    winscard.SCardReleaseContext(h_ctx.value)
                except Exception:
                    pass

    def on_card_inserted_event(self, slot_info):
        if not self.auto_sense_var.get() or self.test_mode_var.get():
            return
        if self.is_querying or (time.time() - self.last_own_read_ts) < 2.5:
            return
        if self._auto_after_id:
            self.root.after_cancel(self._auto_after_id)
        self.lbl_card_bar.configure(
            text=f"💳 感應到健保卡推入（{slot_info}），緩衝 2 秒後自動背景靜默檢查..."
        )
        self._auto_after_id = self.root.after(
            2000, lambda: self.start_check_thread(auto_triggered=True)
        )

    # =========================================================================
    # 背景服務 2：本機網頁外掛連動埠 (127.0.0.1:18765) ＋ 單一實例喚醒
    # =========================================================================
    def start_local_bridge_server(self):
        app_ref = self

        class BridgeHandler(BaseHTTPRequestHandler):
            def log_message(self, fmt, *args):
                pass

            def _send_json(self, code, obj):
                body = json.dumps(obj, ensure_ascii=False).encode("utf-8")
                self.send_response(code)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
                self.send_header("Access-Control-Allow-Headers", "Content-Type")
                self.send_header("Content-Length", str(len(body)))
                self.end_headers()
                self.wfile.write(body)

            def do_OPTIONS(self):
                self._send_json(200, {"ok": True})

            def do_GET(self):
                if self.path.startswith("/ping"):
                    self._send_json(200, {"ok": True, "app": "nhi_ddi_clinic_5x"})
                else:
                    self._send_json(404, {"ok": False})

            def do_POST(self):
                length = int(self.headers.get("Content-Length", "0") or 0)
                raw = self.rfile.read(length) if length > 0 else b"{}"
                try:
                    payload = json.loads(raw.decode("utf-8", errors="ignore"))
                except Exception:
                    payload = {}

                if self.path.startswith("/show_window"):
                    app_ref.root.after(
                        0, lambda: app_ref.popup_alert_window(play_sound=False)
                    )
                    self._send_json(200, {"ok": True})
                elif self.path.startswith("/lab_sync"):
                    app_ref.root.after(
                        0, lambda p=payload: app_ref.apply_web_lab_sync(p)
                    )
                    self._send_json(200, {"ok": True})
                else:
                    self._send_json(404, {"ok": False})

        try:
            self.http_server = ThreadingHTTPServer(
                (LOCAL_BRIDGE_HOST, LOCAL_BRIDGE_PORT), BridgeHandler
            )
            threading.Thread(target=self.http_server.serve_forever, daemon=True).start()
        except Exception:
            self.http_server = None

    def apply_web_lab_sync(self, payload):
        """接收瀏覽器外掛抓取的 GOT / GPT / Creatinine / eGFR 並同步點亮上方色塊。"""
        if not isinstance(payload, dict):
            return
        pat_id_hint = str(payload.get("patIdHint") or "").strip()
        pat_name_hint = str(payload.get("patNameHint") or "").strip()
        has_lab_tab = payload.get("hasLabTab", True)
        labs = payload.get("labs") or {}

        same_pat = is_compatible_masked_text(self.current_pat_id, pat_id_hint) and (
            is_compatible_masked_text(self.current_pat_name, pat_name_hint)
        )
        hcv_msgs = self.current_hcv_msgs if same_pat else []
        self.current_web_labs = payload

        got = labs.get("got")
        gpt = labs.get("gpt")
        cr = labs.get("cr")
        egfr = labs.get("egfr")

        abnormal_synced = False

        if got or gpt:
            got_v = got.get("value", "-") if got else "未驗"
            gpt_v = gpt.get("value", "-") if gpt else "未驗"
            got_n = got.get("num") if got else None
            gpt_n = gpt.get("num") if gpt else None
            got_f = str(got.get("flag") or "") if got else ""
            gpt_f = str(gpt.get("flag") or "") if gpt else ""

            ref_item = got or gpt
            dt_str = ref_item.get("date") or "近期"
            hosp_str = ref_item.get("hosp") or "雲端院所"

            is_crit = (got_n is not None and got_n >= 80) or (
                gpt_n is not None and gpt_n >= 80
            )
            is_high = (
                (got_n is not None and got_n > 40)
                or (gpt_n is not None and gpt_n > 40)
                or bool(got_f)
                or bool(gpt_f)
            )

            if hcv_msgs:
                lv_level = "CRITICAL"
                lv_title = f"🚨 C肝陽性｜GOT {got_v} / GPT {gpt_v}"
                lv_detail = f"{' '.join(hcv_msgs)}｜檢驗：{dt_str} ({hosp_str})"
                abnormal_synced = True
            elif is_crit:
                lv_level = "CRITICAL"
                lv_title = f"🚨 肝指數顯著異常：GOT {got_v} / GPT {gpt_v}"
                lv_detail = f"🌐 雲端網頁同步：{dt_str} ({hosp_str}) — 口服抗黴菌藥請高度留心！"
                abnormal_synced = True
            elif is_high:
                lv_level = "HIGH"
                lv_title = f"⚠️ 肝指數偏高：GOT {got_v} / GPT {gpt_v}"
                lv_detail = f"🌐 雲端網頁同步：{dt_str} ({hosp_str}) — 開藥前請評估肝功能"
                abnormal_synced = True
            else:
                lv_level = "NORMAL"
                lv_title = f"✅ 肝指數正常：GOT {got_v} / GPT {gpt_v}"
                lv_detail = f"🌐 雲端網頁同步：{dt_str} ({hosp_str})｜無 C 肝陽性警示"

            self.set_banner("liver", lv_level, lv_title, lv_detail)
        elif has_lab_tab is False and not hcv_msgs:
            self.set_banner(
                "liver",
                "INFO",
                "✅ 無 C 肝警示｜近期無檢驗",
                "🌐 雲端網頁確認：近半年無「檢查與檢驗」紀錄",
            )

        if cr or egfr:
            parts = []
            if cr:
                parts.append(f"Cr: {cr.get('value')} {cr.get('unit', '')}".strip())
            if egfr:
                parts.append(f"eGFR: {egfr.get('value')}")
            ref_r = egfr or cr
            r_dt = ref_r.get("date") or ""
            web_line = f"🌐 網頁同步：{'｜'.join(parts)}" + (f" ({r_dt})" if r_dt else "")

            egfr_n = egfr.get("num") if egfr else None
            cr_n = cr.get("num") if cr else None

            if not same_pat or self.banner_levels["renal"] in ("IDLE", "UNKNOWN"):
                if egfr_n is not None:
                    if egfr_n < 30:
                        self.set_banner(
                            "renal", "CRITICAL", f"🚨 重度腎不全｜eGFR: {egfr['value']}", web_line
                        )
                        abnormal_synced = True
                    elif egfr_n < 45:
                        self.set_banner(
                            "renal", "HIGH", f"⚠️ 中度腎異常｜eGFR: {egfr['value']}", web_line
                        )
                        abnormal_synced = True
                    elif egfr_n < 60:
                        self.set_banner(
                            "renal", "MODERATE", f"⚠️ 輕中度腎異常｜eGFR: {egfr['value']}", web_line
                        )
                        abnormal_synced = True
                    else:
                        self.set_banner(
                            "renal", "NORMAL", f"✅ 腎功能正常｜eGFR: {egfr['value']}", web_line
                        )
                elif cr_n is not None:
                    lvl = "CRITICAL" if cr_n >= 2.0 else ("MODERATE" if cr_n > 1.3 else "NORMAL")
                    if lvl != "NORMAL":
                        abnormal_synced = True
                    icon = "🚨" if lvl == "CRITICAL" else ("⚠️" if lvl == "MODERATE" else "✅")
                    self.set_banner("renal", lvl, f"{icon} 肌酸酐 Cr: {cr['value']}", web_line)
            else:
                base_detail = self.current_renal_detail or ""
                if "🌐 網頁同步" not in base_detail:
                    merged_detail = f"{base_detail}\n{web_line}".strip()
                else:
                    merged_detail = re.sub(r"🌐 網頁同步.*$", web_line, base_detail)
                self.set_banner(
                    "renal", self.current_renal_level, self.current_renal_title, merged_detail
                )

        if abnormal_synced:
            self.flash_taskbar_icon()

    # =========================================================================
    # 背景服務 3：Windows 全域熱鍵 Alt + F9
    # =========================================================================
    def start_global_hotkey_thread(self):
        threading.Thread(target=self._hotkey_loop, daemon=True).start()

    def _hotkey_loop(self):
        try:
            user32 = ctypes.windll.user32
            kernel32 = ctypes.windll.kernel32
            self._hotkey_tid = kernel32.GetCurrentThreadId()
            if not user32.RegisterHotKey(None, 1, MOD_ALT | MOD_NOREPEAT, VK_F9):
                return
            msg = wintypes.MSG()
            while not self._stop_event.is_set():
                ret = user32.GetMessageW(ctypes.byref(msg), None, 0, 0)
                if ret == 0 or ret == -1:
                    break
                if msg.message == WM_HOTKEY and msg.wParam == 1:
                    self.root.after(0, self.on_manual_hotkey)
        except Exception:
            pass
        finally:
            try:
                ctypes.windll.user32.UnregisterHotKey(None, 1)
            except Exception:
                pass

    def on_manual_hotkey(self):
        self.popup_alert_window(play_sound=False)
        self.start_check_thread(auto_triggered=False)

    def on_close(self):
        self._stop_event.set()
        if self._hotkey_tid:
            try:
                ctypes.windll.user32.PostThreadMessageW(self._hotkey_tid, WM_QUIT, 0, 0)
            except Exception:
                pass
        if self.http_server:
            try:
                self.http_server.shutdown()
            except Exception:
                pass
        self.root.destroy()

    # =========================================================================
    # 讀卡與簽章：C:\NHI\LIB\CSHIS.dll (5.1.5.7) + CsHisX.nhicshisx.1
    # =========================================================================
    def read_cards_and_sign_5x(self):
        if self.py_bits != 32:
            raise RuntimeError(
                "目前是以 64 位元 Python 執行，無法載入 32 位元的 CSHIS.dll！\n"
                "請改用 32 位元 Python 啟動：py -3.13-32 nhi_ddi_clinic_5x.py"
            )
        if not HAS_WIN32COM:
            raise RuntimeError(
                "尚未安裝 pywin32，請執行：py -3.13-32 -m pip install requests pywin32"
            )
        if not os.path.exists(CSHIS_DLL_PATH):
            raise FileNotFoundError(f"找不到讀卡機控制軟體：{CSHIS_DLL_PATH}")

        cshis = ctypes.WinDLL(CSHIS_DLL_PATH)
        com_idx = int(self.config.get("com_port_index", 0))

        open_ret = cshis.csOpenCom(ctypes.c_int(com_idx))
        try:
            b_pat = ctypes.create_string_buffer(128)
            l_pat = ctypes.c_int(128)
            ret_pat = cshis.hisGetBasicData(b_pat, ctypes.byref(l_pat))
            if ret_pat != 0:
                raise RuntimeError(
                    f"讀取病人健保卡失敗 (csOpenCom={open_ret}, hisGetBasicData={ret_pat})。\n"
                    "請確認病人健保卡已插妥，且展望系統目前未正在讀寫卡片。"
                )

            raw = b_pat.raw
            pat_card_id = raw[0:12].decode("ascii", errors="ignore").strip()
            pat_name = raw[12:32].decode("big5", errors="ignore").strip()
            pat_id = raw[32:42].decode("ascii", errors="ignore").strip()
            pat_bday = raw[42:49].decode("ascii", errors="ignore").strip()
            pat_sex = raw[49:50].decode("ascii", errors="ignore").strip()

            b_hosp = ctypes.create_string_buffer(32)
            l_hosp = ctypes.c_int(32)
            cshis.csGetHospID(b_hosp, ctypes.byref(l_hosp))
            hosp_id = (
                b_hosp.raw[0:10].decode("ascii", errors="ignore").strip()
                or self.config["hosp_id_fallback"]
            )

            b_sam = ctypes.create_string_buffer(32)
            l_sam = ctypes.c_int(32)
            cshis.csGetCardNo(ctypes.c_int(1), b_sam, ctypes.byref(l_sam))
            sam_id = b_sam.raw[0:12].decode("ascii", errors="ignore").strip()

            b_hcasn = ctypes.create_string_buffer(32)
            l_hcasn = ctypes.c_int(32)
            ret_hca = cshis.hpcGetHPCSN(b_hcasn, ctypes.byref(l_hcasn))
            hca_card_id = b_hcasn.raw[0:12].decode("ascii", errors="ignore").strip()
            if ret_hca != 0 or not hca_card_id:
                raise RuntimeError(
                    f"讀取醫師卡失敗 (hpcGetHPCSN={ret_hca})，請確認醫事人員卡已插妥並完成 PIN 認證。"
                )

            b_hcassn = ctypes.create_string_buffer(32)
            l_hcassn = ctypes.c_int(32)
            cshis.hpcGetHPCSSN(b_hcassn, ctypes.byref(l_hcassn))
            hca_id = (
                b_hcassn.raw[0:10].decode("ascii", errors="ignore").strip()
                or self.config["hca_id_fallback"]
            )

            b_hcaname = ctypes.create_string_buffer(64)
            l_hcaname = ctypes.c_int(64)
            cshis.hpcGetHPCCNAME(b_hcaname, ctypes.byref(l_hcaname))
            doc_name = b_hcaname.value.decode("big5", errors="ignore").strip() or hca_id
        finally:
            cshis.csCloseCom()
            self.last_own_read_ts = time.time()

        vpn_ax = win32com.client.Dispatch(CSHISX_PROG_ID)
        client_random = str(vpn_ax.VPNGetRandomX()).strip()
        if len(client_random) != 20:
            raise RuntimeError(f"VPNGetRandomX 取得亂數異常，回傳：[{client_random}]")

        signature = str(vpn_ax.VPNH_SignX(client_random, "4", "91")).strip()
        self.last_own_read_ts = time.time()
        if len(signature) < 256:
            raise RuntimeError(
                f"VPNH_SignX 簽章失敗（回傳：[{signature}]）\n"
                "請確認 SAM 卡、醫事人員卡與健保卡皆在線且已完成 PIN 認證。"
            )

        sex_str = "男" if pat_sex == "M" else ("女" if pat_sex == "F" else "")
        bday_age = format_roc_bday_and_age(pat_bday)
        summary = (
            f"🏥 {hosp_id} ｜ 👨‍⚕️ {doc_name} ｜ "
            f"👤 病人：{pat_name} ({sex_str}・{bday_age}・{pat_id})"
        )

        pat_meta = {"pat_id": pat_id, "pat_name": pat_name}
        return (
            summary,
            pat_meta,
            {
                "sHospId": hosp_id,
                "sHcaId": hca_id,
                "sPatId": pat_id,
                "sPatCardType": "2",
                "sHcaCardId": hca_card_id,
                "sPatCardId": pat_card_id,
                "sClientRandom": client_random,
                "sSignature": signature,
                "vhcCloudToken": "",
                "sSamId": sam_id,
            },
        )

    # =========================================================================
    # 背景查詢健保雲端 API（含寫卡佔線自動重試、5 分鐘防重查、05 隔離拆單）
    # =========================================================================
    def start_check_thread(self, auto_triggered=False):
        if self.is_querying:
            return
        self.config["hosp_id_fallback"] = self.ent_hosp.get().strip()
        self.config["hca_id_fallback"] = self.ent_hca.get().strip()
        self.save_config()

        selected_drugs = [d for v, d in self.drug_vars if v.get()]
        test_mode = self.test_mode_var.get()

        self.is_querying = True
        self.btn_check.configure(state="disabled", text="⏳ 正在讀卡並連線健保雲端查詢中...")
        if not auto_triggered:
            self.txt_result.delete("1.0", tk.END)
            self.txt_result.insert(tk.END, "連線查詢中，請稍候...\n")

        threading.Thread(
            target=self._worker_query,
            args=(selected_drugs, test_mode, auto_triggered),
            daemon=True,
        ).start()

    def _post_cloud(self, base_auth, sub_blocks):
        payload = dict(base_auth)
        payload["sub"] = sub_blocks
        resp = requests.post(
            NHI_CLOUD_API_URL,
            json=payload,
            headers={"Content-Type": "application/json"},
            timeout=15,
            verify=False,
        )
        resp.raise_for_status()
        return resp.json()

    def _worker_query(self, selected_drugs, test_mode, auto_triggered):
        if HAS_WIN32COM:
            pythoncom.CoInitialize()
        try:
            if test_mode:
                hosp = self.config["hosp_id_fallback"]
                hca = self.config["hca_id_fallback"]
                summary = f"🧪 官方測試模式 ｜ 機構：{hosp} ｜ 醫師：{hca} ｜ 測試病人：Z299999992"
                pat_meta = {"pat_id": "Z299999992", "pat_name": "測試個案"}
                base_auth = {
                    "sHospId": hosp,
                    "sHcaId": hca,
                    "sPatId": "Z299999992",
                    "sPatCardType": "2",
                    "sHcaCardId": "000000243387",
                    "sPatCardId": "000073983649",
                    "sClientRandom": "0" * 20,
                    "sSignature": "0" * 512,
                    "vhcCloudToken": "",
                    "sSamId": "001000000103",
                }
            else:
                try:
                    summary, pat_meta, base_auth = self.read_cards_and_sign_5x()
                except Exception:
                    if not auto_triggered:
                        raise
                    time.sleep(2.8)
                    summary, pat_meta, base_auth = self.read_cards_and_sign_5x()

                pid = pat_meta["pat_id"]
                if (
                    auto_triggered
                    and pid
                    and pid == self.last_auto_pat_id
                    and (time.time() - self.last_auto_pat_ts) < AUTO_DEDUP_SECONDS
                ):
                    self.root.after(
                        0, lambda s=summary: self._on_skip_duplicate(s)
                    )
                    return

                self.last_auto_pat_id = pid
                self.last_auto_pat_ts = time.time()

            cur_probe = self.config.get("working_nsaid_probe", NSAID_PROBE_CANDIDATES[0])
            organ_blocks = [
                {"sType": "02", "sub": [{"sOrder": "X"}]},
                {"sType": "11", "sub": [{"sOrder": "X"}]},
                {"sType": "07", "sub": [{"sOrder": cur_probe}]},
            ]

            west_orders = [
                {"sOrder": d["code"]}
                for d in selected_drugs
                if d.get("category", "08") == "08"
            ]
            tcm_orders = [
                {"sOrder": d["code"]}
                for d in selected_drugs
                if d.get("category") == "09"
            ]
            ddi_blocks = []
            if west_orders:
                ddi_blocks.append({"sType": "08", "sub": west_orders})
            if tcm_orders:
                ddi_blocks.append({"sType": "09", "sub": tcm_orders})

            data = self._post_cloud(base_auth, organ_blocks + ddi_blocks)
            invalid_drugs = []
            renal_probe_ok = True

            if ci_get(data, "rtnCode") == "05":
                base_res = self._post_cloud(
                    base_auth,
                    [
                        {"sType": "02", "sub": [{"sOrder": "X"}]},
                        {"sType": "11", "sub": [{"sOrder": "X"}]},
                    ],
                )
                data = base_res
                merged_subs = list(ci_get(base_res, "sub", []))

                renal_probe_ok = False
                candidates = [cur_probe] + [
                    c for c in NSAID_PROBE_CANDIDATES if c != cur_probe
                ]
                for cand in candidates:
                    r07 = self._post_cloud(
                        base_auth, [{"sType": "07", "sub": [{"sOrder": cand}]}]
                    )
                    if ci_get(r07, "rtnCode") == "00":
                        renal_probe_ok = True
                        if cand != cur_probe:
                            self.config["working_nsaid_probe"] = cand
                            self.save_config()
                        merged_subs.extend(ci_get(r07, "sub", []))
                        break

                for d in selected_drugs:
                    cat = d.get("category", "08")
                    s_res = self._post_cloud(
                        base_auth, [{"sType": cat, "sub": [{"sOrder": d["code"]}]}]
                    )
                    if ci_get(s_res, "rtnCode") == "00":
                        merged_subs.extend(ci_get(s_res, "sub", []))
                    else:
                        invalid_drugs.append(d)

                data["sub"] = merged_subs

            self.root.after(
                0,
                lambda: self.render_results(
                    summary,
                    pat_meta,
                    data,
                    selected_drugs,
                    invalid_drugs,
                    renal_probe_ok,
                    auto_triggered,
                ),
            )
        except Exception as e:
            err_msg = str(e)
            self.root.after(
                0, lambda: self.render_error(err_msg, auto_triggered=auto_triggered)
            )
        finally:
            self.is_querying = False
            if HAS_WIN32COM:
                pythoncom.CoUninitialize()

    def _on_skip_duplicate(self, summary):
        self.btn_check.configure(
            state="normal",
            text="🔍 手動立即讀卡檢查（F9 / Alt+F9）➔ 檢查【肝腎・過敏・DDI 交互作用】",
        )
        self.lbl_card_bar.configure(
            text=f"{summary} ｜ ℹ️ 5 分鐘內已完成檢查（如需重查請按按鈕或 F9）"
        )

    def render_error(self, err_msg, auto_triggered=False):
        self.btn_check.configure(
            state="normal",
            text="🔍 手動立即讀卡檢查（F9 / Alt+F9）➔ 檢查【肝腎・過敏・DDI 交互作用】",
        )
        if auto_triggered:
            first_line = err_msg.splitlines()[0] if err_msg else "讀取未完成"
            self.lbl_card_bar.configure(text=f"ℹ️ 背景插槽讀取略過：{first_line}")
            return

        self.lbl_card_bar.configure(text="⚠️ 讀卡或連線未完成（請見下方錯誤說明）")
        for t in ("renal", "liver", "allergy"):
            self.set_banner(t, "IDLE", "⚠️ 尚未取得資料", "請見下方錯誤訊息")
        self.txt_result.delete("1.0", tk.END)
        self.txt_result.insert(tk.END, f"❌ 查詢失敗：\n{err_msg}\n")

    # =========================================================================
    # 更新三格警示色塊、DDI 明細與自動置頂判斷
    # =========================================================================
    def update_top_three_banners(self, sub_list, renal_probe_ok):
        renal_msgs, hcv_msgs, active_allergies = [], [], []

        for block in sub_list:
            otype = ci_get(block, "oType")
            items = ci_get(block, "sub", [])

            if otype == "07":
                for it in items:
                    if ci_get(it, "oMsg"):
                        renal_msgs.append(ci_get(it, "oMsg"))
            elif otype == "11":
                for it in items:
                    if ci_get(it, "oMsg"):
                        hcv_msgs.append(ci_get(it, "oMsg"))
            elif otype == "02":
                for it in items:
                    if ci_get(it, "upload_Flag") == "N":
                        raw_txt = ci_get(it, "allergyOrder_text", "").strip()
                        cleaned = re.sub(
                            r"^(其他|藥物|顯影劑|食物)註記內容:\s*", "", raw_txt
                        ).strip()
                        if cleaned:
                            active_allergies.append(
                                (
                                    cleaned,
                                    ci_get(it, "hospName", ""),
                                    ci_get(it, "upload_date", ""),
                                )
                            )

        self.current_hcv_msgs = hcv_msgs

        if not renal_probe_ok:
            r_lvl = "UNKNOWN"
            r_title = "⚠️ 腎功能探測碼需更新"
            r_detail = "內建探測碼皆停用，請於 JSON 設定檔更新 working_nsaid_probe"
        elif renal_msgs:
            msg = renal_msgs[0]
            if "慢性透析" in msg:
                r_lvl = "CRITICAL"
                r_title = "🚨 腎功能極高危：慢性透析(洗腎)"
                r_detail = "嚴格注意腎代謝藥物減量，避免腎毒性藥物！"
            elif "慢性腎臟病" in msg:
                m_st = re.search(r"可能為\[?(第?[345][AB]?)\]?期慢性腎臟病", msg)
                m_egfr = re.search(
                    r"eGFR\s*值為\s*\[?([0-9.]+)\]?", msg, re.IGNORECASE
                )
                m_dt = re.search(r"採檢日期為\s*\[?([0-9/]+)\]?", msg)
                st_str = m_st.group(1) if m_st else "未明"
                egfr_v = float(m_egfr.group(1)) if m_egfr else None
                egfr_s = f"eGFR: {m_egfr.group(1)}" if m_egfr else "依診斷碼判定"
                dt_s = f" (採檢日 {m_dt.group(1)})" if m_dt else ""

                if (
                    "4" in st_str
                    or "5" in st_str
                    or (egfr_v is not None and egfr_v < 30)
                ):
                    r_lvl = "CRITICAL"
                    r_title = f"🚨 重度腎不全：CKD {st_str} 期｜{egfr_s}{dt_s}"
                    r_detail = "eGFR < 30，腎排泄藥物（含部分抗組織胺/抗病毒藥）務必調整劑量！"
                elif "3B" in st_str.upper() or (
                    egfr_v is not None and egfr_v < 45
                ):
                    r_lvl = "HIGH"
                    r_title = f"⚠️ 中度腎異常：CKD {st_str} 期｜{egfr_s}{dt_s}"
                    r_detail = "eGFR 30~45，請留意腎代謝藥物劑量與口服 NSAIDs 天數。"
                else:
                    r_lvl = "MODERATE"
                    r_title = f"⚠️ 輕中度腎異常：CKD {st_str} 期｜{egfr_s}{dt_s}"
                    r_detail = "eGFR 45~60，長期用藥建議追蹤腎功能。"
            elif "無血清肌酸酐檢測值" in msg:
                m_m = re.search(r"近\s*(\d+)\s*個月", msg)
                m_txt = f"近 {m_m.group(1)} 個月" if m_m else "近期"
                r_lvl = "UNKNOWN"
                r_title = f"ℹ️ 腎功能不明：{m_txt}無雲端紀錄"
                r_detail = "雲端無近期肌酸酐/eGFR，高齡病患開藥仍請留意。"
            else:
                r_lvl = "MODERATE"
                r_title = "⚠️ 腎功能提示"
                r_detail = msg
        else:
            r_lvl = "NORMAL"
            r_title = "✅ 腎功能近期正常 (eGFR ≥ 60)"
            r_detail = "近 6 個月內有雲端檢驗且無慢性腎臟病或洗腎警示"

        self.current_renal_level = r_lvl
        self.current_renal_title = r_title
        self.current_renal_detail = r_detail
        self.set_banner("renal", r_lvl, r_title, r_detail)

        if hcv_msgs:
            self.set_banner(
                "liver",
                "CRITICAL",
                "🚨 肝炎警示：近 3 個月 C 肝陽性！",
                " ".join(hcv_msgs),
            )
        else:
            self.set_banner(
                "liver",
                "INFO",
                "✅ 無 C 肝陽性警示",
                "無近 3 個月 C 肝陽性註記（點閱雲端網頁將自動同步 GOT/GPT 數值）",
            )

        if active_allergies:
            uniq = list(dict.fromkeys([a[0] for a in active_allergies]))
            preview = "、".join(uniq[:4]) + ("..." if len(uniq) > 4 else "")
            self.set_banner(
                "allergy",
                "CRITICAL",
                f"🚨 藥物過敏紀錄 ({len(uniq)} 項)",
                f"過敏品項：{preview}（詳見下方明細）",
            )
        else:
            self.set_banner(
                "allergy",
                "NORMAL",
                "✅ 查無雲端過敏紀錄",
                "目前雲端與健保卡上傳紀錄無有效過敏藥物註記",
            )

        return active_allergies

    def render_results(
        self,
        summary,
        pat_meta,
        data,
        selected_drugs,
        invalid_drugs,
        renal_probe_ok=True,
        auto_triggered=False,
    ):
        self.btn_check.configure(
            state="normal",
            text="🔍 手動立即讀卡檢查（F9 / Alt+F9）➔ 檢查【肝腎・過敏・DDI 交互作用】",
        )

        new_pid = pat_meta.get("pat_id", "")
        if new_pid and new_pid != self.current_pat_id:
            self.current_web_labs = None
        self.current_pat_id = new_pid
        self.current_pat_name = pat_meta.get("pat_name", "")

        self.lbl_card_bar.configure(text=summary)
        self.txt_result.delete("1.0", tk.END)

        rtn_code = ci_get(data, "rtnCode")
        err_map = {
            "01": "參數解析失敗",
            "02": "個案驗章失敗 (實體健保卡三卡簽章驗證失敗)",
            "03": "連線數過多，請稍候再試",
            "04": "健保署系統發生異常",
            "05": "非適用特定醫囑代碼之醫令範圍",
            "06": "資料類別錯誤",
            "07": "虛擬健保卡驗章失敗",
            "08": "回傳資料筆數過多",
            "09": "非院所的專兼任醫師(藥師) — 請檢查上方「醫師身分證號」是否填寫正確！",
        }

        if rtn_code != "00":
            msg = err_map.get(rtn_code, "未知錯誤")
            for t in ("renal", "liver", "allergy"):
                self.set_banner(t, "IDLE", f"⚠️ API 異常 ({rtn_code})", msg)
            self.txt_result.insert(
                tk.END, f"⚠️ 健保雲端 API 回傳異常 (代碼 {rtn_code})：{msg}\n"
            )
            return

        sub_list = ci_get(data, "sub", [])
        active_allergies = self.update_top_three_banners(sub_list, renal_probe_ok)

        if self.current_web_labs and is_compatible_masked_text(
            self.current_pat_id, self.current_web_labs.get("patIdHint", "")
        ):
            self.apply_web_lab_sync(self.current_web_labs)

        if active_allergies:
            self.txt_result.insert(tk.END, "🚨 【病人雲端有效過敏藥物明細】\n")
            for drug_txt, hosp, udate in active_allergies:
                self.txt_result.insert(
                    tk.END, f"   • {drug_txt}  （登錄院所：{hosp}，日期：{udate}）\n"
                )
            self.txt_result.insert(tk.END, "=" * 76 + "\n")

        if invalid_drugs:
            for bad in invalid_drugs:
                self.txt_result.insert(
                    tk.END,
                    f"⚠️ 略過停用或非健保代碼：[{bad['name']} ({bad['code']})] 已不在當期適用清單內。\n",
                )
            self.txt_result.insert(tk.END, "-" * 76 + "\n")

        code_to_name = {d["code"]: d["name"] for d in self.config["drugs"]}
        found_ddi = False

        if not selected_drugs:
            self.txt_result.insert(
                tk.END,
                "ℹ️ 已完成上方【腎功能・肝功能・過敏史】檢查（下方未勾選藥品，未執行 DDI 比對）。\n",
            )
        else:
            for block in sub_list:
                otype = ci_get(block, "oType")
                if otype == "08" and int(ci_get(block, "rtnNum", "0") or 0) > 0:
                    for order_item in ci_get(block, "sub", []):
                        my_code = ci_get(order_item, "oOrder")
                        my_name = code_to_name.get(my_code, my_code)
                        for ddi in ci_get(order_item, "sub", []):
                            found_ddi = True
                            ddi_order = ci_get(ddi, "ddiOrder")
                            hosp_str = "、".join(
                                [
                                    f"{ci_get(h, 'hospName')}({ci_get(h, 'funcDT')})"
                                    for h in ci_get(ddi, "hosPsub", [])
                                ]
                            )
                            self.txt_result.insert(
                                tk.END,
                                f"⚡ 【交互作用警示】本院擬開：{my_name} [{my_code}]  vs  病人餘藥：[{ddi_order}]\n"
                                f"   📍 餘藥來源院所(日期)：{hosp_str}\n",
                            )
                            for det in ci_get(ddi, "ddIsub", []):
                                self.txt_result.insert(
                                    tk.END,
                                    f"   💊 成分比對：{ci_get(det, 'ingName')}  ⚡  {ci_get(det, 'ddiIngName')}\n"
                                    f"   ⚠️ 交互結果：{ci_get(det, 'effect')}\n"
                                    f"   🔬 作用機轉：{ci_get(det, 'mechanism')}\n"
                                    f"   💡 處置建議：{ci_get(det, 'management')}\n"
                                    f"   🔄 替代建議：{ci_get(det, 'alternatives')}\n",
                                )
                            self.txt_result.insert(tk.END, "=" * 76 + "\n")

                elif otype == "09" and int(ci_get(block, "rtnNum", "0") or 0) > 0:
                    for item in ci_get(block, "sub", []):
                        found_ddi = True
                        my_code = ci_get(item, "oOrder")
                        my_name = code_to_name.get(my_code, my_code)
                        hosp_str = "、".join(
                            [
                                f"{ci_get(h, 'hospName')}({ci_get(h, 'funcDT')})"
                                for h in ci_get(item, "sub", [])
                            ]
                        )
                        self.txt_result.insert(
                            tk.END,
                            f"⚡ 【中西藥交互作用】擬開立：{my_name} [{my_code}] (含{ci_get(item, 'cDrugName')})\n"
                            f"   📍 西藥餘藥來源：{hosp_str}\n"
                            f"   ⚠️ 交互結果：{ci_get(item, 'effect')}\n"
                            f"   💡 處置建議：{ci_get(item, 'management')}\n"
                            + "=" * 76
                            + "\n",
                        )

            if not found_ddi and len(invalid_drugs) < len(selected_drugs):
                self.txt_result.insert(
                    tk.END,
                    "✅ 安全！所勾選之有效藥品與病人目前雲端餘藥（含西藥、中藥）皆【查無交互作用】。\n",
                )

        has_clinical_alert = (
            self.banner_levels["renal"] in ("CRITICAL", "HIGH", "MODERATE")
            or self.banner_levels["liver"] in ("CRITICAL", "HIGH", "MODERATE")
            or bool(active_allergies)
            or found_ddi
        )

        if auto_triggered and has_clinical_alert:
            self.popup_alert_window(play_sound=True)


if __name__ == "__main__":
    if wake_existing_instance_if_running():
        sys.exit(0)
    root = tk.Tk()
    app = ClinicNHIDDIApp(root)
    if "--minimized" in sys.argv:
        root.after(100, root.iconify)
    root.mainloop()