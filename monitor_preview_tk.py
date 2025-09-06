import threading
import time
import ctypes
import ctypes.wintypes as wintypes
import tkinter as tk
from tkinter import ttk, messagebox, colorchooser
from typing import Optional, List, Dict, Tuple

import mss
import numpy as np
import cv2
from PIL import Image, ImageTk

# --- Windows API / pywin32 (opsiyonel ama önerilir) --------------------------
try:
    import win32con, win32gui, win32api
    HAS_WIN32 = True
except Exception:
    HAS_WIN32 = False

APP_TITLE = "CANMOSE TV/Monitör Önizleme"

# --- DPI AWARENESS ------------------------------------------------------------
user32 = ctypes.windll.user32 if hasattr(ctypes, "windll") else None
shcore = ctypes.windll.shcore if hasattr(ctypes, "windll") else None

DPI_AWARE_CTX_PER_MONITOR_V2 = ctypes.c_void_p(-4).value  # PER_MONITOR_AWARE_V2
MDT_EFFECTIVE_DPI = 0
MONITOR_DEFAULTTONEAREST = 2

class POINT(ctypes.Structure):
    _fields_ = [("x", ctypes.c_long), ("y", ctypes.c_long)]

def make_process_dpi_aware():
    if not user32:
        return
    try:
        user32.SetProcessDpiAwarenessContext(ctypes.c_void_p(DPI_AWARE_CTX_PER_MONITOR_V2))
    except Exception:
        try:
            if shcore:
                shcore.SetProcessDpiAwareness(2)  # PER_MONITOR_AWARE
        except Exception:
            try:
                user32.SetProcessDPIAware()
            except Exception:
                pass

def get_system_scale_factor() -> float:
    if user32 and hasattr(user32, "GetDpiForSystem"):
        try:
            dpi = user32.GetDpiForSystem()
            return dpi / 96.0
        except Exception:
            pass
    return 1.0

def monitor_from_point(px: int, py: int):
    if not user32:
        return None
    pt = POINT(px, py)
    return user32.MonitorFromPoint(pt, MONITOR_DEFAULTTONEAREST)

def get_dpi_for_monitor(hmon) -> float:
    if not shcore:
        return get_system_scale_factor()
    dpiX = wintypes.UINT()
    dpiY = wintypes.UINT()
    try:
        hr = shcore.GetDpiForMonitor(hmon, MDT_EFFECTIVE_DPI,
                                     ctypes.byref(dpiX), ctypes.byref(dpiY))
        if hr == 0 and dpiX.value:
            return float(dpiX.value) / 96.0
    except Exception:
        pass
    return get_system_scale_factor()

def try_get_physical_cursor_pos() -> Optional[Tuple[int,int]]:
    if user32 and hasattr(user32, "GetPhysicalCursorPos"):
        pt = POINT()
        if user32.GetPhysicalCursorPos(ctypes.byref(pt)):
            return (pt.x, pt.y)
    return None

def get_physical_cursor_pos_robust() -> Tuple[int,int,dict]:
    """Fiziksel piksel koordinatı döndür (DPI sağlam) + debug meta."""
    meta = {"method": "", "scale": None}
    phys = try_get_physical_cursor_pos()
    if phys:
        meta["method"] = "GetPhysicalCursorPos"
        meta["scale"] = 1.0
        return phys[0], phys[1], meta

    if HAS_WIN32:
        x_log, y_log = win32api.GetCursorPos()
        hmon = monitor_from_point(x_log, y_log)
        scale = get_dpi_for_monitor(hmon) if hmon else get_system_scale_factor()
        x = int(round(x_log * scale))
        y = int(round(y_log * scale))
        meta["method"] = "GetCursorPos * per-monitor DPI"
        meta["scale"] = scale
        return x, y, meta

    return 0, 0, {"method": "none", "scale": None}

# --- yardımcılar --------------------------------------------------------------
def list_monitors() -> List[Dict]:
    with mss.mss() as sct:
        return sct.monitors  # [0]=tüm, 1..N=tek tek

def monitors_to_options(monitors: List[Dict]) -> List[str]:
    opts = []
    for idx, m in enumerate(monitors):
        if idx == 0:
            label = f"[0] Tüm Ekranlar  {m['width']}x{m['height']}"
        else:
            label = f"[{idx}] Monitör {idx}  {m['width']}x{m['height']} @({m['left']},{m['top']})"
        opts.append(label)
    return opts

def parse_monitor_index(text: str) -> int:
    try:
        return int(text.split("]")[0].strip("["))
    except Exception:
        return 1

def set_window_topmost(root: tk.Tk, enable: bool):
    root.wm_attributes("-topmost", 1 if enable else 0)
    if HAS_WIN32:
        try:
            hwnd = win32gui.FindWindow(None, root.title())
            if hwnd:
                flag = win32con.HWND_TOPMOST if enable else win32con.HWND_NOTOPMOST
                win32gui.SetWindowPos(
                    hwnd, flag, 0, 0, 0, 0,
                    win32con.SWP_NOMOVE | win32con.SWP_NOSIZE
                )
        except Exception:
            pass

def hex_to_bgr(hex_color: str) -> Tuple[int,int,int]:
    """'#RRGGBB' -> (B, G, R)"""
    hex_color = hex_color.strip()
    if hex_color.startswith("#"):
        hex_color = hex_color[1:]
    if len(hex_color) != 6:
        return (0, 0, 255)  # default red
    r = int(hex_color[0:2], 16)
    g = int(hex_color[2:4], 16)
    b = int(hex_color[4:6], 16)
    return (b, g, r)

# --- worker thread ------------------------------------------------------------
class PreviewWorker(threading.Thread):
    def __init__(self, ui, monitor_idx: int, scale: float, fps: int,
                 show_cursor: bool, debug: bool,
                 arrow_len: int, arrow_dir: str, arrow_color_bgr: Tuple[int,int,int],
                 anchor_mode: str, arrow_offset: int):
        super().__init__(daemon=True)
        self.ui = ui
        self.monitor_idx = monitor_idx
        self.scale = scale
        self.fps = max(1, int(fps))
        self.show_cursor = show_cursor
        self.debug = debug
        self.arrow_len = int(arrow_len)
        self.arrow_dir = arrow_dir
        self.arrow_color_bgr = arrow_color_bgr
        self.anchor_mode = anchor_mode
        self.arrow_offset = int(arrow_offset)
        self._stop = threading.Event()

    def stop(self):
        self._stop.set()

    def run(self):
        frame_interval = 1.0 / self.fps
        prev_time = 0.0
        try:
            with mss.mss() as sct:
                monitors = sct.monitors
                if self.monitor_idx < 0 or self.monitor_idx >= len(monitors):
                    self.ui.on_worker_stopped("Geçersiz monitör index")
                    return
                mon = monitors[self.monitor_idx]
                L, T, W, H = mon["left"], mon["top"], mon["width"], mon["height"]

                while not self._stop.is_set():
                    now = time.time()
                    if now - prev_time < frame_interval:
                        time.sleep(max(0, frame_interval - (now - prev_time)))
                    prev_time = time.time()

                    # --- yakalama (BGRA) -> BGR ve contiguous/yazılabilir ---
                    img = np.asarray(sct.grab(mon))
                    frame = img[:, :, :3].copy()
                    frame = np.ascontiguousarray(frame, dtype=np.uint8)

                    dbg_text = ""
                    # --- imleç ok çizimi ---
                    if HAS_WIN32 and self.show_cursor:
                        try:
                            cx_phys, cy_phys, meta = get_physical_cursor_pos_robust()
                            inside = (L <= cx_phys < L + W) and (T <= cy_phys < T + H)
                            cx = cx_phys - L
                            cy = cy_phys - T

                            if not inside:
                                # yedek: ham mantıksal
                                x_raw, y_raw = win32api.GetCursorPos()
                                if (L <= x_raw < L + W) and (T <= y_raw < T + H):
                                    cx = x_raw - L
                                    cy = y_raw - T
                                    inside = True
                                    meta["method"] += " | fallback raw"

                            if inside:
                                cx = int(cx); cy = int(cy)
                                # 1) yön vektörü
                                d = self.arrow_dir
                                vx, vy = 0, 0
                                if d == "sağ":                  vx, vy =  1,  0
                                elif d == "sol":               vx, vy = -1,  0
                                elif d == "yukarı":            vx, vy =  0, -1
                                elif d == "aşağı":             vx, vy =  0,  1
                                elif d == "çapraz sağ-aşağı":  vx, vy =  1,  1
                                elif d == "çapraz sol-aşağı":  vx, vy = -1,  1
                                elif d == "çapraz sağ-yukarı": vx, vy =  1, -1
                                elif d == "çapraz sol-yukarı": vx, vy = -1, -1

                                # normalize
                                length = max(1.0, (vx*vx + vy*vy)**0.5)
                                ux, uy = vx/length, vy/length

                                # 2) toplam uzunluk = arrow_len + offset
                                base_len = max(5, int(self.arrow_len))
                                total_len = base_len + int(self.arrow_offset)

                                # 3) anchor modu
                                if self.anchor_mode == "uçtan dışarı çiz":
                                    # başlangıç imleçte, uç dışarı
                                    x1 = cx
                                    y1 = cy
                                    x2 = int(np.clip(cx + ux*total_len, 0, frame.shape[1]-1))
                                    y2 = int(np.clip(cy + uy*total_len, 0, frame.shape[0]-1))
                                else:
                                    # "dışarıdan uca çiz": okun ucu imleçte biter
                                    x2 = cx
                                    y2 = cy
                                    x1 = int(np.clip(cx - ux*total_len, 0, frame.shape[1]-1))
                                    y1 = int(np.clip(cy - uy*total_len, 0, frame.shape[0]-1))

                                # 4) çiz
                                cv2.arrowedLine(
                                    frame, (x1, y1), (x2, y2),
                                    color=self.arrow_color_bgr, thickness=2,
                                    line_type=cv2.LINE_AA, tipLength=0.35
                                )

                            if self.debug:
                                dbg_text = (f"CUR:{cx_phys},{cy_phys} method={meta.get('method')} "
                                            f"MON:[{L},{T},{W},{H}] inside={inside} "
                                            f"anchor={self.anchor_mode} off={self.arrow_offset}")
                        except Exception as e:
                            if self.debug:
                                dbg_text = f"cursor_err:{e}"

                    # --- RGB'ye çevir, PIL image oluştur ---
                    frame_rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
                    image = Image.fromarray(frame_rgb)

                    # Canvas'a sığdır
                    canvas_w = self.ui.canvas.winfo_width()
                    canvas_h = self.ui.canvas.winfo_height()
                    if canvas_w > 0 and canvas_h > 0:
                        image.thumbnail((canvas_w, canvas_h), Image.LANCZOS)

                    photo = ImageTk.PhotoImage(image)
                    self.ui.root.after(0, self.ui.update_frame, photo)

                    if self.debug:
                        self.ui.root.after(0, self.ui.set_status, f"DEBUG: {dbg_text}")

        except Exception as e:
            self.ui.on_worker_stopped(f"Hata: {e}")

# --- GUI ----------------------------------------------------------------------
class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        root.title(APP_TITLE)
        root.geometry("1120x680")
        root.minsize(860, 520)

        self.worker: Optional[PreviewWorker] = None
        self.is_running = False
        self.photo_ref = None

        # ÜST PANEL
        top = ttk.Frame(root, padding=8)
        top.pack(side=tk.TOP, fill=tk.X)

        ttk.Label(top, text="Monitör:").pack(side=tk.LEFT, padx=(0,6))
        self.monitors = list_monitors()
        self.monitor_opts = monitors_to_options(self.monitors)
        default_val = self.monitor_opts[1] if len(self.monitor_opts) > 1 else self.monitor_opts[0]
        self.monitor_var = tk.StringVar(value=default_val)
        self.monitor_combo = ttk.Combobox(top, values=self.monitor_opts,
                                          textvariable=self.monitor_var, state="readonly", width=48)
        self.monitor_combo.pack(side=tk.LEFT)
        ttk.Button(top, text="Yenile", command=self.refresh_monitors).pack(side=tk.LEFT, padx=6)

        ttk.Label(top, text="Ölçek:").pack(side=tk.LEFT, padx=(10,6))
        self.scale_var = tk.DoubleVar(value=0.5)
        ttk.Spinbox(top, from_=0.1, to=1.5, increment=0.05,
                    textvariable=self.scale_var, width=6).pack(side=tk.LEFT)

        ttk.Label(top, text="FPS:").pack(side=tk.LEFT, padx=(10,6))
        self.fps_var = tk.IntVar(value=30)
        ttk.Spinbox(top, values=(15,20,24,30,45,60),
                    textvariable=self.fps_var, width=6).pack(side=tk.LEFT)

        self.topmost_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(top, text="Hep üstte",
                        variable=self.topmost_var, command=self.on_topmost).pack(side=tk.LEFT, padx=(10,6))

        self.cursor_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(top, text="İmleci Göster",
                        variable=self.cursor_var).pack(side=tk.LEFT, padx=(6,6))

        self.debug_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(top, text="Debug",
                        variable=self.debug_var).pack(side=tk.LEFT)

        # ORTA PANEL — OK AYARLARI
        mid = ttk.Frame(root, padding=(8,0,8,8))
        mid.pack(side=tk.TOP, fill=tk.X)

        ttk.Label(mid, text="Ok Boyu:").pack(side=tk.LEFT)
        self.arrow_len_var = tk.IntVar(value=24)
        ttk.Spinbox(mid, from_=8, to=200, increment=2,
                    textvariable=self.arrow_len_var, width=6).pack(side=tk.LEFT, padx=(6,12))

        ttk.Label(mid, text="Ok Yönü:").pack(side=tk.LEFT)
        self.arrow_dir_var = tk.StringVar(value="sağ")
        ttk.Combobox(mid, textvariable=self.arrow_dir_var, state="readonly",
                     values=["sağ","sol","yukarı","aşağı",
                             "çapraz sağ-aşağı","çapraz sol-aşağı",
                             "çapraz sağ-yukarı","çapraz sol-yukarı"],
                     width=18).pack(side=tk.LEFT, padx=(6,12))

        ttk.Label(mid, text="Bağlantı:").pack(side=tk.LEFT)
        self.anchor_mode_var = tk.StringVar(value="dışarıdan uca çiz")
        ttk.Combobox(mid, textvariable=self.anchor_mode_var, state="readonly",
                     values=["uçtan dışarı çiz", "dışarıdan uca çiz"],
                     width=18).pack(side=tk.LEFT, padx=(6,12))

        ttk.Label(mid, text="Ofset:").pack(side=tk.LEFT)
        self.arrow_offset_var = tk.IntVar(value=0)
        ttk.Spinbox(mid, from_=-60, to=120, increment=2,
                    textvariable=self.arrow_offset_var, width=6).pack(side=tk.LEFT, padx=(6,12))

        ttk.Label(mid, text="Ok Rengi:").pack(side=tk.LEFT)
        self.arrow_color_hex = tk.StringVar(value="#FF0000")  # kırmızı
        self.color_btn = ttk.Button(mid, text="Renk Seç", command=self.pick_color)
        self.color_btn.pack(side=tk.LEFT, padx=(6,6))
        self.color_preview = tk.Canvas(mid, width=28, height=18, bg="#FF0000",
                                       highlightthickness=1, highlightbackground="#888")
        self.color_preview.pack(side=tk.LEFT)

        # BUTONLAR
        btns = ttk.Frame(root, padding=(8,0,8,8))
        btns.pack(side=tk.TOP, fill=tk.X)
        self.start_btn = ttk.Button(btns, text="Önizlemeyi Başlat", command=self.start_preview)
        self.stop_btn  = ttk.Button(btns,  text="Durdur",            command=self.stop_preview, state=tk.DISABLED)
        self.start_btn.pack(side=tk.LEFT)
        self.stop_btn.pack(side=tk.LEFT, padx=6)

        # ÖNİZLEME
        self.canvas = tk.Canvas(root, bg="black")
        self.canvas.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0,8))

        # STATUS
        self.status_var = tk.StringVar(value="Hazır")
        ttk.Label(root, textvariable=self.status_var, relief=tk.SUNKEN, anchor="w")\
            .pack(side=tk.BOTTOM, fill=tk.X)

        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    # ---- GUI event handlers ----
    def pick_color(self):
        color = colorchooser.askcolor(color=self.arrow_color_hex.get(), title="Ok Rengini Seç")
        if color and color[1]:
            self.arrow_color_hex.set(color[1])
            self.color_preview.configure(bg=color[1])

    def refresh_monitors(self):
        try:
            self.monitors = list_monitors()
            self.monitor_opts = monitors_to_options(self.monitors)
            self.monitor_combo["values"] = self.monitor_opts
            self.monitor_var.set(self.monitor_opts[1] if len(self.monitor_opts) > 1 else self.monitor_opts[0])
            self.set_status("Monitör listesi yenilendi.")
        except Exception as e:
            messagebox.showerror("Hata", f"Monitörler alınamadı: {e}")

    def on_topmost(self):
        set_window_topmost(self.root, self.topmost_var.get())

    def start_preview(self):
        if self.is_running:
            return
        try:
            mon_idx = parse_monitor_index(self.monitor_var.get())
            scale   = float(self.scale_var.get())
            fps     = int(self.fps_var.get())
            show_cursor = bool(self.cursor_var.get())
            debug = bool(self.debug_var.get())
            arrow_len = int(self.arrow_len_var.get())
            arrow_dir = self.arrow_dir_var.get()
            anchor_mode = self.anchor_mode_var.get()
            arrow_offset = int(self.arrow_offset_var.get())
            arrow_color_bgr = hex_to_bgr(self.arrow_color_hex.get())

            if not (0.1 <= scale <= 1.5):
                raise ValueError("Ölçek 0.1–1.5 arası olmalı.")

            self.worker = PreviewWorker(
                self, mon_idx, scale, fps,
                show_cursor, debug,
                arrow_len, arrow_dir, arrow_color_bgr,
                anchor_mode, arrow_offset
            )
            self.worker.start()
            self.is_running = True
            self.start_btn.config(state=tk.DISABLED)
            self.stop_btn.config(state=tk.NORMAL)
            self.set_status(f"Önizleme başladı (Monitör {mon_idx}, ölçek {scale}, {fps} FPS).")
        except Exception as e:
            messagebox.showerror("Başlatma Hatası", str(e))

    def stop_preview(self):
        if self.worker:
            self.worker.stop()
        self.is_running = False
        self.start_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.DISABLED)
        self.set_status("Önizleme durduruldu.")

    def update_frame(self, photo: ImageTk.PhotoImage):
        self.photo_ref = photo  # GC koruması
        self.canvas.delete("all")
        w = self.canvas.winfo_width()
        h = self.canvas.winfo_height()
        img_w = photo.width()
        img_h = photo.height()
        x = (w - img_w) // 2 if w > img_w else 0
        y = (h - img_h) // 2 if h > img_h else 0
        self.canvas.create_image(x, y, anchor="nw", image=photo)

    def on_worker_stopped(self, reason: str):
        self.stop_preview()
        if reason:
            self.set_status(f"Önizleme sonlandı: {reason}")

    def set_status(self, text: str):
        self.status_var.set(text)

    def on_close(self):
        try:
            if self.worker:
                self.worker.stop()
        except Exception:
            pass
        self.root.destroy()

# --- main ---------------------------------------------------------------------
def main():
    make_process_dpi_aware()
    root = tk.Tk()
    try:
        from tkinter import font as tkfont
        tkfont.nametofont("TkDefaultFont").configure(size=10)
    except Exception:
        pass
    app = App(root)
    root.mainloop()

if __name__ == "__main__":
    main()
