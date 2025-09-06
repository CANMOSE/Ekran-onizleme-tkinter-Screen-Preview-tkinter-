# build_exe.py  —  MonitorPreview.exe'yi no-console üretir
from PyInstaller.__main__ import run
import os

# İstersen kendi ikonun varsa koy (örn. app.ico)
ICON = "app.ico"  # yoksa None bırak
ARGS = [
    "--noconsole",            # konsolsuz
    "--onefile",              # tek dosya exe
    "--clean",                # eski build artıkları silinsin
    "--name=MonitorPreview",  # çıktı adı
    # Bazı paketler için toplama/hidden-import ipuçları:
    "--collect-all=cv2",
    "--collect-all=PIL",
    "--collect-all=mss",
    "--collect-all=pywin32",
    "--hidden-import=win32api",
    "--hidden-import=win32gui",
    "--hidden-import=win32con",
    # kaynağımız:
    "monitor_preview_tk.py",
]

if ICON and os.path.exists(ICON):
    ARGS.insert(0, f"--icon={ICON}")

run(ARGS)
