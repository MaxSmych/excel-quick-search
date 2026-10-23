"""Установщик Excel Search 2 одним файлом (23.10.2026): программа (Excel-Search-Qt.zip) лежит внутри.
Разворачивает папку в %LOCALAPPDATA%\\Programs\\Excel Search 2 (без прав администратора, как VS Code),
кладёт ярлыки на рабочий стол и в «Пуск», запускает программу. Дальше программа обновляется сама.
Настроек и _backup в архиве нет — повторная установка их не трогает."""

import os
import subprocess
import sys
import threading
import tkinter as tk
import zipfile
from tkinter import messagebox, ttk

APP = "Excel Search 2"
EXE = "Excel Search 2.exe"
TARGET = os.path.join(os.environ.get("LOCALAPPDATA") or os.path.expanduser(r"~\AppData\Local"), "Programs", APP)


def bundled(name):
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, name)


def running():
    """Запущена ли программа ИЗ папки установки (тогда её файлы заняты)."""
    ps = (f"@(Get-Process -Name '{APP}' -ErrorAction SilentlyContinue | "
          f"? {{ $_.Path -like '{TARGET}\\*' }}).Count")
    out = subprocess.run(["powershell", "-NoProfile", "-Command", ps], capture_output=True, text=True,
                         creationflags=0x08000000).stdout.strip()
    return out not in ("", "0")


def shortcuts():
    """Ярлыки на рабочем столе (путь у Windows — бывает перенаправлен) и в «Пуск»."""
    exe = os.path.join(TARGET, EXE).replace("'", "''")
    wd = TARGET.replace("'", "''")
    ps = ("$sh = New-Object -ComObject WScript.Shell;"
          "foreach ($d in @([Environment]::GetFolderPath('Desktop'), [Environment]::GetFolderPath('Programs'))) {"
          f"  $l = $sh.CreateShortcut((Join-Path $d '{APP}.lnk'));"
          f"  $l.TargetPath = '{exe}'; $l.WorkingDirectory = '{wd}'; $l.IconLocation = '{exe},0'; $l.Save() }}")
    subprocess.run(["powershell", "-NoProfile", "-Command", ps], creationflags=0x08000000, timeout=60)


def main():
    root = tk.Tk()
    root.title(f"Установка {APP}")
    root.geometry("420x120")
    root.resizable(False, False)
    if running():
        root.withdraw()
        messagebox.showwarning(f"Установка {APP}", "Excel Search 2 открыт — закройте его и запустите установку снова.")
        root.destroy()
        return
    lbl = ttk.Label(root, text="Подготовка…")
    lbl.pack(padx=16, pady=(18, 8), anchor="w")
    bar = ttk.Progressbar(root, length=388, mode="determinate")
    bar.pack(padx=16)
    state = {"err": None}

    def work():
        try:
            os.makedirs(TARGET, exist_ok=True)
            with zipfile.ZipFile(bundled("Excel-Search-Qt.zip")) as z:
                names = z.infolist()
                for i, info in enumerate(names, 1):
                    z.extract(info, TARGET)
                    if i % 20 == 0 or i == len(names):
                        root.after(0, lambda i=i, n=len(names): (
                            lbl.config(text=f"Распаковка: {i} из {n} файлов"), bar.config(value=100 * i / n)))
            root.after(0, lambda: lbl.config(text="Ярлыки…"))
            shortcuts()
        except Exception as exc:
            state["err"] = exc
        root.after(0, done)

    def done():
        if state["err"]:
            messagebox.showerror(f"Установка {APP}", f"Не установилось:\n{state['err']}", parent=root)
        else:
            subprocess.Popen([os.path.join(TARGET, EXE)], cwd=TARGET)
            messagebox.showinfo(f"Установка {APP}", f"Готово.\n\nПапка: {TARGET}\n"
                                "Ярлык «Excel Search 2» — на рабочем столе и в «Пуск».\n"
                                "Обновления программа предлагает сама.", parent=root)
        root.destroy()

    threading.Thread(target=work, daemon=True).start()
    root.mainloop()


if __name__ == "__main__":
    main()
