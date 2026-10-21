"""Заставка при запуске: знак SM разгоняется, вращается и замирает анфас.

Взято из dahua-recovery (знак автора SM). Кадры готовит tools/make_splash.py там же (assets/splash/NNN.png, 60 кадров/с);
здесь их только проигрываем средствами tkinter, без сторонних библиотек.

Две фазы: Splash(root) сразу показывает неподвижный первый кадр — пока
строится главное окно (~1 с), экран не пустой; play() запускает вращение.
"""
import glob
import os
import time
import tkinter as tk

import sys

def _bundle_dir():
    """Кадры лежат внутри onefile (_MEIPASS) или рядом со скриптом."""
    return getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))


FPS = 60                    # как в tools/make_splash.py
FADE_IN, HOLD, FADE_OUT = 0.12, 0.1, 0.25   # секунды


class Splash:
    def __init__(self, root):
        self.files = sorted(glob.glob(
            os.path.join(_bundle_dir(), "assets", "splash", "*.png")))
        if not self.files:
            self.win = None
            return
        self.win = win = tk.Toplevel(root)
        win.overrideredirect(True)
        win.attributes("-topmost", True)
        win.attributes("-alpha", 0.0)
        img = tk.PhotoImage(file=self.files[0])
        w, h = img.width(), img.height()
        x = (win.winfo_screenwidth() - w) // 2
        y = (win.winfo_screenheight() - h) // 2
        win.geometry(f"{w}x{h}+{x}+{y}")
        self.label = tk.Label(win, image=img, bd=0)
        self.label.pack()
        self.label.image = img
        # проявление — прямо здесь: цикл событий ещё не запущен, после нас
        # конструктор главного окна займёт поток почти на секунду
        t0 = time.perf_counter()
        while (t := time.perf_counter() - t0) < FADE_IN:
            win.attributes("-alpha", t / FADE_IN)
            win.update()
        win.attributes("-alpha", 1.0)
        win.update()

    def play(self, on_done):
        """Проиграть вращение и по окончании вызвать on_done."""
        if self.win is None:
            on_done()
            return
        files, label, win = self.files, self.label, self.win
        t0 = time.perf_counter()
        end_anim = len(files) / FPS
        st = {"shown": 0, "revealed": False, "fade_at": None}

        def tick():
            now = time.perf_counter()
            t = now - t0
            if t < end_anim:
                # кадр по часам, а не по счётчику: притормозил цикл — кадры
                # пропускаются, длительность не растягивается
                i = min(int(t * FPS), len(files) - 1)
                if i != st["shown"]:
                    label.image = tk.PhotoImage(file=files[i])
                    label.configure(image=label.image)
                    st["shown"] = i
            elif not st["revealed"]:
                # знак встал — последний кадр и открываем главное окно ПОД
                # заставкой (она topmost): первая отрисовка окна длится ~1 с,
                # пусть идёт, пока знак ещё на экране
                label.image = tk.PhotoImage(file=files[-1])
                label.configure(image=label.image)
                st["revealed"] = True
                on_done()
                win.lift()
            elif st["fade_at"] is None:
                st["fade_at"] = now + HOLD      # окно отрисовано — отсчёт паузы
            else:
                alpha = 1 - (now - st["fade_at"]) / FADE_OUT
                if alpha <= 0:
                    win.destroy()
                    return
                win.attributes("-alpha", min(1.0, alpha))
            win.after(5, tick)

        win.after(0, tick)
