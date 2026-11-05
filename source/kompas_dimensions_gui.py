# -*- coding: utf-8 -*-
"""Нативное окно Windows: рекурсивная обработка папок с деталями .m3d."""
import contextlib
import io
import json
import os
from pathlib import Path
import queue
import sys
import threading
import traceback
from types import SimpleNamespace

import pythoncom
import win32api
import win32con as C
import win32gui as G
import win32com
from win32com.client import gencache
import win32com.gen_py
from win32com.shell import shell, shellcon

import kompas_dimensions as core

APP_NAME = "Габариты деталей 1.3 — обработка исполнений — КОМПАС-3D v24"
DATA = Path(os.environ.get("LOCALAPPDATA", str(Path.home()))) / "KompasDimensions"
SCHEMA = DATA / "properties"
UPDATE = C.WM_APP + 1


def prepare_cache():
    """Генерация интерфейсов COM в доступной для записи папке."""
    cache = DATA / "gen_py"
    cache.mkdir(parents=True, exist_ok=True)
    (cache / "__init__.py").touch()
    win32com.__gen_path__ = str(cache)
    win32com.gen_py.__path__ = [str(cache)]
    gencache.is_readonly = False


class LogWriter(io.TextIOBase):
    def __init__(self, callback):
        self.callback = callback
    def write(self, value):
        if value:
            self.callback(value)
        return len(value)
    def flush(self):
        pass


class Window:
    def __init__(self):
        self.busy = False
        self.events = queue.Queue()
        self.controls = {}
        self.mode = "bbox"
        self.output = Path(shell.SHGetFolderPath(
            0, shellcon.CSIDL_PERSONAL, None, 0)) / "Габариты КОМПАС"
        instance = win32api.GetModuleHandle(None)
        wc = G.WNDCLASS()
        wc.hInstance = instance
        wc.lpszClassName = "KompasDimensionsWindow"
        wc.lpfnWndProc = self.proc
        wc.hCursor = G.LoadCursor(0, C.IDC_ARROW)
        wc.hbrBackground = C.COLOR_BTNFACE + 1
        G.RegisterClass(wc)
        style = C.WS_OVERLAPPED | C.WS_CAPTION | C.WS_SYSMENU | C.WS_MINIMIZEBOX
        self.hwnd = G.CreateWindow(wc.lpszClassName, APP_NAME, style,
                                  C.CW_USEDEFAULT, C.CW_USEDEFAULT, 800, 735,
                                  0, 0, instance, None)
        self.font = G.GetStockObject(17)  # DEFAULT_GUI_FONT
        self.add("STATIC", "1. Выберите папку с деталями .m3d. Вложенные папки проверяются автоматически.",
                 20, 18, 750, 23)
        self.add("STATIC", "2. Нажмите «Записать свойства + Excel». Открытая сборка не требуется.",
                 20, 44, 750, 23)
        self.add("STATIC", "Папка с деталями:", 20, 80, 750, 23)
        self.input_edit = self.add("EDIT", "", 20, 106, 580, 27,
                                  303, C.WS_BORDER | C.ES_AUTOHSCROLL)
        self.add("BUTTON", "Выбрать папку...", 610, 105, 145, 30, 104)
        self.add("BUTTON", "Габариты готовой модели", 20, 157, 350, 28,
                 201, C.BS_RADIOBUTTON)
        self.add("BUTTON", "Размеры заготовки по переменным", 380, 157, 380, 28,
                 202, C.BS_RADIOBUTTON)
        G.SendMessage(self.controls[201], C.BM_SETCHECK, C.BST_CHECKED, 0)
        self.add("STATIC", "Для заготовки нужны переменные BLANK_L, BLANK_W, BLANK_H (мм).",
                 20, 192, 750, 23)
        self.add("STATIC", "Папка отчётов:", 20, 229, 150, 22)
        self.path_edit = self.add("EDIT", str(self.output), 20, 254, 735, 26,
                                  301, C.WS_BORDER | C.ES_AUTOHSCROLL)
        self.add("BUTTON", "Сформировать Excel", 20, 299, 220, 38, 101)
        self.add("BUTTON", "Записать свойства + Excel", 250, 299, 285, 38, 102)
        self.add("BUTTON", "Открыть отчёты", 545, 299, 210, 38, 103)
        self.add("STATIC", "Стандартные и прочие изделия пропускаются. Размеры — числа в мм.",
                 20, 355, 750, 24)
        self.add("STATIC", "Журнал обработки:", 20, 389, 750, 23)
        self.log = self.add("EDIT", "Выберите папку с деталями.\r\n", 20, 417, 735, 230,
                             302, C.WS_BORDER | C.ES_MULTILINE | C.ES_READONLY |
                             C.WS_VSCROLL | C.ES_AUTOVSCROLL)
        G.SendMessage(self.log, C.EM_SETLIMITTEXT, 2000000, 0)
        self.add("STATIC", "Все исполнения измеряются отдельно. Перед записью создаются резервные копии.",
                 20, 660, 750, 23)
        G.ShowWindow(self.hwnd, C.SW_SHOW)
        G.UpdateWindow(self.hwnd)

    def add(self, cls, text, x, y, width, height, ident=0, extra=0):
        style = C.WS_CHILD | C.WS_VISIBLE | extra
        if cls in ("BUTTON", "EDIT"):
            style |= C.WS_TABSTOP
        hwnd = G.CreateWindow(cls, text, style, x, y, width, height,
                              self.hwnd, ident, win32api.GetModuleHandle(None), None)
        G.SendMessage(hwnd, C.WM_SETFONT, self.font, True)
        if ident:
            self.controls[ident] = hwnd
        return hwnd

    def post(self, kind, value):
        self.events.put((kind, value))
        G.PostMessage(self.hwnd, UPDATE, 0, 0)

    def append(self, text):
        text = text.replace("\r\n", "\n").replace("\n", "\r\n")
        G.SendMessage(self.log, C.EM_SETSEL, -1, -1)
        G.SendMessage(self.log, C.EM_REPLACESEL, False, text)

    def worker(self, write, output, mode, input_folder):
        pythoncom.CoInitialize()
        try:
            writer = LogWriter(lambda text: self.post("log", text))
            with contextlib.redirect_stdout(writer), contextlib.redirect_stderr(writer):
                prepare_cache()
                k = core.Kompas()
                if write:
                    seed = SCHEMA / "dimensions_seed.m3d"
                    registry = SCHEMA / "properties.json"
                    if seed.exists() != registry.exists():
                        raise RuntimeError("Повреждён источник определений свойств: " + str(SCHEMA))
                    if not seed.exists():
                        print("Создание общих определений свойств...")
                        core.init_properties(k, SCHEMA)
                print("Папка моделей: " + str(input_folder))
                args = SimpleNamespace(write=write, mode=mode,
                                       schema=str(SCHEMA), output=str(output), folder=str(input_folder))
                code = core.run(k, args)
                print("Завершено." if code == 0 else
                      "Завершено с пропусками или ошибками. См. журнал в отчёте.")
        except Exception as exc:
            self.post("log", "\nОШИБКА: " + str(exc) + "\n")
            DATA.mkdir(parents=True, exist_ok=True)
            (DATA / "last_error.txt").write_text(traceback.format_exc(), encoding="utf-8")
        finally:
            pythoncom.CoUninitialize()
            self.post("done", None)

    def start(self, write):
        if self.busy:
            return
        raw_input = G.GetWindowText(self.input_edit).strip().strip('"')
        input_folder = Path(raw_input).expanduser()
        if not raw_input or not input_folder.is_absolute() or not input_folder.is_dir():
            G.MessageBox(self.hwnd, "Выберите существующую папку с деталями .m3d.", APP_NAME, C.MB_ICONWARNING)
            return
        output = Path(G.GetWindowText(self.path_edit).strip()).expanduser()
        if not output.is_absolute():
            G.MessageBox(self.hwnd, "Укажите полный путь к папке отчётов.", APP_NAME, C.MB_ICONWARNING)
            return
        self.output = output
        self.busy = True
        for ident in (101, 102, 104, 201, 202, 301, 303):
            G.EnableWindow(self.controls[ident], False)
        self.append("\n" + ("Запись свойств и Excel" if write else "Только Excel") + "\n")
        threading.Thread(target=self.worker, args=(write, output, self.mode, input_folder), daemon=True).start()

    def choose_folder(self):
        pythoncom.CoInitialize()
        try:
            result = shell.SHBrowseForFolder(self.hwnd, None,
                "Выберите папку с деталями .m3d (включая вложенные папки)",
                shellcon.BIF_RETURNONLYFSDIRS | shellcon.BIF_NEWDIALOGSTYLE)
            if not result:
                return
            chosen = result[0]
            if chosen:
                path = shell.SHGetPathFromIDList(chosen)
                if isinstance(path, bytes):
                    path = path.decode("mbcs")
                G.SetWindowText(self.input_edit, path)
        except Exception as exc:
            G.MessageBox(self.hwnd, "Не удалось открыть выбор папки. Введите путь вручную.\n" + str(exc),
                         APP_NAME, C.MB_ICONWARNING)
        finally:
            pythoncom.CoUninitialize()

    def proc(self, hwnd, msg, wp, lp):
        if msg == C.WM_COMMAND:
            ident = wp & 0xffff
            if ident in (201, 202):
                self.mode = "bbox" if ident == 201 else "parameters"
                for choice in (201, 202):
                    G.SendMessage(self.controls[choice], C.BM_SETCHECK,
                                  C.BST_CHECKED if choice == ident else C.BST_UNCHECKED, 0)
            elif ident in (101, 102):
                self.start(ident == 102)
            elif ident == 104:
                self.choose_folder()
            elif ident == 103:
                try:
                    self.output.mkdir(parents=True, exist_ok=True)
                    os.startfile(str(self.output))
                except Exception as exc:
                    G.MessageBox(hwnd, str(exc), APP_NAME, C.MB_ICONERROR)
            return 0
        if msg == UPDATE:
            while not self.events.empty():
                kind, value = self.events.get_nowait()
                if kind == "log":
                    self.append(value)
                elif kind == "done":
                    self.busy = False
                    for ident in (101, 102, 104, 201, 202, 301, 303):
                        G.EnableWindow(self.controls[ident], True)
            return 0
        if msg == C.WM_CLOSE:
            if self.busy:
                G.MessageBox(hwnd, "Дождитесь завершения обработки деталей.", APP_NAME, C.MB_ICONINFORMATION)
            else:
                G.DestroyWindow(hwnd)
            return 0
        if msg == C.WM_DESTROY:
            G.PostQuitMessage(0)
            return 0
        return G.DefWindowProc(hwnd, msg, wp, lp)


def self_test(folder):
    from openpyxl import load_workbook
    import embodiment_processing  # проверяет наличие модуля исполнений в .exe
    folder = Path(folder)
    folder.mkdir(parents=True, exist_ok=True)
    prepare_cache()
    assert core.dimensions((100, 2500, 100)) == (2500, 100, 100)
    core.export(folder, [["test.m3d", "=TEST", "Брус", 1, 2500, 100, 100, "bbox", "нет"]], [], {"Тест": 1})
    wb = load_workbook(folder / "dimensions.xlsx", data_only=False)
    assert wb.active["B2"].value == "=TEST"
    assert not any(c.data_type == 'f' for s in wb for row in s for c in row)
    (folder / "self_test.json").write_text(json.dumps({"ok": True, "python": sys.version,
        "frozen": bool(getattr(sys, "frozen", False))}, ensure_ascii=False, indent=2), encoding="utf-8")


if __name__ == "__main__":
    try:
        if len(sys.argv) > 2 and sys.argv[1] == "--self-test":
            self_test(sys.argv[2])
        else:
            app = Window()
            if "--smoke-gui" in sys.argv:
                timer = threading.Timer(3, lambda: G.PostMessage(app.hwnd, C.WM_CLOSE, 0, 0))
                timer.start()
            G.PumpMessages()
    except Exception:
        DATA.mkdir(parents=True, exist_ok=True)
        text = traceback.format_exc()
        (DATA / "startup_error.txt").write_text(text, encoding="utf-8")
        G.MessageBox(0, text, APP_NAME, C.MB_ICONERROR)
        sys.exit(1)
