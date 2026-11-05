# -*- coding: utf-8 -*-
"""КОМПАС-3D v24 / Python 3.10+ / pywin32 / openpyxl.
run: python kompas_dimensions.py [--write] [--mode parameters]
init: python kompas_dimensions.py --init
"""
import argparse
from collections import Counter
from datetime import datetime
import json
import math
import os
from pathlib import Path
import shutil
import sys
import xml.etree.ElementTree as ET

NAMES = ("Длина", "Ширина", "Высота")
VARIABLES = ("BLANK_L", "BLANK_W", "BLANK_H")


def array(value):
    return [] if value is None else list(value) if isinstance(value, (tuple, list)) else [value]


def dimensions(values, digits=3):
    values = [float(v) for v in values]
    if len(values) != 3 or any(not math.isfinite(v) or v <= 0 for v in values):
        raise ValueError(f"Некорректные размеры: {values}")
    result = tuple(round(v, digits) for v in sorted(values, reverse=True))
    if min(result) <= 0:
        raise ValueError("Размер меньше точности округления")
    return result


def parse_section(xml):
    """Имена разделов — метаданные КОМПАС, не имена файлов."""
    if not xml:
        return ""
    root = ET.fromstring(xml)
    for node in root.iter():
        if node.get("id") == "sectionName":
            return node.get("value", "").strip()
    if root.tag == "document" and not list(root):
        return ""  # раздел явно не задан
    raise ValueError("Неизвестный формат раздела спецификации")


class Kompas:
    def __init__(self):
        import pythoncom
        from win32com.client import GetActiveObject, Dispatch, gencache
        self.api = gencache.EnsureModule(
            "{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0)
        obj = None
        failures = []
        for connect, progid in ((GetActiveObject, "Kompas.Application.7"),
                                (GetActiveObject, "Kompas.Application.5"),
                                (Dispatch, "Kompas.Application.7"),
                                (Dispatch, "Kompas.Application.5")):
            try:
                obj = connect(progid)
                if progid.endswith(".5"):
                    obj = obj.ksGetApplication7()
                break
            except pythoncom.com_error as exc:
                failures.append(str(exc))
                obj = None
        if obj is None:
            raise RuntimeError("Не удалось подключиться к КОМПАС-3D или запустить его. "
                               "Проверьте установку КОМПАС v24. " + "; ".join(failures))
        self.app = self.api.IApplication(obj._oleobj_.QueryInterface(
            self.api.IApplication.CLSID, pythoncom.IID_IDispatch))
        self.mng = self.api.IPropertyMng(self.app)
        self.app.Visible = True

    def top(self, doc):
        return self.api.IKompasDocument3D(doc).TopPart

    def props(self, doc):
        return [self.api.IProperty(p) for p in array(self.mng.GetProperties(doc))]

    def prop(self, doc, ident):
        # float обязателен: int в GetProperty означает индекс, а не ID!
        p = self.mng.GetProperty(doc, float(ident))
        return self.api.IProperty(p) if p else None

    def read(self, part, prop):
        result = self.api.IPropertyKeeper(part).GetPropertyValue(prop, None, False, False)
        if not isinstance(result, tuple) or len(result) != 3 or not result[0]:
            raise RuntimeError(f"Не прочитано свойство {prop.Name}: {result!r}")
        return result[1]

    def category(self, doc, part):
        if bool(part.Standard):
            return "standard", "IPart7.Standard"
        p = self.prop(doc, 20.0)
        if not p:
            raise RuntimeError("Недоступно системное свойство 20: Раздел спецификации")
        keeper = self.api.IPropertyKeeper(part)
        if keeper.IsComplexPropertyValue(p):
            result = keeper.GetComplexPropertyValue(p, False)
            if not isinstance(result, tuple) or len(result) != 2:
                raise RuntimeError(f"Неизвестный ответ GetComplexPropertyValue: {result!r}")
            name = parse_section(result[0])
        else:
            value = self.read(part, p)
            if value is not None and not isinstance(value, str):
                raise RuntimeError(f"Неизвестный формат раздела: {value!r}")
            name = (value or "").strip()
        if name.casefold() == "стандартные изделия":
            return "standard", name
        if name.casefold() == "прочие изделия":
            return "other", name
        return "normal", name

    def measure(self, part, mode):
        if mode == "bbox":
            r = part.GetGabarit(False, False, 0., 0., 0., 0., 0., 0.)
            if not isinstance(r, tuple) or len(r) != 7 or not r[0]:
                raise RuntimeError(f"GetGabarit не выдал габариты: {r!r}")
            _, x1, y1, z1, x2, y2, z2 = r
            return dimensions((x2 - x1, y2 - y1, z2 - z1))
        feature = self.api.IFeature7(part)
        vals = []
        for name in VARIABLES:
            var = feature.Variable(False, True, name)
            if not var:
                raise RuntimeError(f"Нет переменной заготовки {name}")
            var = self.api.IVariable7(var)
            if var.IsAngular:
                raise RuntimeError(f"{name} — угловая переменная")
            vals.append(var.Value)
        return dimensions(vals)

    def opened(self):
        docs = self.app.Documents
        return {key(d.PathName): d for d in
                (docs.Item(i) for i in range(docs.Count)) if d.PathName}

    def open(self, path, readonly):
        doc = self.app.Documents.Open(str(path), False, readonly)
        if not doc:
            raise RuntimeError(f"Не удалось открыть {path}")
        return doc


def key(path):
    return os.path.normcase(os.path.realpath(path))


def save(doc):
    # Save возвращает void, проверка `if not doc.Save()` была бы ошибкой.
    doc.Save()
    if doc.Changed:
        raise RuntimeError("После Save документ остался изменённым")


def close(doc):
    if not doc.Close(0):  # kdDoNotSaveChanges
        raise RuntimeError("Не удалось закрыть документ")


def init_properties(k, folder):
    folder.mkdir(parents=True, exist_ok=True)
    seed = folder / "dimensions_seed.m3d"
    registry = folder / "properties.json"
    if seed.exists() or registry.exists():
        raise RuntimeError("Источник определений уже существует; повторно создавать его нельзя")
    doc = k.app.Documents.Add(4, False)  # ksDocumentPart
    try:
        ids = {}
        for name in NAMES:
            if any(p.Name == name for p in k.props(doc)):
                raise RuntimeError(f"Шаблон уже содержит свойство {name}")
            p = k.api.IProperty(k.mng.AddProperty(doc, None))
            p.Name = name
            p.DataType = 2  # ksPropertyDataTypeDouble: число в миллиметрах
            p.Comment = "Автоматические размеры, числовое значение в мм"
            if not p.Update():
                raise RuntimeError(f"Не создано свойство {name}")
            ids[name] = float(p.Id)  # Id доступен только для чтения
        doc.SaveAs(str(seed.resolve()))
        if doc.Changed or not seed.exists():
            raise RuntimeError("Не сохранён источник определений")
        registry.write_text(json.dumps(ids, ensure_ascii=False, indent=2), encoding="utf-8")
    finally:
        close(doc)
    print(f"Созданы {seed} и {registry}. Храните их вместе и используйте во всех сборках.")


def write_values(k, doc, part, prototypes, values):
    # Проверяем весь набор до первого изменения.
    existing = k.props(doc)
    selected = []
    for name in NAMES:
        proto = prototypes[name]
        p = k.prop(doc, proto.Id)
        collisions = [v for v in existing if v.Name == name and v.Id != proto.Id]
        if collisions:
            raise RuntimeError(f"Конфликт: {name} уже существует с другим ID")
        if p and (p.Name != name or p.DataType != 2 or p.ReadOnly or p.IsVariable):
            raise RuntimeError(f"Конфликт определения/формула/защита свойства {name}")
        selected.append(p)
    keeper = k.api.IPropertyKeeper(part)
    for name, p, value in zip(NAMES, selected, values):
        if not p:
            p = k.api.IProperty(k.mng.AddProperty(doc, prototypes[name]))
        if float(p.Id) != float(prototypes[name].Id):
            raise RuntimeError(f"КОМПАС изменил ID при копировании {name}")
        if not keeper.SetPropertyValue(p, float(value), False):
            raise RuntimeError(f"Не записано свойство {name}")
        if not math.isclose(float(k.read(part, p)), value, abs_tol=1e-8):
            raise RuntimeError(f"Контроль чтения {name} не пройден")
    if not part.Update():
        raise RuntimeError("Не обновлены свойства верхнего компонента")


def collect(k, root, log):
    records = {}
    def walk(parent, trace, stack, inherited=None):
        parts = parent.Parts
        for i in range(parts.Count):
            part = parts.Part(i)
            label = f"{trace}/{i + 1}"
            raw = str(part.FileName or "")
            rec = {"path": raw, "trace": label, "status": "", "reason": ""}
            try:
                if raw and not os.path.isabs(raw):
                    raise RuntimeError("API вернул относительный путь: нельзя надёжно найти источник")
                cat, section = k.category(root, part)
                cat = inherited or cat
                rec.update(section=section, category=cat)
                if raw:
                    group = records.setdefault(key(raw), {"path": raw, "items": [], "blocked": False})
                    group["items"].append(rec)
                    group["blocked"] |= cat != "normal"
                if cat != "normal":
                    rec["status"] = cat
                if not part.Detail:
                    # Обходим также содержимое исключённой подсборки, наследуя запрет.
                    ident = key(raw) if raw else label
                    if ident in stack:
                        raise RuntimeError("Циклическая ссылка сборки")
                    if not rec["status"]:
                        rec["status"] = "assembly"
                    walk(part, label, stack | {ident}, cat if cat != "normal" else None)
                elif not rec["status"] and (part.IsLocal or Path(raw).suffix.lower() != ".m3d"):
                    rec.update(status="unsupported", reason="Нет отдельного исходного файла .m3d")
            except Exception as exc:
                rec.update(status="error", reason=str(exc))
                if raw and key(raw) in records:
                    records[key(raw)]["blocked"] = True
            log.append(rec)
    walk(k.top(root), "Сборка", {key(root.PathName)})
    return records


def export(folder, rows, log, counts):
    from openpyxl import Workbook
    from openpyxl.styles import Font
    wb = Workbook()
    ws = wb.active
    ws.title = "Габариты"
    ws.append(["Файл", "Обозначение", "Наименование", "Количество", "Длина, мм",
               "Ширина, мм", "Высота, мм", "Метод", "Запись свойств"])
    for row in rows:
        ws.append(row)
    journal = wb.create_sheet("Журнал")
    journal.append(["Вставка", "Файл", "Раздел", "Статус", "Причина"])
    for r in log:
        journal.append([r.get(n, "") for n in ("trace", "path", "section", "status", "reason")])
    summary = wb.create_sheet("Итоги")
    summary.append(["Показатель", "Значение"])
    for name, value in counts.items():
        summary.append([name, value])
    for sheet in wb:
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = sheet.dimensions
        for cell in sheet[1]:
            cell.font = Font(bold=True)
        for column in sheet.columns:
            sheet.column_dimensions[column[0].column_letter].width = min(65, max(16,
                max(len(str(c.value or "")) for c in column) + 2))
        for row in sheet.iter_rows():
            for cell in row:
                if isinstance(cell.value, str):
                    cell.data_type = "s"  # даже строка, начинающаяся с '=', не формула
                elif isinstance(cell.value, float):
                    cell.number_format = "0.###"
    wb.save(folder / "dimensions.xlsx")
    (folder / "journal.json").write_text(json.dumps(
        {"summary": counts, "items": log}, ensure_ascii=False, indent=2), encoding="utf-8")


def folder_groups(folder, excluded, log):
    """Снимок списка файлов до открытия моделей, без переходов по ссылкам."""
    folder = Path(folder).resolve()
    if not folder.is_dir():
        raise ValueError(f"Папка не найдена: {folder}")
    excluded = [Path(p).resolve() for p in excluded]
    def ignored(path):
        path = Path(path).resolve()
        return any(path == base or base in path.parents for base in excluded)
    if ignored(folder):
        raise ValueError("Папка моделей не должна находиться внутри папки отчётов или определений")
    def onerror(exc):
        log.append({"path": str(exc.filename or folder), "status": "scan_error", "reason": str(exc)})
    groups = {}
    for current, dirs, files in os.walk(folder, followlinks=False, onerror=onerror):
        retained = []
        for name in sorted(dirs, key=str.casefold):
            child = Path(current) / name
            try:
                attrs = getattr(child.lstat(), "st_file_attributes", 0)
                if not ignored(child) and not child.is_symlink() and not (attrs & 0x400):
                    retained.append(name)
            except OSError as exc:
                onerror(exc)
        dirs[:] = retained
        for name in sorted(files, key=str.casefold):
            path = Path(current) / name
            if path.suffix.lower() != ".m3d" or path.is_symlink() or ignored(path):
                continue
            try:
                if not path.is_file():
                    raise OSError(f"Файл недоступен: {path}")
                rec = {"path": str(path), "trace": str(path.relative_to(folder)),
                       "status": "", "reason": ""}
                groups.setdefault(key(path), {"path": str(path), "items": [rec], "blocked": False})
                log.append(rec)
            except OSError as exc:
                onerror(exc)
    return groups


def run(k, args):
    input_folder = getattr(args, "folder", None)
    root = None if input_folder else k.app.ActiveDocument
    if not input_folder:
        if not root or int(root.DocumentType) != 5:
            raise RuntimeError("Активным документом должна быть сборка .a3d")
        if not root.PathName or root.Changed:
            raise RuntimeError("Сначала сохраните сборку")
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    folder = Path(args.output).resolve() / stamp
    folder.mkdir(parents=True)
    log, rows = [], []
    seed_doc = None
    prototypes = {}
    try:
        if args.write:
            base = Path(args.schema).resolve()
            ids = json.loads((base / "properties.json").read_text(encoding="utf-8"))
            seed_doc = k.open(base / "dimensions_seed.m3d", True)
            for name in NAMES:
                p = k.prop(seed_doc, ids[name])
                if not p or p.Name != name or p.DataType != 2 or p.IsVariable:
                    raise RuntimeError(f"Повреждён источник определений {name}")
                prototypes[name] = p
        groups = folder_groups(input_folder, (args.output, args.schema), log) if input_folder else collect(k, root, log)
        print(f"Найдено уникальных файлов: {len(groups)}", flush=True)
        if not groups:
            print("Файлы .m3d не найдены. Проверьте выбранную папку и журнал чтения папок.", flush=True)
        for index, group in enumerate(groups.values(), 1):
            pending = [r for r in group["items"] if not r["status"]]
            if not pending:
                continue
            if group["blocked"]:
                for r in pending:
                    r.update(status="conflict", reason="У того же файла есть исключённая или ошибочная вставка")
                continue
            path = Path(group["path"])
            doc, owned = None, False
            try:
                print(f"[{index}/{len(groups)}] Деталь: {path}", flush=True)
                already = k.opened().get(key(path))
                if already and (already.Changed or args.write):
                    raise RuntimeError("Закройте отдельное окно детали перед записью; сохраните изменения перед измерением")
                doc = already or k.open(path, not args.write)
                owned = already is None
                part = k.top(doc)
                cat, section = k.category(doc, part)
                if cat != "normal":
                    for r in pending:
                        r.update(status=cat, section=section, reason="Категория исходного файла")
                    print(f"  Пропущено: {section or cat}", flush=True)
                    continue
                if args.write and doc.ReadOnly:
                    raise RuntimeError("Файл открыт только для чтения")
                count = k.api.IEmbodimentsManager(part).EmbodimentCount
                if count > 1:
                    raise RuntimeError("Файл содержит несколько исполнений; требуется отдельная обработка исполнений")
                values = k.measure(part, args.mode)
                marking = str(part.Marking or "")
                name = str(part.Name or "")
                saved = "не запрошена"
                if args.write:
                    backup = folder / "backups" / f"{index:06d}" / path.name
                    backup.parent.mkdir(parents=True)
                    shutil.copy2(path, backup)
                    write_values(k, doc, part, prototypes, values)
                    save(doc)
                    # Закрытие и повторное открытие проверяет значения именно на диске.
                    close(doc)
                    doc = None
                    doc = k.open(path, True)
                    reread = k.top(doc)
                    for n, value in zip(NAMES, values):
                        p = k.prop(doc, prototypes[n].Id)
                        if not p or not math.isclose(float(k.read(reread, p)), value, abs_tol=1e-8):
                            raise RuntimeError(f"Проверка после сохранения не пройдена: {n}; копия: {backup}")
                    saved = "сохранены и проверены"
                rows.append([str(path), marking, name, len(pending), *values, args.mode, saved])
                for r in pending:
                    r.update(status="processed", reason=saved)
            except Exception as exc:
                print(f"  Ошибка: {exc}", flush=True)
                for r in pending:
                    r.update(status="error", reason=str(exc))
            finally:
                if doc is not None and owned:
                    try:
                        close(doc)
                    except Exception as exc:
                        log.append({"path": str(path), "status": "close_error", "reason": str(exc)})
        tally = Counter(r["status"] for r in log)
        counts = {
            "Обработано уникальных файлов": len(rows),
            "Обработано файлов" if input_folder else "Обработано вставок": tally["processed"],
            "Пропущено стандартных изделий": tally["standard"],
            "Пропущено прочих изделий": tally["other"],
            "Ошибок обработки": tally["error"],
            "Конфликтов категорий (вставок)": tally["conflict"],
            "Без отдельного файла .m3d (вставок)": tally["unsupported"],
            "Ошибок закрытия": tally["close_error"],
            "Ошибок чтения папок": tally["scan_error"],
            "Источник": str(input_folder or root.PathName),
            "Метод": args.mode,
        }
        export(folder, rows, log, counts)
        print(json.dumps(counts, ensure_ascii=False, indent=2))
        print(f"Результаты: {folder}")
        return 2 if not groups or any(tally[s] for s in ("error", "conflict", "unsupported", "close_error", "scan_error")) else 0
    finally:
        # JSON сохраняется даже при сбое экспорта Excel.
        (folder / "processing.json").write_text(json.dumps(log, ensure_ascii=False, indent=2), encoding="utf-8")
        if seed_doc is not None:
            close(seed_doc)
        if root is not None:
            root.Active = True


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--init", action="store_true", help="Однократно создать общие определения")
    parser.add_argument("--write", action="store_true", help="Записать свойства и сохранить детали")
    parser.add_argument("--mode", choices=("bbox", "parameters"), default="bbox")
    parser.add_argument("--schema", default="dimension_properties")
    parser.add_argument("--output", default="dimension_reports")
    parser.add_argument("--folder", help="Папка с деталями .m3d, включая вложенные папки")
    args = parser.parse_args()
    import pythoncom
    pythoncom.CoInitialize()
    try:
        k = Kompas()
        if args.init:
            init_properties(k, Path(args.schema).resolve())
            return 0
        return run(k, args)
    finally:
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception as exc:
        print(f"ОШИБКА: {exc}", file=sys.stderr)
        sys.exit(1)
