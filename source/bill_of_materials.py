"""Ведомость материалов по фактическим вставкам выбранной сборки."""
from datetime import datetime
import json
import math
from pathlib import Path
import re

from kompas_dimensions import array, close, key
from embodiment_processing import select

SECTIONS = ('Сортамент', 'Метизы', 'Прочие изделия', 'Без материала / не распределено')


def normalized(value):
    text = re.sub(r'\s+', ' ', str(value or '')).strip().casefold().replace('ё', 'е')
    text = re.sub(r'(?<=\d),(?=\d)', '.', text)
    return re.sub(r'\s*[х×x]\s*(?=\d)', 'x', text)


def material_type(material):
    text = normalized(material)
    sheet = re.search(r'\b(лист\w*|фанер\w*|дсп|лдсп|мдф|двп|осб|osb|поликарбонат\w*)\b', text)
    linear = re.search(r'\b(труб\w*|профил\w*|угол\w*|швеллер\w*|двутавр\w*|круг\w*|'
                 r'прут\w*|брус\w*|доск\w*|рейк\w*|полос\w*|арматур\w*|квадрат\w*|шестигранник\w*)\b', text)
    if sheet and (not linear or sheet.start() < linear.start()):
        return 'sheet'
    if linear:
        return 'linear'
    return 'unknown'


def number(value):
    return f'{value:.3f}'.rstrip('0').rstrip('.')


def aggregate(items):
    """Не объединяет разные марки материалов или листы разного размера."""
    groups = {}
    for item in items:
        qty = item.get('quantity', 1)
        if not isinstance(qty, int) or qty <= 0:
            raise ValueError('Количество должно быть положительным целым числом')
        material = str(item.get('material') or '').strip()
        category = item.get('category', 'normal')
        name = str(item.get('name') or 'Без наименования').strip()
        marking = str(item.get('marking') or '').strip()
        dims = item.get('dimensions')
        identity = (normalized(name), normalized(marking))
        label = f'{name} — {marking}' if marking else name
        kind = material_type(material)
        if item.get('error'):
            section = SECTIONS[3]
            label += ' [не удалось определить данные]'
            ident = (section, key(item.get('path') or label), identity)
            length, pieces = None, qty
        elif category in ('standard', 'other'):
            section = SECTIONS[1] if category == 'standard' else SECTIONS[2]
            ident = (section, identity, normalized(material))
            if material:
                label += '; ' + material
            length, pieces = None, qty
        elif material and kind in ('linear', 'sheet') and dims:
            if len(dims) != 3 or any(not math.isfinite(float(v)) or v <= 0 for v in dims):
                raise ValueError('Некорректные размеры заготовки')
            section = SECTIONS[0]
            if kind == 'linear':
                ident = (section, kind, normalized(material))
                label = material
                length, pieces = float(dims[0]) * qty, None
            else:
                length_mm, width, thickness = (round(float(v), 3) for v in dims)
                ident = (section, kind, normalized(material), thickness, length_mm, width)
                label = (f'{material}; t={number(thickness)} мм; '
                         f'{number(length_mm)} × {number(width)} мм')
                length, pieces = None, qty
        else:
            section = SECTIONS[3]
            label += f'; материал: {material}' if material else '; материал не задан'
            ident = (section, identity, normalized(material), tuple(dims or ()))
            length, pieces = None, qty
        if ident not in groups:
            groups[ident] = {'section': section, 'label': label, 'length': length,
                            'quantity': pieces, 'members': [], 'sort': ident[1:]}
        else:
            if length is not None:
                groups[ident]['length'] += length
            if pieces is not None:
                groups[ident]['quantity'] += pieces
        groups[ident]['members'].append(item)
    result = []
    for section in SECTIONS:
        group = [row for row in groups.values() if row['section'] == section]
        if section == SECTIONS[0]:
            group.sort(key=lambda r: r['sort'])
        else:
            group.sort(key=lambda r: normalized(r['label']))
        for row in group:
            if row['length'] is not None:
                row['length'] = round(row['length'], 3)
            row.pop('sort', None)
            result.append(row)
    return result


def collect(k, doc, journal):
    """PartsEx(0) содержит все вставки, включая копии. Каждая считается один раз."""
    items = []
    def walk(parent, trail, ancestors):
        try:
            parts = array(parent.PartsEx(0))  # ksAllParts, не ksUniqueParts
            if not parts and trail != 'Сборка':
                raise RuntimeError('Состав подсборки пуст или недоступен')
        except Exception as exc:
            journal.append({'path': trail, 'status': 'error', 'reason': str(exc)})
            return
        for i, obj in enumerate(parts, 1):
            record = {'path': '', 'trace': f'{trail}/{i}', 'quantity': 1,
                      'name': '', 'marking': '', 'material': '', 'category': 'normal'}
            try:
                part = k.api.IPart7(obj)
                if k.api.IFeature7(part).Excluded:
                    journal.append({'path': record['trace'], 'status': 'excluded',
                                    'reason': 'Исключён из расчёта сборки'})
                    continue
                if int(part.LoadState) != 0:
                    if not part.Load(False) or int(part.LoadState) != 0:
                        raise RuntimeError('Компонент не удалось полностью загрузить')
                record.update(path=str(part.FileName or ''), name=str(part.Name or ''),
                              marking=str(part.Marking or ''))
                category, section = k.category(doc, part)
                record.update(category=category, section=section)
                if category != 'normal':
                    # Для покупных изделий геометрия и список исполнений не нужны.
                    # Пустой материал не исключает их из соответствующего раздела.
                    items.append(record)
                    continue
                record['material'] = str(part.Material or '')
                manager = k.api.IEmbodimentsManager(part)
                record['embodiment'] = str(manager.GetCurrentEmbodimentMarking(-1, False) or '')
                if part.Detail:
                    record['local'] = bool(part.IsLocal)
                    items.append(record)
                else:
                    ident = (key(record['path']) if record['path'] else record['trace'],
                             record['embodiment'])
                    if ident in ancestors:
                        raise RuntimeError('Циклическая ссылка подсборки')
                    walk(part, record['trace'], ancestors | {ident})
            except Exception as exc:
                record['error'] = str(exc)
                record['name'] = record['name'] or 'Не удалось прочитать компонент'
                items.append(record)
                journal.append({'path': record['trace'], 'status': 'error', 'reason': str(exc)})
    root = k.top(doc)
    identity = (key(doc.PathName), k.api.IEmbodimentsManager(root).GetCurrentEmbodimentMarking(-1, False))
    walk(root, 'Сборка', {identity})
    return items


def source_dimensions(k, item, mode):
    path = Path(item['path'])
    if item.get('local') or not path.is_absolute() or path.suffix.lower() != '.m3d':
        raise RuntimeError('Нет отдельного исходного файла .m3d для измерения')
    if key(path) in k.opened():
        raise RuntimeError('Закройте отдельное окно исходной детали перед формированием ведомости')
    doc = k.open(path, True)
    manager, original = None, None
    try:
        manager = k.api.IEmbodimentsManager(k.top(doc))
        original = int(manager.CurrentEmbodimentIndex)
        wanted = normalized(item.get('embodiment'))
        count = int(manager.EmbodimentCount)
        if count > 1:
            matches = [i for i in range(count)
                       if normalized(manager.GetEmbodimentMarking(i, -1, False)) == wanted]
            if not wanted or len(matches) != 1:
                raise RuntimeError('Не удалось однозначно сопоставить исполнение вставки с исходным файлом')
            select(manager, matches[0])
            if normalized(manager.GetCurrentEmbodimentMarking(-1, False)) != wanted:
                raise RuntimeError('КОМПАС выбрал другое исполнение исходной детали')
        k.rebuild(doc)
        return k.measure(k.top(doc), mode)
    finally:
        try:
            if manager is not None and original is not None:
                select(manager, original)
        finally:
            close(doc)


def export(folder, rows, journal, assembly, marking):
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
    wb = Workbook()
    ws = wb.active
    ws.title = 'Ведомость'
    errors = sum(r.get('status') == 'error' for r in journal)
    title = f'Ведомость: {Path(assembly).name}; исполнение: {marking or "основное"}'
    if errors:
        title = f'НЕПОЛНЫЙ ОТЧЁТ — ошибок: {errors}. ' + title
    ws.merge_cells('A1:D1')
    ws['A1'] = title
    ws['A1'].font = Font(bold=True, color='B91C1C' if errors else '000000', size=12)
    ws['A1'].alignment = Alignment(wrap_text=True, vertical='center')
    ws.row_dimensions[1].height = 42
    ws.merge_cells('A2:D2')
    ws['A2'] = 'Линейный сортамент — суммарная длина; листы и покупные изделия — количество.'
    ws['A2'].alignment = Alignment(wrap_text=True)
    thin = Side(style='thin', color='000000')
    current = 4
    for section in SECTIONS:
        headers = ('№', section, 'Длина (мм)\nна изделие', 'Кол-во (шт.)\nна изделие')
        for col, value in enumerate(headers, 1):
            cell = ws.cell(current, col, value)
            cell.font = Font(bold=True)
            cell.fill = PatternFill('solid', fgColor='E8EEF5')
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        ws.row_dimensions[current].height = 34
        current += 1
        section_rows = [row for row in rows if row['section'] == section]
        for i, row in enumerate(section_rows, 1):
            for col, value in enumerate((i, row['label'], row['length'], row['quantity']), 1):
                cell = ws.cell(current, col, value)
                cell.alignment = Alignment(horizontal='left' if col == 2 else 'center',
                                           vertical='center', wrap_text=True)
                if isinstance(value, float):
                    cell.number_format = '0.###'
            ws.row_dimensions[current].height = max(30,15*math.ceil(len(row['label'])/65))
            current += 1
        if not section_rows:
            current += 1
        current += 1
    for row in ws.iter_rows(min_row=4, max_row=current-2, max_col=4):
        for cell in row:
            cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for row in ws:
        for cell in row:
            if isinstance(cell.value, str):
                cell.data_type = 's'
    for col, width in (('A',6),('B',76),('C',19),('D',21)):
        ws.column_dimensions[col].width = width
    ws.freeze_panes = 'C5'
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.orientation = 'landscape'
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.print_options.horizontalCentered = True
    ws.print_area = f'A1:D{current-2}'
    output = folder/'Ведомость.xlsx'
    wb.save(output)
    (folder/'Ведомость.json').write_text(json.dumps(
        {'assembly':str(assembly),'embodiment':marking,'errors':errors,'rows':rows,'journal':journal},
        ensure_ascii=False,indent=2),encoding='utf-8')
    return output, title


def preview(rows, title):
    lines = [title, '']
    for section in SECTIONS:
        lines.extend([section.upper(), '№\tНаименование / материал\tДлина, мм\tКол-во'])
        for i,row in enumerate((r for r in rows if r['section']==section),1):
            label = re.sub(r'\s+',' ',row['label'])
            short = label if len(label) <= 75 else label[:72]+'...'
            length = number(row['length']) if row['length'] is not None else ''
            qty = str(row['quantity']) if row['quantity'] is not None else ''
            lines.append(f'{i}\t{short}\t{length}\t{qty}')
            if len(label)>75:
                lines.append('\t'+label)
        lines.append('')
    return '\n'.join(lines)


def run(k, assembly, output, mode='bbox'):
    assembly = Path(assembly).resolve()
    if assembly.suffix.lower() != '.a3d' or not assembly.is_file():
        raise ValueError('Выберите существующую сборку .a3d')
    existing = k.opened().get(key(assembly))
    if existing and existing.Changed:
        raise RuntimeError('Сохраните изменения выбранной сборки перед созданием ведомости')
    doc = existing or k.open(assembly, True)
    folder = Path(output).resolve()/('Ведомость_'+datetime.now().strftime('%Y%m%d_%H%M%S_%f'))
    folder.mkdir(parents=True)
    journal = []
    try:
        marking = str(k.api.IEmbodimentsManager(k.top(doc)).GetCurrentEmbodimentMarking(-1,False) or '')
        print('Ведомость по сборке: '+str(assembly),flush=True)
        items = collect(k,doc,journal)
        cache = {}
        for i,item in enumerate(items,1):
            print(f'  [{i}/{len(items)}] {item["name"]} — {item["marking"]}',flush=True)
            if item.get('error') or item['category'] in ('standard','other'):
                continue
            if material_type(item['material']) == 'unknown':
                continue
            cache_key = (key(item['path']) if item['path'] else '', normalized(item.get('embodiment')))
            try:
                if cache_key not in cache:
                    try:
                        cache[cache_key] = {'dimensions':source_dimensions(k,item,mode)}
                    except Exception as exc:
                        cache[cache_key] = {'error':str(exc)}
                if 'error' in cache[cache_key]:
                    raise RuntimeError(cache[cache_key]['error'])
                item['dimensions'] = cache[cache_key]['dimensions']
            except Exception as exc:
                item['error'] = str(exc)
                journal.append({'path':item['trace'],'status':'error','reason':str(exc)})
        rows = aggregate(items)
        if not items:
            journal.append({'path':str(assembly),'status':'error','reason':'В сборке не найдены компоненты'})
        workbook,title = export(folder,rows,journal,assembly,marking)
        print('Сохранена ведомость: '+str(workbook),flush=True)
        return {'workbook':str(workbook),'text':preview(rows,title),'errors':sum(r.get('status')=='error' for r in journal)}
    finally:
        (folder/'Журнал.json').write_text(json.dumps(journal,ensure_ascii=False,indent=2),encoding='utf-8')
        if existing is None:
            close(doc)
