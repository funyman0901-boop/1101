"""Измерение и запись свойств исполнений одного файла КОМПАС-3D."""
import math
from pathlib import Path
import shutil


def select(manager, index):
    if int(manager.CurrentEmbodimentIndex) != index:
        if not manager.SetCurrentEmbodiment(int(index)):
            raise RuntimeError(f"Не удалось переключить исполнение {index}")
    if int(manager.CurrentEmbodimentIndex) != index:
        raise RuntimeError(f"КОМПАС оставил другое исполнение вместо {index}")
    embodiment = manager.CurrentEmbodiment
    if embodiment is None:
        raise RuntimeError(f"Нет хранилища исполнения {index}")
    return embodiment


def snapshot(k, doc, storage, prototypes):
    """Значения исключённого исполнения; незаданные значения равны None."""
    result = []
    for proto in prototypes.values():
        prop = k.prop(doc, proto.Id)
        if prop is None:
            result.append(None)
            continue
        value = k.api.IPropertyKeeper(storage).GetPropertyValue(prop, None, False, False)
        if not isinstance(value, tuple) or len(value) != 3:
            raise RuntimeError("Не удалось прочитать свойства исключённого исполнения")
        if not value[0] and value[1] is not None:
            raise RuntimeError("Ошибка чтения свойств исключённого исполнения")
        if not value[0] and value[2]:
            raise RuntimeError("Не разрешён источник свойств исключённого исполнения")
        result.append(value[1] if value[0] else None)
    return result


def verify(k, doc, manager, measured, skipped, prototypes):
    for item in measured:
        source = select(manager, item['index'])
        for name, expected in zip(prototypes, item['values']):
            prop = k.prop(doc, prototypes[name].Id)
            if not prop or not math.isclose(float(k.read(source, prop)), expected, abs_tol=1e-8):
                raise RuntimeError(f"Размер {name} не сохранён отдельно для исполнения {item['label']}")
    for item in skipped:
        source = select(manager, item['index'])
        if snapshot(k, doc, source, prototypes) != item['snapshot']:
            raise RuntimeError("Размеры у исключённого исполнения изменились по наследованию: " + item['label'])


def process(k, doc, path, args, prototypes, backup, trace):
    """Владеет doc: закрывает его при любом исходе. Сохранение один раз на файл."""
    from kompas_dimensions import write_values, save, close
    records, measured, skipped = [], [], []
    manager = None
    original = None
    saved = False
    try:
        manager = k.api.IEmbodimentsManager(k.top(doc))
        original = int(manager.CurrentEmbodimentIndex)
        count = int(manager.EmbodimentCount)
        if count < 2:
            raise RuntimeError("Изменился список исполнений файла")
        for index in range(count):
            rec = {'path': str(path), 'trace': trace, 'embodiment': f'Индекс {index}',
                   'status': '', 'reason': '', 'section': ''}
            records.append(rec)
            try:
                source = select(manager, index)
                marking = str(manager.GetCurrentEmbodimentMarking(-1, False) or '')
                label = marking or str(source.Name or f'Индекс {index}')
                rec['embodiment'] = label
                print(f"  Исполнение [{index + 1}/{count}]: {label}", flush=True)
                part = k.top(doc)
                category, section = k.category(doc, part, source=source)
                rec['section'] = section or 'Не задан'
                if category != 'normal':
                    rec.update(status=category, reason='Категория исполнения')
                    if args.write:
                        skipped.append({'index': index, 'label': label,
                                        'snapshot': snapshot(k, doc, source, prototypes)})
                    print(f"    Пропущено: {section or category}", flush=True)
                    continue
                if args.write and doc.ReadOnly:
                    raise RuntimeError('Файл открыт только для чтения')
                k.rebuild(doc)
                # После переключения и перестроения запрашиваем текущие объекты заново.
                source = select(manager, index)
                part = k.top(doc)
                values = k.measure(part, args.mode)
                item = {'index': index, 'label': label, 'values': values, 'rec': rec,
                        'marking': marking, 'name': str(part.Name or '')}
                measured.append(item)
                rec['status'] = 'measured'
            except Exception as exc:
                rec.update(status='error', reason=str(exc))
                print(f"    Ошибка: {exc}", flush=True)

        if args.write and any(r['status'] == 'error' for r in records):
            raise RuntimeError('Не все исполнения удалось проверить: файл целиком оставлен без сохранения')
        if args.write and measured:
            backup = Path(backup)
            backup.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(path, backup)
            for item in measured:
                source = select(manager, item['index'])
                write_values(k, doc, k.top(doc), prototypes, item['values'], storage=source)
            # Проверяем все исполнения после всех записей: это обнаружит перезапись
            # соседнего исполнения и распространение значений по наследованию.
            verify(k, doc, manager, measured, skipped, prototypes)
            select(manager, original)
            k.rebuild(doc)
            save(doc)
            saved = True
            close(doc)
            doc = None
            manager = None
            doc = k.open(path, True)
            manager = k.api.IEmbodimentsManager(k.top(doc))
            if int(manager.EmbodimentCount) != count:
                raise RuntimeError('После сохранения изменилось количество исполнений')
            if int(manager.CurrentEmbodimentIndex) != original:
                raise RuntimeError('После сохранения не восстановлено исходное исполнение')
            verify(k, doc, manager, measured, skipped, prototypes)

        rows = []
        for item in measured:
            saved_text = 'сохранены и проверены' if args.write else 'не запрошена'
            item['rec'].update(status='processed', reason=saved_text)
            rows.append([str(path), item['marking'], item['name'], 1,
                         *item['values'], args.mode, saved_text, item['label']])
        return rows, records, bool(measured)
    except Exception as exc:
        reason = str(exc)
        if saved or (args.write and Path(backup).exists()):
            reason += f'; сохранение не подтверждено, резервная копия: {backup}'
        print(f'  Ошибка файла: {reason}', flush=True)
        for item in measured:
            item['rec'].update(status='error', reason=reason)
        records.append({'path': str(path), 'trace': trace, 'status': 'file_error', 'reason': reason})
        return [], records, False
    finally:
        if doc is not None:
            try:
                if manager is not None and original is not None:
                    select(manager, original)
            except Exception as exc:
                records.append({'path': str(path), 'status': 'restore_error', 'reason': str(exc)})
            try:
                close(doc)
            except Exception as exc:
                records.append({'path': str(path), 'status': 'close_error', 'reason': str(exc)})
