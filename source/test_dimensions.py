import tempfile
import unittest
from unittest.mock import patch
from pathlib import Path
from types import SimpleNamespace as NS
from openpyxl import load_workbook
import kompas_dimensions as m


class Parts:
    def __init__(self, *items):
        self.items = items
        self.Count = len(items)
    def Part(self, i):
        return self.items[i]


def part(path, cat="normal", children=()):
    return NS(FileName=path, cat=cat, Detail=not children,
              IsLocal=False, Parts=Parts(*children))


class Tests(unittest.TestCase):
    def category_fixture(self, result=(False, None, False), complex_value=None, storage=None, standard=False, has_property=True):
        keeper = NS(IsComplexPropertyValue=lambda p: complex_value is not None,
                    GetPropertyValue=lambda *a: result,
                    GetComplexPropertyValue=lambda *a: (complex_value, False), Properties=storage)
        k = object.__new__(m.Kompas)
        k.api = NS(IPropertyKeeper=lambda p: keeper)
        k.prop = lambda *a: NS(Name='Раздел спецификации') if has_property else None
        return k, NS(Standard=standard)

    def test_unset_category_response_from_user_log(self):
        k, part = self.category_fixture()
        self.assertEqual(k.category(None, part), ('normal', 'Не задан'))

    def test_missing_category_definition(self):
        k, part = self.category_fixture(has_property=False)
        self.assertEqual(k.category(None, part)[0], 'normal')

    def test_empty_category_success(self):
        for value in (None, ''):
            k, part = self.category_fixture(result=(True, value, False))
            self.assertEqual(k.category(None, part)[0], 'normal')

    def test_explicit_excluded_categories_preserved(self):
        for name, expected in (('Прочие изделия','other'),('Стандартные изделия','standard')):
            xml = f'<document><property id="sectionName" value="{name}"/></document>'
            k, part = self.category_fixture(complex_value=xml)
            self.assertEqual(k.category(None, part)[0], expected)
            k, part = self.category_fixture(result=(True, name, False))
            self.assertEqual(k.category(None, part)[0], expected)
            storage = f'<infObject><property id="SPCSection"><property id="sectionName" value="{name}"/></property></infObject>'
            k, part = self.category_fixture(storage=storage)
            self.assertEqual(k.category(None, part)[0], expected)
        k, part = self.category_fixture(standard=True)
        self.assertEqual(k.category(None, part)[0], 'standard')

    def test_unreadable_category_does_not_default(self):
        for result in ((False, 'unknown', False), (False, None, True), (False, None)):
            k, part = self.category_fixture(result=result)
            with self.assertRaises(RuntimeError):
                k.category(None, part)

    def test_numeric_property_reads_stay_strict(self):
        k, part = self.category_fixture()
        with self.assertRaises(RuntimeError):
            k.read(part, NS(Name='Длина'))

    def test_recursive_folder_and_exclusions(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            (root / 'sub' / 'deep').mkdir(parents=True)
            (root / 'reports' / 'backups').mkdir(parents=True)
            (root / 'schema').mkdir()
            for name in ('one.m3d', 'sub/two.M3D', 'sub/deep/three.m3d',
                         'sub/ignored.a3d', 'reports/backups/copy.m3d', 'schema/seed.m3d'):
                (root / name).write_bytes(b'model')
            log = []
            groups = m.folder_groups(root, (root/'reports', root/'schema'), log)
            self.assertEqual(len(groups), 3)
            self.assertEqual({Path(g['path']).name for g in groups.values()},
                             {'one.m3d', 'two.M3D', 'three.m3d'})
            self.assertTrue(all(not g['blocked'] for g in groups.values()))

    def test_missing_folder_rejected(self):
        with tempfile.TemporaryDirectory() as tmp:
            with self.assertRaises(ValueError):
                m.folder_groups(Path(tmp)/'missing', [], [])
            with self.assertRaises(ValueError):
                m.folder_groups(tmp, [tmp], [])

    def test_folder_does_not_follow_symlink(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            models = root/'models'
            models.mkdir()
            (root/'outside').mkdir()
            (root/'outside'/'do_not_open.m3d').write_bytes(b'x')
            try:
                (models/'link').symlink_to(root/'outside', target_is_directory=True)
            except OSError:
                self.skipTest('Symlinks unavailable on this system')
            self.assertEqual(m.folder_groups(models, [], []), {})

    def test_folder_run_without_assembly_continues_after_error(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            models = root/'models'
            models.mkdir()
            for name in ('good.m3d','broken.m3d','other.m3d','standard.m3d'):
                (models/name).write_bytes(b'model')
            opened, measured, closed = [], [], []
            class Doc:
                Changed = False
                ReadOnly = False
                def __init__(self, path):
                    self.path = Path(path)
                    self.part = NS(Marking='ABC.001', Name='Брус')
                def Close(self, option):
                    closed.append(self.path.name)
                    return True
            class Fake:
                # app deliberately has no ActiveDocument
                app = NS()
                api = NS(IEmbodimentsManager=lambda p: NS(EmbodimentCount=1))
                def opened(self): return {}
                def open(self, path, readonly):
                    opened.append(Path(path).name)
                    return Doc(path)
                def top(self, doc):
                    doc.part.filename = doc.path.name
                    return doc.part
                def category(self, doc, part):
                    cat = {'other.m3d':'other', 'standard.m3d':'standard'}.get(doc.path.name, 'normal')
                    return cat, cat
                def measure(self, part, mode):
                    measured.append(part.filename)
                    if part.filename == 'broken.m3d':
                        raise RuntimeError('Cannot measure')
                    return (2500,100,100)
            args = NS(folder=str(models), output=str(root/'reports'), schema=str(root/'schema'), write=False, mode='bbox')
            self.assertEqual(m.run(Fake(), args), 2)
            report = next((root/'reports').glob('*/journal.json'))
            import json
            journal = json.loads(report.read_text(encoding='utf-8'))
            self.assertEqual(journal['summary']['Обработано уникальных файлов'], 1)
            self.assertEqual(journal['summary']['Пропущено прочих изделий'], 1)
            self.assertEqual(journal['summary']['Пропущено стандартных изделий'], 1)
            self.assertEqual(set(measured), {'good.m3d','broken.m3d'})
            self.assertEqual(len(closed), 4)

    def test_empty_folder_still_exports_journal(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            models = root/'models'
            models.mkdir()
            args = NS(folder=str(models), output=str(root/'reports'), schema=str(root/'schema'), write=False, mode='bbox')
            self.assertEqual(m.run(NS(app=NS()), args), 2)
            self.assertEqual(len(list((root/'reports').glob('*/dimensions.xlsx'))), 1)

    def test_folder_write_backup_and_saved_readback(self):
        with tempfile.TemporaryDirectory() as tmp:
            import json
            root = Path(tmp)
            models, schema = root/'models', root/'schema'
            schema.mkdir()
            (schema/'properties.json').write_text(json.dumps(dict(zip(m.NAMES, (101,102,103)))), encoding='utf-8')
            for sub in ('a','b'):
                (models/sub).mkdir(parents=True)
                (models/sub/'same.m3d').write_bytes(sub.encode())
            saved, reopened = [], []
            values_by_path = {}
            class Doc:
                Changed = False
                ReadOnly = False
                def __init__(self, path):
                    self.path = Path(path)
                    self.part = NS(Marking='ABC.001', Name='Брус', path=str(self.path))
                def Close(self, option): return True
                def Save(self):
                    saved.append(str(self.path))
                    values_by_path[str(self.path)] = (2500,100,100)
            class Fake:
                app = NS()
                api = NS(IEmbodimentsManager=lambda p: NS(EmbodimentCount=1))
                def opened(self): return {}
                def open(self, path, readonly):
                    reopened.append((str(path), readonly))
                    return Doc(path)
                def top(self, doc): return doc.part
                def category(self, doc, part): return 'normal', ''
                def measure(self, part, mode): return (2500,100,100)
                def prop(self, doc, ident):
                    return NS(Id=ident, Name=m.NAMES[int(ident)-101], DataType=2, IsVariable=False)
                def read(self, part, prop): return values_by_path[part.path][int(prop.Id)-101]
            def write(k, doc, part, prototypes, values):
                if doc.path.parent.name == 'a':
                    raise RuntimeError('Simulated write failure')
            args = NS(folder=str(models), output=str(root/'reports'), schema=str(schema), write=True, mode='bbox')
            with patch.object(m, 'write_values', write):
                self.assertEqual(m.run(Fake(), args), 2)
            self.assertEqual(saved, [str(models/'b'/'same.m3d')])
            backups = sorted((root/'reports').glob('*/backups/*/same.m3d'))
            self.assertEqual([p.read_bytes() for p in backups], [b'a', b'b'])
            self.assertIn((str(models/'b'/'same.m3d'), True), reopened)

    def test_axis_independent(self):
        for dims in ((100, 100, 2500), (100, 2500, 100), (2500, 100, 100)):
            self.assertEqual(m.dimensions(dims), (2500, 100, 100))

    def test_bad_geometry(self):
        for dims in ((0, 1, 2), (1, -2, 3), (float("nan"), 1, 2),
                     (float("inf"), 2, 3), (0.000001, 1, 2)):
            with self.assertRaises(ValueError):
                m.dimensions(dims)

    def test_sections(self):
        for name in ("Прочие изделия", "Стандартные изделия", "Детали"):
            xml = f'<document><property id="sectionName" value="{name}"/></document>'
            self.assertEqual(m.parse_section(xml), name)
        self.assertEqual(m.parse_section('<document/>'), "")
        with self.assertRaises(ValueError):
            m.parse_section('<document><unknown/></document>')

    def test_nested_duplicates_and_conflict(self):
        a = part('/parts/a.m3d')
        b = part('/parts/b.m3d', "other")
        c = part('/parts/c.m3d', "standard")
        conflict = part('/parts/a.m3d', "other")
        assembly = part('/parts/sub.a3d', children=(a, b, c))
        top = part('/parts/root.a3d', children=(assembly, a, conflict))
        k = NS(top=lambda doc: top, category=lambda doc, p: (p.cat, p.cat))
        log = []
        records = m.collect(k, NS(PathName='/parts/root.a3d'), log)
        self.assertEqual(len(records[m.key('/parts/a.m3d')]['items']), 3)
        self.assertTrue(records[m.key('/parts/a.m3d')]['blocked'])
        self.assertEqual(sum(r['status'] == 'other' for r in log), 2)
        self.assertEqual(sum(r['status'] == 'standard' for r in log), 1)

    def test_excluded_assembly_propagates(self):
        top = part('/root.a3d', children=(part('/sub.a3d', 'other',
                   (part('/nested.m3d'),)),))
        log = []
        k = NS(top=lambda d: top, category=lambda d, p: (p.cat, p.cat))
        records = m.collect(k, NS(PathName='/root.a3d'), log)
        self.assertTrue(records[m.key('/nested.m3d')]['blocked'])
        self.assertTrue(all(r['status'] == 'other' for r in log))

    def test_bbox_com_response(self):
        k = object.__new__(m.Kompas)
        p = NS(GetGabarit=lambda *args: (True, -10, 5, 20, 2490, 105, 120))
        self.assertEqual(k.measure(p, 'bbox'), (2500, 100, 100))
        with self.assertRaises(RuntimeError):
            k.measure(NS(GetGabarit=lambda *a: (False, 0, 0, 0, 0, 0, 0)), 'bbox')

    def test_excel_literals_and_empty_report(self):
        with tempfile.TemporaryDirectory() as tmp:
            folder = Path(tmp)
            m.export(folder, [['=file', '=mark', '=1+2', 2, 2500., 100., 100., 'bbox', 'нет']],
                     [{'path': '=file', 'status': 'error', 'reason': '=cmd'}], {'Обработано': 1})
            wb = load_workbook(folder / 'dimensions.xlsx', data_only=False)
            self.assertEqual(wb.active['C2'].value, '=1+2')
            self.assertEqual(wb.active['E2'].value, 2500)
            self.assertFalse(any(c.data_type == 'f' for s in wb for r in s for c in r))
            m.export(folder, [], [], {'Обработано': 0})


if __name__ == '__main__':
    unittest.main()
