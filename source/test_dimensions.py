import tempfile
import unittest
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
