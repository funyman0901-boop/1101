import json
from pathlib import Path
import tempfile
from types import SimpleNamespace as NS
import unittest
from unittest.mock import patch
from openpyxl import load_workbook

import bill_of_materials as bom
from test_embodiments import Fake, Manager


def component(path, material='', name='Деталь', category='normal', children=(),
              embodiment='ABC.001', excluded=False):
    def parts(kind):
        assert kind == 0, 'Must include all occurrences and patterned copies'
        return children
    return NS(FileName=str(path), Material=material, Name=name, Marking=embodiment,
              category=category, Detail=not children, IsLocal=False, PartsEx=parts,
              Excluded=excluded, LoadState=0, embodiment=embodiment)


def tree_api():
    return NS(IPart7=lambda p:p, IFeature7=lambda p:p,
              IEmbodimentsManager=lambda p:NS(GetCurrentEmbodimentMarking=lambda *a:p.embodiment))


class BillTests(unittest.TestCase):
    def test_linear_sums_all_lengths_times_quantity(self):
        rows = bom.aggregate([
            {'material':'Труба 50х50х3 Ст3','dimensions':(1000,50,50),'quantity':4},
            {'material':'труба 50 × 50 × 3 Ст3','dimensions':(500,50,50),'quantity':2}])
        self.assertEqual(len(rows),1)
        self.assertEqual(rows[0]['length'],5000)
        self.assertIsNone(rows[0]['quantity'])

    def test_different_grades_and_sections_do_not_merge(self):
        rows=bom.aggregate([{'material':m,'dimensions':(1000,50,50)} for m in
                           ('Труба 50х50х3 Ст3','Труба 50х50х3 09Г2С','Труба 60х60х3 Ст3')])
        self.assertEqual(len(rows),3)

    def test_sheet_rows_by_material_thickness_and_size(self):
        rows=bom.aggregate([
            {'material':'Лист Ст3','dimensions':(500,300,4),'quantity':2},
            {'material':'Лист Ст3','dimensions':(500,300,4),'quantity':3},
            {'material':'Лист Ст3','dimensions':(500,300,5),'quantity':1},
            {'material':'Лист Ст3','dimensions':(600,300,4),'quantity':1}])
        self.assertEqual(len(rows),3)
        first=next(r for r in rows if '500 × 300' in r['label'] and 't=4 ' in r['label'])
        self.assertEqual(first['quantity'],5)
        self.assertTrue(all(r['length'] is None for r in rows))

    def test_piece_sections_and_unclassified_last(self):
        rows=bom.aggregate([
            {'name':'Без материала'},
            {'name':'Болт М8','category':'standard','quantity':4},
            {'name':'Петля','category':'other','quantity':2},
            {'material':'Труба 50х50х3','dimensions':(1000,50,50)},
            {'name':'Стальная деталь','material':'Сталь 3'}])
        self.assertEqual([r['section'] for r in rows[:3]],list(bom.SECTIONS[:3]))
        self.assertTrue(all(r['section']==bom.SECTIONS[3] for r in rows[3:]))
        self.assertEqual(rows[1]['quantity'],4)
        self.assertIsNone(rows[1]['length'])

    def test_piece_different_marking_not_merged(self):
        rows=bom.aggregate([{'name':'Петля','category':'other','marking':m} for m in ('А1','А2')])
        self.assertEqual(len(rows),2)

    def test_unknown_material_is_not_assumed_linear(self):
        self.assertEqual(bom.material_type('Труба листовая сварная 50х50х3'),'linear')
        self.assertEqual(bom.material_type('Лист профилированный 4 мм'),'sheet')
        rows=bom.aggregate([{'material':'Сталь 20','dimensions':(1000,50,50)}])
        self.assertEqual(rows[0]['section'],bom.SECTIONS[3])
        self.assertIsNone(rows[0]['length'])

    def test_nested_repeated_subassemblies_and_copies(self):
        tubes=[component('/pipe.m3d','Труба 50х50х3') for i in range(4)]
        sub1=component('/sub.a3d',children=tubes,embodiment='SUB-01')
        sub2=component('/sub.a3d',children=tubes,embodiment='SUB-01')
        root=component('/root.a3d',children=(sub1,sub2),embodiment='ROOT')
        k=NS(api=tree_api(),top=lambda d:root,category=lambda d,p:(p.category,''))
        journal=[]
        items=bom.collect(k,NS(PathName='/root.a3d'),journal)
        self.assertEqual(len(items),8)
        self.assertEqual(journal,[])

    def test_purchased_subassembly_is_one_piece(self):
        bought=component('/hinge.a3d',name='Петля',category='other',
                         children=(component('/screw.m3d',category='standard'),))
        root=component('/root.a3d',children=(bought,))
        k=NS(api=tree_api(),top=lambda d:root,category=lambda d,p:(p.category,''))
        items=bom.collect(k,NS(PathName='/root.a3d'),[])
        self.assertEqual(len(items),1)
        self.assertEqual(items[0]['name'],'Петля')

    def test_excluded_from_assembly_is_not_counted(self):
        root=component('/root.a3d',children=(component('/skip.m3d',excluded=True),
                                          component('/good.m3d')))
        k=NS(api=tree_api(),top=lambda d:root,category=lambda d,p:(p.category,''))
        journal=[]
        self.assertEqual(len(bom.collect(k,NS(PathName='/root.a3d'),journal)),1)
        self.assertEqual(journal[0]['status'],'excluded')

    def test_correct_source_embodiment_only(self):
        with tempfile.TemporaryDirectory() as tmp:
            path=Path(tmp)/'part.m3d'
            path.write_bytes(b'model')
            k=Fake(path)
            class SourceManager(Manager):
                def GetEmbodimentMarking(self,index,*args): return f'ABC.001-{index:02d}'
            k.api.IEmbodimentsManager=lambda p:SourceManager(p.doc)
            dims=bom.source_dimensions(k,{'path':str(path),'embodiment':'ABC.001-02'},'bbox')
            self.assertEqual(dims,(1200,100,50))
            self.assertEqual(k.measured,[2])
            self.assertEqual(k.closed[-1][1],1)
            self.assertEqual(k.saves,0)

    def test_ambiguous_execution_does_not_measure(self):
        with tempfile.TemporaryDirectory() as tmp:
            path=Path(tmp)/'part.m3d'
            path.write_bytes(b'model')
            k=Fake(path)
            class BadManager(Manager):
                def GetEmbodimentMarking(self,index,*args): return 'SAME'
            k.api.IEmbodimentsManager=lambda p:BadManager(p.doc)
            with self.assertRaises(RuntimeError):
                bom.source_dimensions(k,{'path':str(path),'embodiment':'SAME'},'bbox')
            self.assertEqual(k.measured,[])

    def test_single_sheet_literal_cells_and_incomplete_banner(self):
        with tempfile.TemporaryDirectory() as tmp:
            rows=bom.aggregate([{'name':'=CMD','category':'standard','quantity':4}])
            path,title=bom.export(Path(tmp),rows,[{'status':'error'}],'test.a3d','TEST-01')
            wb=load_workbook(path,data_only=False)
            self.assertEqual(wb.sheetnames,['Ведомость'])
            self.assertIn('НЕПОЛНЫЙ ОТЧЁТ',wb.active['A1'].value)
            self.assertFalse(any(c.data_type=='f' for s in wb for row in s for c in row))

    def test_complete_run_quantity_cache_and_no_model_save(self):
        with tempfile.TemporaryDirectory() as tmp:
            rootpath=Path(tmp)/'root.a3d'
            rootpath.write_bytes(b'assembly')
            pipe=component(Path(tmp)/'pipe.m3d','Труба 50х50х3')
            bolt=component(Path(tmp)/'bolt.m3d',name='Болт М8',category='standard')
            root=component(rootpath,children=(pipe,pipe,pipe,pipe,bolt,bolt),embodiment='ROOT-01')
            doc=NS(PathName=str(rootpath),Changed=False)
            k=NS(api=tree_api(),top=lambda d:root,category=lambda d,p:(p.category,''),
                 opened=lambda:{bom.key(rootpath):doc})
            with patch.object(bom,'source_dimensions',return_value=(1000,50,50)) as measure:
                result=bom.run(k,rootpath,Path(tmp)/'reports')
                self.assertEqual(measure.call_count,1)
            self.assertEqual(result['errors'],0)
            data=json.loads(Path(result['workbook']).with_suffix('.json').read_text(encoding='utf-8'))
            self.assertEqual(next(r for r in data['rows'] if r['section']=='Сортамент')['length'],4000)
            self.assertEqual(next(r for r in data['rows'] if r['section']=='Метизы')['quantity'],2)


if __name__=='__main__':
    unittest.main()
