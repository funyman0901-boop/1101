import copy
import json
from pathlib import Path
import tempfile
from types import SimpleNamespace as NS
import unittest

import kompas_dimensions as core
import embodiment_processing as ep


def prop(index):
    return NS(Id=101+index, Name=core.NAMES[index], DataType=2,
              ReadOnly=False, IsVariable=False)


class Storage:
    def __init__(self, doc, index):
        self.doc, self.index = doc, index
        self.Name = f'Исполнение {index}'
    def GetPropertyValue(self, p, *args):
        value = self.doc.data[self.index].get(p.Id)
        if value is None and self.doc.owner.inherit and self.index > 0:
            value = self.doc.data[0].get(p.Id)
        return (value is not None, value, False)
    def SetPropertyValue(self, p, value, base):
        targets = range(self.doc.count) if self.doc.owner.overwrite_all else [self.index]
        for index in targets:
            self.doc.data[index][p.Id] = value
        self.doc.Changed = True
        return True


class Document:
    def __init__(self, owner, path, readonly=False):
        self.owner, self.path = owner, str(path)
        self.ReadOnly = readonly
        self.Changed = False
        self.count = owner.count
        state = owner.saved_state
        self.active = state['active'] if state else owner.original
        self.data = copy.deepcopy(state['data']) if state else [{} for i in range(self.count)]
        self.defs = copy.deepcopy(state['defs']) if state else {}
        if Path(path).name == 'dimensions_seed.m3d':
            self.defs = {p.Id:p for p in owner.prototypes.values()}
    def Save(self):
        if self.owner.save_failure:
            raise RuntimeError('Save failed')
        self.owner.saves += 1
        self.owner.saved_state = copy.deepcopy({'active':self.active,'data':self.data,'defs':self.defs})
        self.Changed = False
    def Close(self, option):
        self.owner.closed.append((self.path, self.active))
        return True


class Manager:
    def __init__(self, doc):
        self.doc = doc
    @property
    def EmbodimentCount(self): return self.doc.count
    @property
    def CurrentEmbodimentIndex(self): return self.doc.active
    @property
    def CurrentEmbodiment(self): return Storage(self.doc, self.doc.active)
    def SetCurrentEmbodiment(self, index):
        if index == self.doc.owner.switch_failure:
            return False
        self.doc.active = index
        return True
    def GetCurrentEmbodimentMarking(self, *args): return f'ABC.001-{self.doc.active:02d}'


class Fake:
    def __init__(self, path, count=3, original=1):
        self.path, self.count, self.original = path, count, original
        self.prototypes = {name: prop(i) for i,name in enumerate(core.NAMES)}
        self.saved_state = None
        self.saves = 0
        self.closed, self.measured = [], []
        self.category_map = {}
        self.measure_failure = self.switch_failure = self.rebuild_failure = None
        self.overwrite_all = self.inherit = self.save_failure = self.corrupt_readback = False
        self.app = NS()
        self.api = NS(IEmbodimentsManager=lambda p: Manager(p.doc),
                      IPropertyKeeper=lambda p:p, IProperty=lambda p:p)
        self.mng = NS(AddProperty=self.add)
    def add(self, doc, p):
        doc.defs[p.Id] = copy.copy(p)
        return doc.defs[p.Id]
    def opened(self): return {}
    def open(self, path, readonly):
        doc = Document(self, path, readonly)
        if self.corrupt_readback and self.saved_state and Path(path).name != 'dimensions_seed.m3d':
            doc.data[0][101] = -1
        return doc
    def top(self, doc):
        def forbidden_update():
            raise AssertionError('IPart7.Update must not be called')
        return NS(doc=doc, Name=f'Брус {doc.active}', Marking=f'ABC.001-{doc.active:02d}',
                  Standard=False, Update=forbidden_update)
    def category(self, doc, part, source=None):
        value = self.category_map.get(doc.active, 'normal')
        return value, value
    def rebuild(self, doc):
        if doc.active == self.rebuild_failure:
            raise RuntimeError('Rebuild failed')
    def measure(self, part, mode):
        index = part.doc.active
        self.measured.append(index)
        if index == self.measure_failure:
            raise RuntimeError('Measure failed')
        return (1000+index*100,100,50)
    def props(self, doc): return list(doc.defs.values())
    def prop(self, doc, ident): return doc.defs.get(ident)
    def read(self, storage, p): return core.Kompas.read(self, storage, p)


class EmbodimentTests(unittest.TestCase):
    def setup_case(self, tmp, write=True):
        root = Path(tmp)
        source = root/'part.m3d'
        source.write_bytes(b'original-model')
        k = Fake(source)
        args = NS(write=write,mode='bbox')
        backup = root/'backups'/'part.m3d'
        return k, args, source, backup

    def process(self, k, args, source, backup):
        return ep.process(k, k.open(source, not args.write), source, args,
                          k.prototypes if args.write else {}, backup, 'part.m3d')

    def test_distinct_values_one_save_and_original_restored(self):
        with tempfile.TemporaryDirectory() as tmp:
            k,args,source,backup = self.setup_case(tmp)
            rows,log,ok = self.process(k,args,source,backup)
            self.assertTrue(ok)
            self.assertEqual(k.saves,1)
            self.assertEqual(k.saved_state['active'],1)
            self.assertEqual([d[101] for d in k.saved_state['data']],[1000,1100,1200])
            self.assertEqual([r[4] for r in rows],[1000,1100,1200])
            self.assertEqual([r[-1] for r in rows],['ABC.001-00','ABC.001-01','ABC.001-02'])
            self.assertEqual(backup.read_bytes(),b'original-model')
            self.assertTrue(all(r['status']=='processed' for r in log))
            self.assertTrue(all(index==1 for path,index in k.closed))

    def test_excel_only_never_saves(self):
        with tempfile.TemporaryDirectory() as tmp:
            k,args,source,backup = self.setup_case(tmp,False)
            rows,log,ok = self.process(k,args,source,backup)
            self.assertEqual(len(rows),3)
            self.assertEqual(k.saves,0)
            self.assertFalse(backup.exists())
            self.assertEqual(k.closed[-1][1],1)

    def test_excluded_execution_untouched(self):
        with tempfile.TemporaryDirectory() as tmp:
            k,args,source,backup = self.setup_case(tmp)
            k.category_map[1]='other'
            rows,log,ok = self.process(k,args,source,backup)
            self.assertTrue(ok)
            self.assertEqual(k.measured,[0,2])
            self.assertEqual(k.saved_state['data'][1],{})
            self.assertEqual(sum(r['status']=='other' for r in log),1)

    def test_all_excluded_no_save_or_backup(self):
        with tempfile.TemporaryDirectory() as tmp:
            k,args,source,backup = self.setup_case(tmp)
            k.category_map={0:'standard',1:'other',2:'standard'}
            rows,log,ok = self.process(k,args,source,backup)
            self.assertFalse(ok)
            self.assertEqual(rows,[])
            self.assertEqual(k.saves,0)
            self.assertFalse(backup.exists())

    def test_measure_failure_aborts_entire_file(self):
        with tempfile.TemporaryDirectory() as tmp:
            k,args,source,backup = self.setup_case(tmp)
            k.measure_failure=1
            rows,log,ok = self.process(k,args,source,backup)
            self.assertFalse(ok)
            self.assertEqual(k.saves,0)
            self.assertEqual(rows,[])
            self.assertEqual(k.measured,[0,1,2])
            self.assertEqual(k.closed[-1][1],1)

    def test_switch_failure_not_saved(self):
        with tempfile.TemporaryDirectory() as tmp:
            k,args,source,backup = self.setup_case(tmp)
            k.original=0
            k.switch_failure=1
            rows,log,ok = self.process(k,args,source,backup)
            self.assertFalse(ok)
            self.assertEqual(k.saves,0)

    def test_cross_execution_overwrite_detected_before_save(self):
        with tempfile.TemporaryDirectory() as tmp:
            k,args,source,backup = self.setup_case(tmp)
            k.overwrite_all=True
            rows,log,ok = self.process(k,args,source,backup)
            self.assertFalse(ok)
            self.assertEqual(k.saves,0)
            self.assertTrue(backup.exists())

    def test_inheritance_to_excluded_execution_aborts_save(self):
        with tempfile.TemporaryDirectory() as tmp:
            k,args,source,backup = self.setup_case(tmp)
            k.category_map[2]='other'
            k.inherit=True
            rows,log,ok = self.process(k,args,source,backup)
            self.assertFalse(ok)
            self.assertEqual(k.saves,0)
            self.assertTrue(any('наследованию' in r.get('reason','') for r in log))

    def test_saved_readback_failure_reports_backup(self):
        with tempfile.TemporaryDirectory() as tmp:
            k,args,source,backup = self.setup_case(tmp)
            k.corrupt_readback=True
            rows,log,ok = self.process(k,args,source,backup)
            self.assertFalse(ok)
            self.assertEqual(k.saves,1)
            self.assertTrue(any('резервная копия' in r.get('reason','') for r in log))

    def test_rebuild_failure_not_saved(self):
        with tempfile.TemporaryDirectory() as tmp:
            k,args,source,backup = self.setup_case(tmp)
            k.rebuild_failure=0
            self.assertFalse(self.process(k,args,source,backup)[2])
            self.assertEqual(k.saves,0)

    def test_save_failure_reported(self):
        with tempfile.TemporaryDirectory() as tmp:
            k,args,source,backup = self.setup_case(tmp)
            k.save_failure=True
            self.assertFalse(self.process(k,args,source,backup)[2])
            self.assertEqual(k.saves,0)
            self.assertTrue(backup.exists())

    def test_folder_integration_writes_rows_per_execution(self):
        with tempfile.TemporaryDirectory() as tmp:
            root=Path(tmp)
            models=root/'models'
            models.mkdir()
            k,args,source,backup=self.setup_case(models)
            schema=root/'schema'
            schema.mkdir()
            (schema/'properties.json').write_text(json.dumps({n:p.Id for n,p in k.prototypes.items()}),encoding='utf-8')
            args.folder=str(models)
            args.schema=str(schema)
            args.output=str(root/'reports')
            self.assertEqual(core.run(k,args),0)
            journal=json.loads(next((root/'reports').glob('*/journal.json')).read_text(encoding='utf-8'))
            self.assertEqual(journal['summary']['Обработано уникальных файлов'],1)
            self.assertEqual(journal['summary']['Обработано исполнений / отдельных деталей'],3)

    def test_single_part_write_never_calls_update(self):
        with tempfile.TemporaryDirectory() as tmp:
            k,args,source,backup=self.setup_case(tmp)
            doc=k.open(source,False)
            storage=Storage(doc,0)
            core.write_values(k,doc,k.top(doc),k.prototypes,(1000,100,50),storage=storage)
            self.assertEqual(doc.data[0][101],1000)


if __name__=='__main__':
    unittest.main()
