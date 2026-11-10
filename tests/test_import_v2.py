"""Static checks and a Python reference model. Does NOT execute VBA or Excel."""
import unittest,re
from pathlib import Path
SOURCE=Path(__file__).resolve().parents[1]/'v2/Inventering_Rapportautomatik.bas'
def plan(existing,incoming,keep=True):
    result=[dict(x) for x in existing];index={}
    for n,row in enumerate(result):
        if row['ID'] in index:raise ValueError('duplicate destination ID')
        index[row['ID']]=n
    seen=set();warnings=[]
    for row in incoming:
        if not row['ID'] or row['ID'] in seen:raise ValueError('invalid source ID')
        seen.add(row['ID'])
    for row in incoming:
        if row['ID'] not in index:index[row['ID']]=len(result);result.append({})
        target=result[index[row['ID']]]
        for field,value in row.items():
            old=target.get(field)
            if keep and field in {'SKICK','KOMMENTAR','ANDELSPARAS','Återbruk i projekt'} and old not in (None,''):
                if value not in (None,'') and old!=value:warnings.append((row['ID'],field))
            else:target[field]=value
    return result,warnings
class ImportTests(unittest.TestCase):
    def test_updates_append_and_preserve_unmatched_and_notes(self):
        old=[{'ID':'001','AREA':1,'SKICK':'Bra','ANDELSPARAS':0,'Återbruk i projekt':'JA','Egen':'behåll'}, {'ID':'002','AREA':7}]
        incoming=[{'AREA':5,'ID':'001','SKICK':'Dåligt','ANDELSPARAS':1,'Återbruk i projekt':'NEJ'}, {'ID':'003','AREA':9}]
        new,warnings=plan(old,incoming)
        self.assertEqual([x['ID'] for x in new],['001','002','003'])
        self.assertEqual(new[0]['AREA'],5);self.assertEqual(new[0]['ANDELSPARAS'],0)
        self.assertEqual(new[0]['Återbruk i projekt'],'JA');self.assertEqual(new[0]['Egen'],'behåll')
        self.assertEqual(new[1],old[1]);self.assertEqual(len(warnings),3)
        self.assertEqual(plan(new,incoming)[0],new)
    def test_duplicates_fail_and_keep_no_updates_assessments(self):
        old=[{'ID':'1','SKICK':'Bra'}]
        with self.assertRaises(ValueError):plan(old,[{'ID':'2'},{'ID':'2'}])
        self.assertEqual(old,[{'ID':'1','SKICK':'Bra'}])
        self.assertEqual(plan(old,[{'ID':'1','SKICK':'Ok'}],False)[0][0]['SKICK'],'Ok')
    def test_1000_rows_and_column_block_runs(self):
        old=[{'ID':str(i),'AREA':i} for i in range(1000)]
        new,_=plan(old,[{'ID':str(i),'AREA':i+1} for i in range(1000)])
        self.assertEqual(len(new),1000);self.assertEqual(sum(r['AREA'] for r in new),500500)
        mask=[True]*1000;runs=[];i=0
        while i<len(mask):
            if mask[i]:
                a=i
                while i<len(mask) and mask[i]:i+=1
                runs.append((a,i))
            else:i+=1
        self.assertEqual(runs,[(0,1000)])
    def test_static_cleanup_and_scope(self):
        s=SOURCE.read_text(encoding='cp1252');a=s.index('Public Sub ImporteraFranVectorworks()');b=s.index('Private Function VImportHeaders',a);t=s[a:b]
        self.assertEqual(t.count('VLastRow(src)'),1)
        self.assertEqual(t.count('dataTable.Resize'),1)
        self.assertLess(t.index('Application.Calculation = xlCalculationManual'),t.index('Workbooks.Open'))
        for statement in ['Application.Calculation = oldCalc','Application.ScreenUpdating = oldScreen','Application.EnableEvents = oldEvents','Application.StatusBar = oldStatus']:
            self.assertIn(statement,t[t.index('CleanUp:'):])
        self.assertNotIn('Application.Calculate\n',t)
        self.assertNotRegex(s,r'(?i)\blocal\s+As\b')
        self.assertIn('target.Value2 = values',t)
if __name__=='__main__':unittest.main(verbosity=2)
