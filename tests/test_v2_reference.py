"""Four core reference checks. Does not claim to execute VBA/Excel.
Usage: python tests/test_v2_reference.py ORIGINAL.xlsm TEST.xlsm
"""
import sys,json,zipfile,hashlib,unittest,collections,copy,re
from pathlib import Path
import openpyxl
ORIGINAL,TEST=sys.argv[1:3];sys.argv=sys.argv[:1]
w=openpyxl.load_workbook(TEST,data_only=False,keep_vba=True)
old=openpyxl.load_workbook(ORIGINAL,data_only=False,keep_vba=True)
def norm(v):return ' '.join(str(v if v is not None else '').replace('\xa0',' ').split()).upper()
def table(name):
 t=w['Regler'].tables[name];a,b,c,d=openpyxl.utils.range_boundaries(t.ref)
 hs=[norm(w['Regler'].cell(b,j).value) for j in range(a,c+1)]
 return [dict(zip(hs,[w['Regler'].cell(i,j).value for j in range(a,c+1)])) for i in range(b+1,d+1)]
def get(rule,key):return rule.get(norm(key),'')
def match(rule,r,n):
 f=get(rule,'Fält '+n)
 if not f:return True
 a,b=norm(r[norm(f)]),norm(get(rule,'Värde '+n));m=norm(get(rule,'Metod '+n))
 if m=='INNEHÅLLER':return b in a
 if m=='ÄR EXAKT':return a==b
 if m=='BÖRJAR MED':return a.startswith(b)
 if m=='INNEHÅLLER INTE':return b not in a
 raise ValueError(m)
def matches(rule,r):
 if norm(get(rule,'Aktiv'))!='JA':return False
 if get(rule,'Lokalt delområde') and norm(get(rule,'Lokalt delområde'))!=norm(r['DELOMRÅDE']):return False
 if get(rule,'Fält X') and match(rule,r,'X'):return False
 a,b,c=[match(rule,r,str(i)) for i in (1,2,3)]
 return {'1':a,'':a,'1 OCH 2':a and b,'1 ELLER 2':a or b,'1 OCH 2 OCH 3':a and b and c,'1 ELLER 2 ELLER 3':a or b or c,'1 OCH (2 ELLER 3)':a and (b or c),'(1 OCH 2) ELLER 3':(a and b) or c,'1 ELLER (2 OCH 3)':a or (b and c),'(1 ELLER 2) OCH 3':(a or b) and c}[norm(get(rule,'Logik'))]
def pick(name,r):
 candidates=[x for x in table(name) if matches(x,r)]
 return max(candidates,key=lambda x:(bool(get(x,'Lokalt delområde')),float(get(x,'Prioritet') or 0))) if candidates else None
def rows():
 hs=[norm(c.value) for c in w['Rådata'][1]]
 return [dict(zip(hs,r)) for r in w['Rådata'].iter_rows(min_row=3,values_only=True) if any(v is not None for v in r)]
def unit(row):
 e=pick('tblEnheterV2',row);c=pick('tblKategorierV2',row);v=get(e or c,'Enhet')
 return v if v!='Auto' else 'lpm' if row['LÄNGD'] else 'm²' if row['AREA'] else 'antal'
class Checks(unittest.TestCase):
 def test_A_CS1_decisions(self):
  source=[r for r in rows() if r['LITTRERING']=='CS1'];self.assertEqual(len(source),55)
  self.assertTrue(all(not r['ÅTERBRUK I PROJEKT'] for r in source))
  synthetic=copy.deepcopy(source)
  for i,r in enumerate(synthetic):r['ÅTERBRUK I PROJEKT']='JA' if i<53 else 'NEJ'
  totals=collections.Counter()
  for r in synthetic:totals[(r['LITTRERING'],r['ÅTERBRUK I PROJEKT'],unit(r))]+=r['ANTAL']
  self.assertEqual(dict(totals),{('CS1','JA','antal'):53,('CS1','NEJ','antal'):2})
 def test_B_KS1(self):
  selected=[r for r in rows() if r['LITTRERING']=='KS1'];self.assertEqual(len(selected),6)
  for r in selected:
   rule=pick('tblKategorierV2',r)
   self.assertEqual((get(rule,'Rubriknivå 1'),get(rule,'Rubriknivå 2')),('BETONGMATERIAL','KANTSTÖD'))
   self.assertEqual(unit(r),'lpm')
  amount=sum(r['LÄNGD'] for r in selected);self.assertAlmostEqual(amount,161.09071791699728)
  self.assertEqual(round(amount/.5)*.5,161)
 def test_C_stone_and_SMG2(self):
  names=['SMÅGATSTEN','SMÅGATSTEN I BETONG','EN RAD SMÅGATSTEN 2 CM FOG']
  for name in names:
   selected=[r for r in rows() if r['MATERIALTYP']==name];self.assertTrue(selected)
   for r in selected:
    rule=pick('tblKategorierV2',r)
    self.assertEqual(tuple(get(rule,f'Rubriknivå {i}') for i in (1,2,3)),('NATURSTENSMATERIAL','GATSTEN','SMÅGATSTEN'))
  selected=[r for r in rows() if r['LITTRERING']=='SMG2'];self.assertEqual(len(selected),22)
  self.assertTrue(all(unit(r)=='lpm' for r in selected))
  self.assertAlmostEqual(sum(r['LÄNGD'] for r in selected),106.11854324583435)
 def test_D_package_preservation_and_complete_source_accounting(self):
  # This verifies package/data preservation, NOT a full Excel report run.
  for row in old['Rådata']:
   for cell in row:self.assertEqual(w['Rådata'][cell.coordinate].value,cell.value)
  with zipfile.ZipFile(ORIGINAL) as a,zipfile.ZipFile(TEST) as b:
   for n in a.namelist():
    if any(x in n for x in ['drawings/','media/','activeX/','ctrlProps/']):self.assertEqual(a.read(n),b.read(n))
  included=[];filtered=[]
  for r in rows():
   if 'X' in norm(r['LITTRERING']) or norm(r['MATERIALTYP'])=='TEXT':filtered.append(r)
   else:included.append(r)
  self.assertEqual(len(included)+len(filtered),len(rows()))
  ids=[norm(r['INFO-ID-REFERENS']) for r in included]
  self.assertTrue(all(ids));self.assertEqual(len(ids),len(set(ids)))
  self.assertTrue(all(pick('tblKategorierV2',r) for r in included))
  print('\nReference accounting:',len(included),'included,',len(filtered),'filtered, 1 verified export-total excluded; Excel update NOT RUN.')
if __name__=='__main__':unittest.main(verbosity=2)
