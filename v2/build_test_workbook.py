"""Build a test copy by editing specific OOXML parts; retain every unrelated part.
Requires openpyxl (read-only analysis), oletools and olefile. No Excel automation.
"""
from pathlib import Path
import sys,struct,io,zipfile,hashlib,copy,re,json
from lxml import etree as E
import olefile
from oletools.olevba import VBA_Project,decompress_stream
NS='http://schemas.openxmlformats.org/spreadsheetml/2006/main'; R='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PK='http://schemas.openxmlformats.org/package/2006/relationships';CT='http://schemas.openxmlformats.org/package/2006/content-types'
E.register_namespace('r',R)
def q(n):return '{'+NS+'}'+n
def compressed(data):
    chunks=[]
    for i in range(0,len(data),4096):
        chunk=data[i:i+4096]
        if len(chunk)<=3640:
            payload=b''.join(b'\0'+chunk[j:j+8] for j in range(0,len(chunk),8))
            chunks.append(struct.pack('<H',0xb000 | (len(payload)-1))+payload)
        else:
            chunks.append(b'\xff\x3f'+chunk.ljust(4096,b' '))
    return b'\x01'+b''.join(chunks)
def rebuild_ole(original,source):
    ole=olefile.OleFileIO(io.BytesIO(original)); proj=VBA_Project(ole,'','PROJECT','VBA/dir');list(proj.parse_modules())
    streams={tuple(p):ole.openstream(p).read() for p in ole.listdir()}
    directory=decompress_stream(streams[('VBA','dir')])
    for m in proj.modules:
        key=('VBA',m.streamname)
        code=source if m.name=='Inventering_Rapportautomatik' else m.code_raw
        streams[key]=compressed(code)
        needle=struct.pack('<HI',0x31,4)+struct.pack('<I',m.textoffset)
        assert needle in directory
        directory=directory.replace(needle,struct.pack('<HI',0x31,4)+b'\0'*4,1)
    streams[('VBA','dir')]=compressed(directory)
    # Invalidate version-specific performance cache; Excel must compile current source.
    streams[('VBA','_VBA_PROJECT')]=streams[('VBA','_VBA_PROJECT')][:2]+b'\xff\xff\x00\x00\x00'
    for p in streams:
        if p[-1].startswith('__SRP_'):streams[p]=b''
    paths={ole._find(list(p)):p for p in streams}
    dirs=ole._open(ole.first_dir_sector).read()
    entries=[bytearray(dirs[i:i+128]) for i in range(0,len(dirs),128)]
    END=0xfffffffe;FREE=0xffffffff;FAT=0xfffffffd
    sectors=[];fat=[];mini=bytearray();mf=[]
    def allocate(data):
        if not data:return END
        start=len(sectors);n=(len(data)+511)//512
        for i in range(n):sectors.append(data[i*512:(i+1)*512].ljust(512,b'\0'));fat.append(start+i+1 if i+1<n else END)
        return start
    for sid,path in paths.items():
        data=streams[path]
        if 0<len(data)<4096:
            start=len(mf);n=(len(data)+63)//64
            mini.extend(data.ljust(n*64,b'\0'));mf.extend([start+i+1 if i+1<n else END for i in range(n)])
        else:start=allocate(data)
        struct.pack_into('<IQ',entries[sid],116,start,len(data))
    rootstart=allocate(bytes(mini));struct.pack_into('<IQ',entries[0],116,rootstart,len(mini))
    mfs=allocate(struct.pack('<'+'I'*len(mf),*mf)) if mf else END
    dir_start=allocate(b''.join(entries))
    nfat=1
    while nfat*128<len(sectors)+nfat:nfat+=1
    assert nfat<=109
    fat_start=len(sectors);fat.extend([FAT]*nfat)
    fat.extend([FREE]*(nfat*128-len(fat)))
    fatdata=struct.pack('<'+'I'*len(fat),*fat)
    sectors.extend([fatdata[i*512:(i+1)*512] for i in range(nfat)])
    header=bytearray(512);header[:8]=bytes.fromhex('D0CF11E0A1B11AE1')
    struct.pack_into('<HHHHH',header,24,0x3e,3,0xfffe,9,6)
    struct.pack_into('<IIIIIIIII',header,40,0,nfat,dir_start,0,4096,mfs,(len(mf)*4+511)//512,END,0)
    struct.pack_into('<109I',header,76,*([fat_start+i for i in range(nfat)]+[FREE]*(109-nfat)))
    result=bytes(header)+b''.join(sectors)
    check=olefile.OleFileIO(io.BytesIO(result));p=VBA_Project(check,'','PROJECT','VBA/dir');list(p.parse_modules())
    for m in p.modules:
        expected=source if m.name=='Inventering_Rapportautomatik' else next(x.code_raw for x in proj.modules if x.name==m.name)
        assert m.code_raw.rstrip()==expected.rstrip(),m.name
    return result

def col(n):
    t=''
    while n:n,k=divmod(n-1,26);t=chr(65+k)+t
    return t

def write_cell(sheet,r,c,value,style=None):
    sd=sheet.find(q('sheetData'));row=next((x for x in sd if x.get('r')==str(r)),None)
    if row is None:
        row=E.Element(q('row'),r=str(r));sd.append(row)
    ref=f'{col(c)}{r}';cell=next((x for x in row if x.get('r')==ref),None)
    if cell is not None:row.remove(cell)
    cell=E.SubElement(row,q('c'),r=ref)
    if style is not None:cell.set('s',str(style))
    if isinstance(value,(int,float)):
        E.SubElement(cell,q('v')).text=str(value)
    elif value is not None:
        cell.set('t','inlineStr');E.SubElement(E.SubElement(cell,q('is')),q('t')).text=str(value)

def serialize(sheet):
    order = ['sheetPr', 'dimension', 'sheetViews', 'sheetFormatPr', 'cols', 'sheetData', 'sheetCalcPr', 'sheetProtection', 'protectedRanges', 'scenarios', 'autoFilter', 'sortState', 'dataConsolidate', 'customSheetViews', 'mergeCells', 'phoneticPr', 'conditionalFormatting', 'dataValidations', 'hyperlinks', 'printOptions', 'pageMargins', 'pageSetup', 'headerFooter', 'rowBreaks', 'colBreaks', 'customProperties', 'cellWatches', 'ignoredErrors', 'smartTags', 'drawing', 'legacyDrawing', 'legacyDrawingHF', 'drawingHF', 'picture', 'oleObjects', 'controls', 'webPublishItems', 'tableParts', 'extLst']
    sheet[:] = sorted(sheet, key=lambda child: order.index(E.QName(child).localname))
    sd=sheet.find(q('sheetData'))
    if sd is not None:
        sd[:]=sorted(sd,key=lambda x:int(x.get('r')))
        for row in sd:
            row[:]=sorted(row,key=lambda x:(len(re.match('[A-Z]+',x.get('r')).group()),re.match('[A-Z]+',x.get('r')).group()))
    return E.tostring(sheet,encoding='utf-8',xml_declaration=True)

def build(src,out,source):
    import openpyxl
    w=openpyxl.load_workbook(src,data_only=False,keep_vba=True)
    with zipfile.ZipFile(src) as z:parts={n:z.read(n) for n in z.namelist()}
    original_parts=dict(parts)
    wb=E.fromstring(parts['xl/workbook.xml']);rels=E.fromstring(parts['xl/_rels/workbook.xml.rels']);types=E.fromstring(parts['[Content_Types].xml'])
    targets={r.get('Id'):r.get('Target').lstrip('/') for r in rels}
    sheets={s.get('name'):('xl/'+targets[s.get('{'+R+'}id')] if not targets[s.get('{'+R+'}id')].startswith('xl/') else targets[s.get('{'+R+'}id')]) for s in wb.find(q('sheets'))}
    def getstyle(name,ref):
        cell=E.fromstring(original_parts[sheets[name]]).find('.//'+q('c')+'[@r="'+ref+'"]')
        return cell.get('s','0') if cell is not None else '0'
    rulepart=sheets['Regler'];rule=E.fromstring(parts[rulepart]);rawpart=sheets['Rådata'];raw=E.fromstring(parts[rawpart])
    write_cell(raw,1,13,'Återbruk i projekt',getstyle('Rådata','L1'))
    raw.find(q('dimension')).set('ref','A1:M1001')
    for name,data in list(parts.items()):
        if name.startswith('xl/tables/') and name.endswith('.xml'):
            tree=E.fromstring(data)
            if tree.get('name')=='tblRadata':
                tree.set('ref','A1:M1001')
                if tree.find(q('autoFilter')) is not None:tree.find(q('autoFilter')).set('ref','A1:M1001')
                cs=tree.find(q('tableColumns'));cs.set('count','13');E.SubElement(cs,q('tableColumn'),id='13',name='Återbruk i projekt');parts[name]=E.tostring(tree,encoding='utf-8',xml_declaration=True)
    dvs=raw.find(q('dataValidations'))
    if dvs is None:dvs=E.SubElement(raw,q('dataValidations'),count='0')
    dv=E.SubElement(dvs,q('dataValidation'),type='list',allowBlank='1',showErrorMessage='1',sqref='M2:M1048576');E.SubElement(dv,q('formula1')).text='"JA,NEJ,EJ FASTSTÄLLT"';dvs.set('count',str(len(dvs)))
    # Existing formatting, old rules, comments and all worksheet relationships retained.
    headers=['Regel-ID','Aktiv','Prioritet','Regelnamn','Logik']+[x for n in ['1','2','3','X'] for x in ['Fält '+n,'Metod '+n,'Värde '+n]]+['Lokalt delområde','Rubriknivå 1','Rubriknivå 2','Rubriknivå 3','Enhet','Ordning','Kommentar','Beskrivning']
    cats=[]
    old=w['Regler'].tables['tblRegler'];oldstart=openpyxl.utils.range_boundaries(old.ref)[1]
    for n,row in enumerate(w['Regler'].iter_rows(min_row=oldstart+1,max_row=openpyxl.utils.range_boundaries(old.ref)[3],values_only=True),1):
        if not isinstance(row[0],(int,float)):continue
        d=dict(zip(headers,['KAT-'+f'{n:03}',row[1],row[0],str(row[3] or 'Reservregel'),'1']))
        d.update({'Fält 1':row[2],'Metod 1':'Innehåller','Värde 1':row[3],'Lokalt delområde':row[4],'Rubriknivå 1':row[5],'Rubriknivå 2':row[6],'Rubriknivå 3':row[7],'Enhet':row[8],'Ordning':row[9],'Kommentar':row[10]})
        cats.append(d)
    common=headers[:18]
    gh=common+['Huvudkategori','Åtgärd','LITTRERING','Materialtyp','SKICK','KOMMENTAR','CLASS','DELOMRÅDE','Beskrivning']
    groups=[]
    for i,category in enumerate(['UTRUSTNING','NATURSTENSMATERIAL','BETONGMATERIAL','PLANTERING',''],1):
        groups.append({'Regel-ID':f'GRP-{i:03}','Aktiv':'Ja','Prioritet':10 if category else 0,'Regelnamn':category or 'Standard','Logik':'1','Huvudkategori':category,'Åtgärd':'Grundregel','LITTRERING':'Ja','Materialtyp':'Nej' if category=='UTRUSTNING' else 'Ja','SKICK':'Nej','KOMMENTAR':'Nej','CLASS':'Nej','DELOMRÅDE':'Nej'})
    groups.append({'Regel-ID':'GRP-006','Aktiv':'Nej','Prioritet':100,'Regelnamn':'Exempel CS1 + skick','Logik':'1','Fält 1':'LITTRERING','Metod 1':'Är exakt','Värde 1':'CS1','Åtgärd':'Komplettera','SKICK':'Ja'})
    uh=common+['Enhet','Beskrivning']
    units=[{'Regel-ID':'ENH-001','Aktiv':'Ja','Prioritet':200,'Regelnamn':'SMG2 i löpmeter','Logik':'1','Fält 1':'LITTRERING','Metod 1':'Är exakt','Värde 1':'SMG2','Enhet':'lpm'}, {'Regel-ID':'ENH-002','Aktiv':'Ja','Prioritet':100,'Regelnamn':'Hårdgjord yta','Logik':'1','Fält 1':'CLASS','Metod 1':'Innehåller','Värde 1':'Hårdgjort','Enhet':'m²'}]
    rh=['Regel-ID','Aktiv','Rådatavärde','Rapportvärde'];reuse=[dict(zip(rh,[f'ATER-{i:03}','Ja',a,b])) for i,(a,b) in enumerate([('JA','JA'),('NEJ','NEJ'),('EJ FASTSTÄLLT','EJ FASTSTÄLLT'),('','EJ FASTSTÄLLT')],1)]
    relpath=rulepart.rsplit('/',1)[0]+'/_rels/'+rulepart.rsplit('/',1)[1]+'.rels'
    rr=E.fromstring(parts[relpath]) if relpath in parts else E.Element('{'+PK+'}Relationships')
    tp=rule.find(q('tableParts'))
    if tp is None:tp=E.SubElement(rule,q('tableParts'),count='0')
    existingids=[int(E.fromstring(v).get('id')) for k,v in parts.items() if k.startswith('xl/tables/') and k.endswith('.xml')]
    nextid=max(existingids)+1
    dvroot=rule.find(q('dataValidations'))
    if dvroot is None:
        dvroot=E.Element(q('dataValidations'));rule.insert(list(rule).index(tp),dvroot)
    config=[]
    for name,start,hs,rows in [('tblKategorierV2',40,headers,cats),('tblGrupperV2',100,gh,groups),('tblAterbrukV2',120,rh,reuse),('tblEnheterV2',135,uh,units)]:
        write_cell(rule,start-2,1,{'tblKategorierV2':'VERSION 2 – KATEGORISERINGSREGLER','tblGrupperV2':'GRUPPERINGSREGLER','tblAterbrukV2':'ÅTERBRUK I PROJEKT','tblEnheterV2':'ENHETSREGLER'}[name],getstyle('Regler','A1'))
        for c,h in enumerate(hs,1):write_cell(rule,start,c,h,getstyle('Regler','A'+str(oldstart)))
        for r,d in enumerate(rows,start+1):
            if 'Beskrivning' in hs:
                d['Beskrivning']='; '.join(f'{d.get("Fält "+n,"")} {d.get("Metod "+n,"")} {d.get("Värde "+n,"")}' for n in ['1','2','3'] if d.get('Fält '+n)) or 'Grundregel / alla objekt'
            for c,h in enumerate(hs,1):write_cell(rule,r,c,d.get(h))
        ref=f'A{start}:{col(len(hs))}{start+len(rows)}'
        table=E.Element(q('table'),id=str(nextid),name=name,displayName=name,ref=ref,totalsRowShown='0')
        E.SubElement(table,q('autoFilter'),ref=ref);cs=E.SubElement(table,q('tableColumns'),count=str(len(hs)))
        for c,h in enumerate(hs,1):E.SubElement(cs,q('tableColumn'),id=str(c),name=h)
        E.SubElement(table,q('tableStyleInfo'),name='TableStyleMedium2',showFirstColumn='0',showLastColumn='0',showRowStripes='1',showColumnStripes='0')
        tablepath=f'xl/tables/table{nextid}.xml';parts[tablepath]=E.tostring(table,encoding='utf-8',xml_declaration=True)
        rid='rIdV2Table'+str(nextid);E.SubElement(rr,'{'+PK+'}Relationship',Id=rid,Type=R+'/table',Target='../tables/table'+str(nextid)+'.xml')
        E.SubElement(tp,q('tablePart'),{'{'+R+'}id':rid});E.SubElement(types,'{'+CT+'}Override',PartName='/'+tablepath,ContentType='application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml')
        for c,h in enumerate(hs,1):
            options=None
            if h=='Aktiv' or h in ['LITTRERING','Materialtyp','SKICK','KOMMENTAR','CLASS','DELOMRÅDE']:options='Ja,Nej'
            if h.startswith('Metod '):options='Innehåller,Är exakt,Börjar med,Innehåller inte'
            if h.startswith('Fält '):options='LITTRERING,Materialtyp,CLASS,DELOMRÅDE,SKICK,Återbruk i projekt,KOMMENTAR'
            if h=='Logik':options='1,1 OCH 2,1 ELLER 2,1 OCH 2 OCH 3,1 ELLER 2 ELLER 3,1 OCH (2 ELLER 3),(1 OCH 2) ELLER 3,1 ELLER (2 OCH 3),(1 ELLER 2) OCH 3'
            if h=='Enhet':options='Auto,antal,m²,lpm'
            if h=='Åtgärd':options='Grundregel,Ersätt,Komplettera'
            if options:
                dv=E.SubElement(dvroot,q('dataValidation'),type='list',allowBlank='1',showErrorMessage='1',sqref=f'{col(c)}{start+1}:{col(c)}{start+len(rows)}');E.SubElement(dv,q('formula1')).text='"'+options+'"'
        config.append((name,hs,rows));nextid+=1
    tp.set('count',str(len(tp)));dvroot.set('count',str(len(dvroot)));rule.find(q('dimension')).set('ref','A1:AA150')
    write_cell(rule,1,1,'Regler V2 från rad 40. Gamla regler ovan bevaras som referens.')
    parts[relpath]=E.tostring(rr,encoding='utf-8',xml_declaration=True)
    pref=E.fromstring(parts[sheets['Rapportval']]);write_cell(pref,26,1,'Återbruksrapportnivå');write_cell(pref,26,2,1);write_cell(pref,27,1,'Nivå 2 och 3 ej implementerade; procent styr inte beslut.')
    pref.find(q('dimension')).set('ref','A1:F27')
    readme=E.fromstring(parts[sheets['Läs mig']]);write_cell(readme,30,1,'TESTVERSION 2 – kör InstalleraKnappar, därefter UppdateraRapporter.')
    write_cell(readme,31,1,'Återbruk i projekt saknades i källfilen. Fyll JA / NEJ / EJ FASTSTÄLLT i Rådata; inga beslut har gissats.')
    write_cell(readme,32,1,'Regler V2 finns från rad 40. Rapporterna i denna fil är ännu originalets rapporter tills makrot körs.')
    write_cell(readme,33,1,'Rapportuppdatering säkerhetskopierar hela filen. Egna rapportanteckningar överförs inte automatiskt.')
    write_cell(readme,34,1,'VBA-källkod och filstruktur kontrollerade; Excel-kompilering och full rapportkörning behöver provas i Windows.')
    readme.find(q('dimension')).set('ref','A1:G34')
    # Maintain worksheet element ordering for raw validation (must precede tableParts).
    if raw.find(q('tableParts')) is not None:
        raw.remove(dvs);raw.insert(list(raw).index(raw.find(q('tableParts'))),dvs)
    for part,tree in [(rawpart,raw),(rulepart,rule),(sheets['Rapportval'],pref),(sheets['Läs mig'],readme)]:parts[part]=serialize(tree)
    calc=wb.find(q('calcPr'))
    if calc is None:calc=E.SubElement(wb,q('calcPr'))
    calc.set('fullCalcOnLoad','1');calc.set('forceFullCalc','1')
    parts['xl/workbook.xml']=E.tostring(wb,encoding='utf-8',xml_declaration=True)
    parts['[Content_Types].xml']=E.tostring(types,encoding='utf-8',xml_declaration=True)
    parts['xl/vbaProject.bin']=rebuild_ole(parts['xl/vbaProject.bin'],source)
    with zipfile.ZipFile(out,'w',zipfile.ZIP_DEFLATED) as z:
        for n,data in parts.items():z.writestr(n,data)
    with zipfile.ZipFile(out) as z:assert z.testzip() is None
    check=openpyxl.load_workbook(out,keep_vba=True)
    assert check['Rådata']['M1'].value=='Återbruk i projekt'
    assert set(['tblKategorierV2','tblGrupperV2','tblAterbrukV2','tblEnheterV2']).issubset(check['Regler'].tables)
    changed=[n for n in original_parts if original_parts[n]!=parts[n]]
    # All existing relationships, pictures, buttons, worksheet reports and template remain byte-identical.
    for n in original_parts:
        if any(x in n for x in ['drawings/','media/','activeX/','ctrlProps/']):assert original_parts[n]==parts[n],n
    for name in ['Rapportmall','Sammanställning']+[s for s in sheets if s.startswith('OMR -')]:assert original_parts[sheets[name]]==parts[sheets[name]]
    manifest={'original_sha256':hashlib.sha256(Path(src).read_bytes()).hexdigest(),'test_sha256':hashlib.sha256(Path(out).read_bytes()).hexdigest(),'modified_parts':changed,'tables':config,'excel_execution':'NOT RUN'}
    Path(str(out)+'.json').write_text(json.dumps(manifest,ensure_ascii=False,indent=2))
    print(json.dumps({k:v for k,v in manifest.items() if k!='tables'},ensure_ascii=False,indent=2))
if __name__=='__main__':build(sys.argv[1],sys.argv[2],Path(sys.argv[3]).read_bytes())
