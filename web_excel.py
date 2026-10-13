"""Portable in-memory XLSX exports. No source templates or user data on disk.

XlsxWriter supplies cached formulas and native chart parts. The pivot package
is added as standard OOXML, keeping the calculated values intact.
"""
from collections import defaultdict
import io
import zipfile
import threading
from xml.etree import ElementTree as ET
import xlsxwriter
from analysis_engine import MONTHS, sort_error, DATA_LABELS

MAIN='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
REL='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PKG='http://schemas.openxmlformats.org/package/2006/relationships'
CT='http://schemas.openxmlformats.org/package/2006/content-types'
EXPORT_LOCK=threading.RLock()

def element(tag,**attributes):
    return ET.Element('{'+MAIN+'}'+tag,{k:str(v) for k,v in attributes.items()})

def child(parent,tag,**attributes):
    item=element(tag,**attributes);parent.append(item);return item

def xml(root):
    ET.register_namespace('',root.tag.split('}')[0][1:])
    ET.register_namespace('r',REL)
    return ET.tostring(root,encoding='utf-8',xml_declaration=True)

def error_value(value):return int(str(value)) if str(value).isdigit() else str(value)

def add_pivot(data,rows,group,pivot_index):
    """Add a refreshable native stock/error PivotTable with cached source rows."""
    if not rows:return data
    headers=['Stok Kodu','Ay','Fiş Grubu','Hata Kodu','Miktar','Kaynak','Orijinal Hata','Mod Ataması']
    values=[[r['code'],MONTHS[r['month']-1],r['voucher'],error_value(r['error']),r.get('pivot_qty',r['qty']),r['source'],str(r['original_error'] or ''),'Evet' if r['assigned'] else 'Hayır'] for r in rows]
    shared=[]
    for i in range(8):
        items=list(dict.fromkeys(row[i] for row in values)) if i in (0,3) else []
        if i==0:items.sort()
        if i==3:items.sort(key=sort_error)
        shared.append(items)
    indexes=[{v:i for i,v in enumerate(items)} for items in shared]
    cache=element('pivotCacheDefinition',recordCount=len(rows),saveData='1',refreshOnLoad='0',createdVersion='6',refreshedVersion='6',minRefreshableVersion='3')
    cache.set('{'+REL+'}id','rId1')
    child(child(cache,'cacheSource',type='worksheet'),'worksheetSource',sheet='Hurda Detayı',ref=f'A1:H{len(rows)+1}')
    fields=child(cache,'cacheFields',count=8)
    for i,name in enumerate(headers):
        si=child(child(fields,'cacheField',name=name),'sharedItems',**({'count':len(shared[i])} if i in (0,3) else {}))
        if i==4:si.attrib.update(containsSemiMixedTypes='0',containsString='0',containsNumber='1')
        if i==3:
            numbers=[v for v in shared[i] if isinstance(v,int)]
            if numbers:si.attrib.update(containsNumber='1',containsInteger='1',containsString='1' if len(numbers)<len(shared[i]) else '0',containsMixedTypes='1' if len(numbers)<len(shared[i]) else '0',containsSemiMixedTypes='1' if len(numbers)<len(shared[i]) else '0',minValue=str(min(numbers)),maxValue=str(max(numbers)))
        if i==6:si.set('containsBlank','1')
        for v in shared[i]:child(si,'n' if isinstance(v,int) else 's',v=v)
    records=element('pivotCacheRecords',count=len(rows))
    for row in values:
        rr=child(records,'r')
        for i,value in enumerate(row):
            if i in (0,3):child(rr,'x',v=indexes[i][value])
            elif i==4:child(rr,'n',v=value)
            elif value=='':child(rr,'m')
            else:child(rr,'s',v=value)
    pivot=element('pivotTableDefinition',name='HataPivot'+group,cacheId=1,dataCaption='Toplam Hurda',rowGrandTotals='1',colGrandTotals='1',compact='0',compactData='0',outline='0',createdVersion='6',updatedVersion='6',minRefreshableVersion='3',useAutoFormatting='1')
    last_col=xlsxwriter.utility.xl_col_to_name(len(shared[3])+1)
    child(pivot,'location',ref=f'A1:{last_col}{len(shared[0])+3}',firstHeaderRow=1,firstDataRow=2,firstDataCol=1)
    pf=child(pivot,'pivotFields',count=8)
    for i in range(8):
        attributes={'showAll':'0'}
        if i in (0,3):attributes.update(axis='axisRow' if i==0 else 'axisCol',sortType='ascending')
        if i==4:attributes['dataField']='1'
        field=child(pf,'pivotField',**attributes)
        if i in (0,3):
            items=child(field,'items',count=len(shared[i])+1)
            for j in range(len(shared[i])):child(items,'item',x=j)
            child(items,'item',t='default')
    for axis,i in [('row',0),('col',3)]:
        child(child(pivot,axis+'Fields',count=1),'field',x=i)
        items=child(pivot,axis+'Items',count=len(shared[i])+1)
        for j in range(len(shared[i])):child(child(items,'i'),'x',v=j)
        child(child(items,'i',t='grand'),'x')
    child(child(pivot,'dataFields',count=1),'dataField',name='Toplam Hurda',fld=4,subtotal='sum',baseField=0,baseItem=0,numFmtId=4)
    child(pivot,'pivotTableStyleInfo',name='PivotStyleMedium9',showRowHeaders='1',showColHeaders='1',showRowStripes='0',showColStripes='0',showLastColumn='1')
    with zipfile.ZipFile(io.BytesIO(data)) as z:parts={n:z.read(n) for n in z.namelist()}
    sheet_path=f'xl/worksheets/sheet{pivot_index}.xml'
    sheet=ET.fromstring(parts[sheet_path])
    pp=child(sheet,'pivotParts',count=1);child(pp,'pivotPart',**{'{'+REL+'}id':'rIdPivot1'})
    parts[sheet_path]=xml(sheet)
    rel_path=f'xl/worksheets/_rels/sheet{pivot_index}.xml.rels'
    sr=ET.fromstring(parts[rel_path]) if rel_path in parts else ET.Element('{'+PKG+'}Relationships')
    ET.SubElement(sr,'{'+PKG+'}Relationship',Id='rIdPivot1',Type=REL+'/pivotTable',Target='../pivotTables/pivotTable1.xml');parts[rel_path]=xml(sr)
    wb=ET.fromstring(parts['xl/workbook.xml']);pc=element('pivotCaches');child(pc,'pivotCache',cacheId=1,**{'{'+REL+'}id':'rIdCache1'})
    calc=wb.find('{'+MAIN+'}calcPr');wb.insert(list(wb).index(calc)+1 if calc is not None else len(wb),pc)
    parts['xl/workbook.xml']=xml(wb)
    wr=ET.fromstring(parts['xl/_rels/workbook.xml.rels']);ET.SubElement(wr,'{'+PKG+'}Relationship',Id='rIdCache1',Type=REL+'/pivotCacheDefinition',Target='pivotCache/pivotCacheDefinition1.xml');parts['xl/_rels/workbook.xml.rels']=xml(wr)
    for path,kind,target in [('xl/pivotTables/_rels/pivotTable1.xml.rels','pivotCacheDefinition','../pivotCache/pivotCacheDefinition1.xml'),('xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels','pivotCacheRecords','pivotCacheRecords1.xml')]:
        root=ET.Element('{'+PKG+'}Relationships');ET.SubElement(root,'{'+PKG+'}Relationship',Id='rId1',Type=REL+'/'+kind,Target=target);parts[path]=xml(root)
    parts['xl/pivotTables/pivotTable1.xml']=xml(pivot)
    parts['xl/pivotCache/pivotCacheDefinition1.xml']=xml(cache)
    parts['xl/pivotCache/pivotCacheRecords1.xml']=xml(records)
    types=ET.fromstring(parts['[Content_Types].xml'])
    for path,kind in [('xl/pivotTables/pivotTable1.xml','pivotTable'),('xl/pivotCache/pivotCacheDefinition1.xml','pivotCacheDefinition'),('xl/pivotCache/pivotCacheRecords1.xml','pivotCacheRecords')]:
        ET.SubElement(types,'{'+CT+'}Override',PartName='/'+path,ContentType='application/vnd.openxmlformats-officedocument.spreadsheetml.'+kind+'+xml')
    parts['[Content_Types].xml']=xml(types)
    output=io.BytesIO()
    with zipfile.ZipFile(output,'w',zipfile.ZIP_DEFLATED) as z:
        for path,content in parts.items():z.writestr(path,content)
    return output.getvalue()

def workbook(result,group,pareto=False):
    output=io.BytesIO()
    book=xlsxwriter.Workbook(output,{'in_memory':True,'strings_to_formulas':False,'strings_to_urls':False})
    head=book.add_format({'bold':True,'bg_color':'#17365D','font_color':'white','text_wrap':True,'border':1})
    num=book.add_format({'num_format':'#,##0.######'})
    money=book.add_format({'num_format':'#,##0.00'})
    pct=book.add_format({'num_format':'0.00%'})
    total=book.add_format({'bold':True,'bg_color':'#DFE9F6','num_format':'#,##0.######','top':1})
    def sheet(name,headers):
        s=book.add_worksheet(name);s.hide_gridlines(2);s.freeze_panes(1,1);s.set_column(0,0,24)
        if len(headers)>1:s.set_column(1,len(headers)-1,18,num)
        s.set_row(0,34);s.write_row(0,0,headers,head);s.set_landscape();s.fit_to_pages(1,0);s.repeat_rows(0);s.set_header(f"&C{group['name']} · {result['year']} · {' / '.join(result['vouchers'])}");return s
    rows=[r for r in result['detail'] if r['code'].startswith(group['group'])]
    cable=group['group']=='202'
    factors={s['code']:s['grams']/1000 if cable and group['pareto_unit']=='kg' else 1 for s in group['stocks']}
    rows=[{**r,'pivot_qty':r['qty']*factors[r['code']]} for r in rows]
    if not pareto:
        summary=sheet('GRAFİK VERİSİ',['Ay',f"Tüketim ({group['unit']})",f"Hurda ({group['unit']})",'PPM','Hedef PPM','Maliyet TL','Maliyet EUR','Veri durumu'])
        for i,m in enumerate(group['monthly'],1):
            available=m.get('data_status','ok') not in ('no_source','no_scope','all_excluded')
            summary.write_row(i,0,[m['label'],m['consumption'] if available else None,m['scrap'] if available else None,m['ppm'],group['target'],m['cost_tl'] if available else None,m['cost_eur'] if available else None,DATA_LABELS.get(m.get('data_status','ok'))])
        n=len(group['monthly'])
        summary.write(n+1,0,'TOPLAM',total)
        if group['stocks']:
            for col in (1,2,5,6):
                letter=xlsxwriter.utility.xl_col_to_name(col);value=sum(m[['','consumption','scrap','','','cost_tl','cost_eur'][col]] for m in group['monthly'])
                summary.write_formula(n+1,col,f'=SUM({letter}2:{letter}{n+1})',total,value)
        summary.write(n+1,7,DATA_LABELS.get(group.get('data_status','ok')))
        summary.set_column(7,7,38)
        if group['consumption']:summary.write_formula(n+1,3,f'=C{n+2}/B{n+2}*1000000',total,group['ppm'])
        chart=book.add_chart({'type':'column'});chart.add_series({'name':'Tüketim','categories':['GRAFİK VERİSİ',1,0,n,0],'values':['GRAFİK VERİSİ',1,1,n,1]})
        line=book.add_chart({'type':'line'});line.add_series({'name':'Gerçekleşen PPM','categories':['GRAFİK VERİSİ',1,0,n,0],'values':['GRAFİK VERİSİ',1,3,n,3],'y2_axis':True,'marker':{'type':'circle','size':6}})
        if group['target'] is not None:line.add_series({'name':'Hedef PPM','categories':['GRAFİK VERİSİ',1,0,n,0],'values':['GRAFİK VERİSİ',1,4,n,4],'y2_axis':True})
        chart.combine(line);chart.set_title({'name':f"{group['name']} · {result['year']} · PPM"});chart.set_y2_axis({'name':'PPM','min':0});chart.set_y_axis({'name':f"Tüketim ({group['unit']})",'min':0});summary.insert_chart('I2',chart,{'x_scale':1.4,'y_scale':1.3})
    for month in range(1,13):
        name=MONTHS[month-1].upper()
        if pareto:
            s=sheet(name,['Hata Kodu','Hata Açıklaması','Miktar ('+group['pareto_unit']+')','Pay','Kümülatif','Maliyet TL'])
            if month not in result['months']:continue
            status=next(m.get('data_status','ok') for m in group['monthly'] if m['month']==month)
            if status in ('no_source','no_scope','all_excluded'):s.write(1,7,DATA_LABELS[status])
            quantities=defaultdict(float);costs=defaultdict(float)
            stocks={v['code']:v for v in group['stocks'] if v['month']==month}
            for r in rows:
                if r['month']!=month:continue
                factor=stocks[r['code']]['grams']/1000 if cable and group['pareto_unit']=='kg' else 1
                quantities[str(r['error'])]+=r['qty']*factor;costs[str(r['error'])]+=r['qty']*stocks[r['code']]['unit_tl']
            descriptions={r['error']:r['description'] for r in group['pareto']}
            ordered=sorted(quantities,key=lambda e:(-quantities[e],sort_error(e)));denom=sum(quantities.values());cumulative=0
            for i,e in enumerate(ordered,1):
                share=quantities[e]/denom if denom else 0;cumulative+=share
                s.write_row(i,0,[e,descriptions.get(e,e),quantities[e]])
                s.write_formula(i,3,f'=C{i+1}/SUM(C$2:C${len(ordered)+1})',pct,share)
                s.write_formula(i,4,f'=SUM(D$2:D{i+1})',pct,cumulative);s.write(i,5,costs[e],money)
            if ordered:
                end=len(ordered)+1;s.write(end,0,'TOPLAM',total)
                for col,value in [(2,denom),(3,1),(5,sum(costs.values()))]:
                    letter=xlsxwriter.utility.xl_col_to_name(col);s.write_formula(end,col,f'=SUM({letter}2:{letter}{end})',pct if col==3 else total,value)
                chart=book.add_chart({'type':'column'});chart.add_series({'name':'Hurda payı','categories':[name,1,0,len(ordered),0],'values':[name,1,3,len(ordered),3]})
                line=book.add_chart({'type':'line'});line.add_series({'name':'Kümülatif','categories':[name,1,0,len(ordered),0],'values':[name,1,4,len(ordered),4],'y2_axis':True,'marker':{'type':'circle'}});chart.combine(line);chart.set_y_axis({'num_format':'0%','min':0,'max':1});chart.set_y2_axis({'num_format':'0%','min':0,'max':1});chart.set_title({'name':f"{group['name']} · {name} · Pareto"});s.insert_chart('H2',chart)
        else:
            headers=['Stok Kodu','Birim','Tüketim','Tüketim kg','Gramaj g/birim','Hurda','Hurda kg','Birim TL','Maliyet TL','Birim EUR','Maliyet EUR'] if cable else ['Stok Kodu','Birim','Tüketim','Hurda','Birim TL','Maliyet TL','Birim EUR','Maliyet EUR']
            s=sheet(name,headers);items=[v for v in group['stocks'] if v['month']==month]
            if month in result['months'] and not items:
                status=next(m.get('data_status','ok') for m in group['monthly'] if m['month']==month)
                s.write(1,len(headers)+1,DATA_LABELS.get(status,'Analize uygun veri yok'))
            for i,v in enumerate(items,1):
                r=i+1
                if cable:
                    s.write_row(i,0,[v['code'],v['unit'],v['consumption']]);s.write_formula(i,3,f'=C{r}*E{r}/1000',num,v['consumption_kg']);s.write_row(i,4,[v['grams'],v['scrap']]);s.write_formula(i,6,f'=E{r}*F{r}/1000',num,v['scrap_kg']);s.write(i,7,v['unit_tl']);s.write_formula(i,8,f'=F{r}*H{r}',money,v['cost_tl']);s.write(i,9,v['unit_eur']);s.write_formula(i,10,f'=F{r}*J{r}',money,v['cost_eur'])
                else:
                    s.write_row(i,0,[v['code'],'AD',v['consumption'],v['scrap'],v['unit_tl']]);s.write_formula(i,5,f'=D{r}*E{r}',money,v['cost_tl']);s.write(i,6,v['unit_eur']);s.write_formula(i,7,f'=D{r}*G{r}',money,v['cost_eur'])
            if items:
                s.write(len(items)+1,0,'TOPLAM',total)
                cols={2:'consumption',3:'consumption_kg',5:'scrap',6:'scrap_kg',8:'cost_tl',10:'cost_eur'} if cable else {2:'consumption',3:'scrap',5:'cost_tl',7:'cost_eur'}
                for col,key in cols.items():
                    if cable and col in (2,5) and any(v['unit']!='MT' for v in items):continue
                    letter=xlsxwriter.utility.xl_col_to_name(col);s.write_formula(len(items)+1,col,f'=SUM({letter}2:{letter}{len(items)+1})',total,sum(v[key] for v in items))
    pivot_index=None
    if not pareto:
        detail=sheet('Hurda Detayı',['Stok Kodu','Ay','Fiş Grubu','Hata Kodu','Miktar','Kaynak','Orijinal Hata','Mod Ataması','Kaynak Miktarı','Kaynak Birimi'])
        for i,r in enumerate(rows,1):detail.write_row(i,0,[r['code'],MONTHS[r['month']-1],r['voucher'],error_value(r['error']),r['pivot_qty'],r['source'],str(r['original_error'] or ''),'Evet' if r['assigned'] else 'Hayır',r['qty'],r['unit']])
        detail.write(0,11,f"Pivot miktar birimi: {group['pivot_unit']}")
        pivot=sheet('Hata Kodu Pivot',[f"Toplam Hurda ({group['pivot_unit']})"]);pivot_index=len(book.worksheets())
        pivot.write_row(1,0,['Stok Kodu']+[error_value(e) for e in group['pivot_codes']]+['TOPLAM'],head)
        for i,r in enumerate(group['pivot'],2):pivot.write_row(i,0,[r['code']]+[r['errors'].get(e,0) for e in group['pivot_codes']]+[r['total']])
        if group['pivot']:
            pivot.write_row(len(group['pivot'])+2,0,['TOPLAM']+[sum(r['errors'].get(e,0) for r in group['pivot']) for e in group['pivot_codes']]+[sum(r['total'] for r in group['pivot'])],total)
        excluded=sheet('Analiz Dışı Kayıtlar',['Stok Kodu','Ay','Neden'])
        for i,r in enumerate((v for v in result['issues'] if v['code'].startswith(group['group'])),1):excluded.write_row(i,0,[r['code'],MONTHS[r['month']-1],r['reason']])
    checks=[r for r in result.get('reconciliation',[]) if r['group']==group['group']]
    if checks:
        check=sheet('Kaynak Kontrolü',['Yıl','Ay','Fiş','Ana miktar (kaynak birimi)','Birleşik miktar (kaynak birimi)','Net ek miktar (kaynak birimi)','Kaynak dosya'])
        for i,r in enumerate(checks,1):check.write_row(i,0,[r['year'],MONTHS[r['month']-1],r['voucher'],r['original_qty'],r['combined_qty'],r['net_added_qty'],r['source']])
        check.set_column(3,5,25,num);check.set_column(6,6,30)
        check.write(len(checks)+2,0,'Birleşik tablo ana miktarı zaten içerir. Ana kayıtlar tekrar eklenmedi.')
        check.write(len(checks)+3,0,'Kaynak miktarları, eksik verili stoklar hariç tutulmadan önceki miktarlardır.')
    period_check=sheet('Dönem Kontrolü',['Ay','Kaynak dosyadaki kayıtlar','Seçilen gruptaki kayıtlar','Analize giren kayıtlar','Veri durumu'])
    for i,m in enumerate(group['monthly'],1):
        source=next((c['source_rows'] for c in result.get('coverage',[]) if c['month']==m['month']),None)
        period_check.write_row(i,0,[m['label'],source,m.get('source_rows'),m.get('eligible_rows'),DATA_LABELS.get(m.get('data_status','ok'))])
    period_check.set_column(4,4,42)
    book.close();data=output.getvalue()
    return add_pivot(data,rows,group['group'],pivot_index) if pivot_index else data

def export_zip(result):
    with EXPORT_LOCK:return _export_zip(result)

def _export_zip(result):
    output=io.BytesIO()
    with zipfile.ZipFile(output,'w',zipfile.ZIP_DEFLATED) as z:
        for group in result['groups']:
            for pareto in (False,True):
                name=f"{group['name'].upper()} HURDA {'PARETO' if pareto else 'ANALİZİ'}.xlsx"
                z.writestr(name,workbook(result,group,pareto))
    return output.getvalue()
