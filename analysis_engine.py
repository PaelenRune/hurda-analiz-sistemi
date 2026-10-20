"""Pure, session-only scrap analysis. Source workbooks are never modified."""
from __future__ import annotations
import calendar
import math
import re
import unicodedata
from collections import Counter, defaultdict
from datetime import datetime, date
from pathlib import Path

MONTHS = ['Ocak','Şubat','Mart','Nisan','Mayıs','Haziran','Temmuz','Ağustos','Eylül','Ekim','Kasım','Aralık']
GROUPS = {'202':'Kablo', '204':'Terminal', '206':'IDC'}
DATA_LABELS={'no_source':'Kaynak dosyada bu dönem yok','no_scope':'Seçilen kapsamda hurda kaydı yok','all_excluded':'Tüm hurda kayıtları eksik veri nedeniyle çıkarıldı','partial':'Eksik veriler nedeniyle kısmi analiz','ok':'Veri kontrolü tamamlandı'}
ERRORS = dict(enumerate(['Terminal Çakım Hatası','IDC Çakım Hatası','Kablo Boyu Hatalı','Taşıma Kaynaklı','Insertion Hatası','Potalama Kaynaklı','Tedarikçi Kaynaklı','Makara Sonu Firesi','Kablo Uç Açımı Hatalı','Yanlış Malzeme Kullanımı','Bıçak İzi','Ayar Firesi','İşçilik Firesi','Sipariş Fazlası Olarak Gelen Kablolar','Revizyon','Ezik / Kırık / Deformeli','Proje Kaynaklı','Operatör Hatası','Yanlış Onay','Elektriksel Test','Kilo Hatası','IDC Kesme Hurdası','Makina Kabul','Enjeksiyon','Fire','Diğer'],1))

class InputError(ValueError): pass
class TableNotFound(InputError): pass
EXCEL_ERRORS={'#NULL!','#DIV/0!','#VALUE!','#REF!','#NAME?','#NUM!','#N/A','#SPILL!','#CALC!'}

def norm(value):
    value = str(value or '').replace('ı','i').replace('İ','I')
    return ''.join(c for c in unicodedata.normalize('NFKD',value) if not unicodedata.combining(c)).upper().strip()

def code(value):
    if value is None: return ''
    if isinstance(value,str) and value.strip().upper() in EXCEL_ERRORS:raise InputError('Excel hata değeri var: '+value+'. Kaynak hücreyi düzeltin.')
    if isinstance(value,(float,int)) and float(value).is_integer(): return str(int(value))
    return str(value).strip().upper()

def numeric(value, label='Değer'):
    if value is None or value == '': return None
    if isinstance(value,str):
        value=value.strip().replace(' ','')
        if ',' in value: value=value.replace('.','').replace(',','.')
    try: result=float(value)
    except (ValueError,TypeError): raise InputError(f'{label}: sayısal olmayan değer ({value}).')
    if not math.isfinite(result): raise InputError(f'{label}: geçersiz sayı.')
    return result

def integer(value,label,lower,upper):
    if isinstance(value,bool):raise InputError(f'{label}: tam sayı seçin.')
    n=numeric(value,label)
    if n is None or n!=int(n) or not lower<=n<=upper:raise InputError(f'{label}: {lower}–{upper} arasında tam sayı seçin.')
    return int(n)

def parse_date(value):
    if isinstance(value,datetime): return value
    if isinstance(value,date): return datetime.combine(value,datetime.min.time())
    if not value: return None
    for fmt in ('%Y-%m-%d %H:%M:%S','%Y-%m-%d','%d.%m.%Y','%d/%m/%Y','%Y-%m-%dT%H:%M:%S'):
        try: return datetime.strptime(str(value).strip(),fmt)
        except ValueError: pass
    return None

def voucher(value):
    raw=code(value)
    if raw.startswith('HRD'): return 'HRD'
    if raw.startswith('Ç'): return 'Ç'
    if raw.startswith('Y'): return 'Y'
    return None

def read_sheets(path, filename=None):
    name=Path(filename or path) if not hasattr(path,'read') else Path(filename or 'input.xlsx')
    if name.suffix.lower()=='.xls':
        import xlrd
        book=xlrd.open_workbook(file_contents=path.read()) if hasattr(path,'read') else xlrd.open_workbook(str(path))
        sheets={}
        for sheet in book.sheets():
            rows=[]
            for r in range(sheet.nrows):
                row=[]
                for c in range(sheet.ncols):
                    cell=sheet.cell(r,c)
                    if cell.ctype==xlrd.XL_CELL_ERROR: row.append(xlrd.error_text_from_code.get(cell.value,'#VALUE!'))
                    elif cell.ctype==xlrd.XL_CELL_DATE: row.append(xlrd.xldate.xldate_as_datetime(cell.value,book.datemode))
                    else: row.append(cell.value if cell.value!='' else None)
                rows.append(row)
            sheets[sheet.name]=rows
        book.release_resources()
        return sheets
    import openpyxl
    book=openpyxl.load_workbook(path,read_only=True,data_only=True)
    try:
        sheets={}
        for s in book:
            s.reset_dimensions()  # ERP exports may incorrectly declare dimension A1.
            sheets[s.title]=list(s.values)
        return sheets
    finally: book.close()

def find_table(sheets, required, preferred=None):
    ordered=list(sheets)
    if preferred:
        ordered.sort(key=lambda n:norm(n)!=norm(preferred))
    for name in ordered:
        rows=sheets[name]
        for i,row in enumerate(rows[:25]):
            headers={norm(v):j for j,v in enumerate(row) if v is not None}
            if all(any(norm(alias) in headers for alias in alternatives) for alternatives in required):
                repeated=Counter(norm(v) for v in row if v is not None)
                important={'STOK KODU','HAM KODU','TARIH','YIL','AY','MIKTAR','TOPLAM MIKTAR','FIS NO','BIRIM','HATA KODU','EKALAN1','GRAMAJ','CEVRIM PAY-1','OLCU BR-1','OLCU BR-2'}|{norm(m) for m in MONTHS}
                duplicate=[k for k,n in repeated.items() if n>1 and (k in important or k.startswith('BIRIM MALZEME MALIYET'))]
                if duplicate:raise InputError(f'{name}!{i+1}: aynı başlık birden fazla sütunda var: '+', '.join(duplicate))
                return name, rows, i, headers
    raise TableNotFound('Beklenen sütunlar bulunamadı: '+', '.join(a[0] for a in required))

def find_unique_table(sheets,required):
    tables=[]
    for name,rows in sheets.items():
        try:tables.append(find_table({name:rows},required))
        except TableNotFound:pass
    if not tables:raise TableNotFound('Beklenen sütunlar bulunamadı: '+', '.join(a[0] for a in required))
    if len(tables)>1:raise InputError('Birden fazla kaynak tablosu bulundu: '+', '.join(t[0] for t in tables)+'. Kaynak hareketlerini tek tabloda birleştirin.')
    return tables[0]

def get(row,headers,*aliases):
    for alias in aliases:
        i=headers.get(norm(alias))
        if i is not None: return row[i] if i<len(row) else None
    return None

def period_hint(sheets):
    text=' '.join(str(v) for rows in sheets.values() for row in rows[:3] for v in row if v is not None)
    years={int(v) for v in re.findall(r'(?:YIL|Yıl|Yil)\s*[:=]?\s*(20\d\d)',text,re.I)}
    months={int(v) for v in re.findall(r'\bAY\s*[:=]?\s*(\d{1,2})(?!\d)',norm(text))}
    if len(years)>1 or len(months)>1:raise InputError('Dosya başlıklarında birden fazla yıl/ay var; dönemler ayrı ve açık şekilde verilmelidir.')
    if any(not 1<=m<=12 for m in months):raise InputError('Dosya başlığındaki ay geçersiz.')
    return (next(iter(years),None),next(iter(months),None))

def record_period(row,h,hint,label):
    raw=get(row,h,'Tarih');dt=parse_date(raw)
    if raw not in (None,'') and not dt:raise InputError(f'{label}: tarih okunamadı; satır başka bir aya atanamaz.')
    if dt:return dt.year,dt.month,dt
    values=[]
    for name,fallback,lower,upper in [('Yıl',hint[0],2000,2100),('Ay',hint[1],1,12)]:
        v=numeric(get(row,h,name),f'{label} {name}')
        if v is None:v=fallback
        if v is not None and (v!=int(v) or not lower<=v<=upper):raise InputError(f'{label}: {name} geçersiz.')
        values.append(int(v) if v is not None else 0)
    return *values,None

def scrap_records(name,rows,start,h,hint,role):
    records=[]
    for n,row in enumerate(rows[start+1:],start+2):
        stock=code(get(row,h,'Stok Kodu'))
        if len(stock)<6:continue
        year,month,dt=record_period(row,h,hint,f'{name}!{n}')
        q=numeric(get(row,h,'Miktar','Toplam Miktar'),f'{name}!{n} miktar')
        if q is None:raise InputError(f'{name}!{n}: hurda miktarı boş.')
        if q<0:raise InputError(f'{name}!{n}: negatif hurda miktarı var; iade/düzeltme kaydı önce ayrılmalı.')
        if role=='scrap' and (not year or not month):raise InputError(f'{name}!{n}: hurda tarihi eksik/geçersiz.')
        records.append(dict(code=stock,name=str(get(row,h,'Stok Adı','Stok İsmi') or ''),year=year,month=month,
            date=dt.isoformat() if dt else None,source_row=n,sheet=name,qty=q,
            error=code(get(row,h,'Hata Kodu','EKALAN1')) or None,
            voucher=voucher(get(row,h,'Fiş No')),voucher_no=code(get(row,h,'Fiş No')) or None,
            unit=norm(get(row,h,'Birim')) or None,source=role))
    return records

def combined_baseline(sheets,summary_name,summary,hint):
    """An embedded transaction table supplies the controls for a combined HURDA table."""
    if norm(summary_name)!='HURDA':return None
    other={name:rows for name,rows in sheets.items() if name!=summary_name}
    if not other:return None
    try:table=find_unique_table(other,[['Stok Kodu'],['Miktar','Toplam Miktar'],['Fiş No']])
    except TableNotFound:return None
    raw=scrap_records(*table,hint,'scrap')
    # Every original stock/period/fiş must remain represented in the combined table.
    totals=defaultdict(float)
    for r in summary:
        if r['code'][:3] in GROUPS:totals[(r['code'],r['year'],r['month'])]+=r['qty']
    candidates={}
    explicit_vouchers={r['voucher'] for r in summary if r['voucher']}
    for v in ('Y','Ç','HRD'):
        if explicit_vouchers and v not in explicit_vouchers:continue
        baseline=[r for r in raw if r['code'][:3] in GROUPS and r['voucher']==v and r['qty']>0]
        controls=defaultdict(float)
        for r in baseline:controls[(r['code'],r['year'],r['month'])]+=r['qty']
        if controls and all(k in totals and totals[k]+1e-7>=q for k,q in controls.items()):candidates[v]=baseline
    if len(candidates)>1:raise InputError('HURDA birleşik tablosunun fiş grubu belirsiz; dosyada birleşik sonucun fiş grubunu belirtin.')
    if not candidates:
        if any(r['code'][:3] in GROUPS and r['qty']>0 for r in raw):
            raise InputError('HURDA tablosu dosyanın kaynak hareketleriyle uyuşmuyor; birleşik sonuç veya yalnız ek miktar içeren ayrı bir dosya yükleyin.')
        return None
    v,baseline=next(iter(candidates.items()))
    if any(r['voucher'] and r['voucher']!=v for r in summary):raise InputError('Birleşik HURDA tablosunun fiş grubu kaynak hareketleriyle uyuşmuyor.')
    for r in summary:r['voucher']=v
    return baseline

def parse_file(path,role,filename=None):
    sheets=read_sheets(path,filename)
    warnings=[]; records=[]; hint=period_hint(sheets);baseline=None
    if role=='consumption':
        name,rows,start,h=find_unique_table(sheets,[['Ham Kodu','Stok Kodu'],['Yıl']])
        for n,row in enumerate(rows[start+1:],start+2):
            stock=code(get(row,h,'Ham Kodu','Stok Kodu'))
            if len(stock)<6: continue
            year=numeric(get(row,h,'Yıl'),f'{name}!{n} yıl')
            if not year or year!=int(year) or not 2000<=year<=2100: raise InputError(f'{name}!{n}: yıl eksik/geçersiz.')
            for month,label in enumerate(MONTHS,1):
                v=numeric(get(row,h,label),f'{stock} {label} tüketim')
                if v is not None: records.append({'code':stock,'year':int(year),'month':month,'qty':v})
    elif role=='weights':
        name,rows,start,h=find_unique_table(sheets,[['Stok Kodu'],['Çevrim Pay-1','Gramaj']])
        for n,row in enumerate(rows[start+1:],start+2):
            stock=code(get(row,h,'Stok Kodu'))
            if len(stock)<6: continue
            factor=numeric(get(row,h,'Çevrim Pay-1','Gramaj'),f'{stock} gramaj')
            origin=norm(get(row,h,'Ölçü Br-1') or 'MT'); target=norm(get(row,h,'Ölçü Br-2') or 'GR')
            if stock.startswith('202') and (origin not in ['MT','AD'] or target not in ['GR','KG']):
                warnings.append(f'{stock}: desteklenmeyen gramaj birimi {origin} → {target}.');continue
            if factor is not None: records.append({'code':stock,'grams':factor*(1000 if target=='KG' else 1),'unit':origin})
    else:
        required=[['Stok Kodu'],['Birim Malzeme Maliyet (TL)','Birim Malzeme Maliyeti (TL)']] if role=='costs' else [['Stok Kodu'],['Miktar','Toplam Miktar']]
        summary=role=='scrap' and any(norm(n)=='HURDA' for n in sheets)
        if summary:table=find_unique_table({n:r for n,r in sheets.items() if norm(n)=='HURDA'},required)
        elif role=='fragmented':table=find_table(sheets,required,'HURDA')
        else:table=find_unique_table(sheets,required)
        name,rows,start,h=table
        if role=='fragmented' and norm(name)!='HURDA':name,rows,start,h=find_unique_table(sheets,required)
        if role in ('scrap','fragmented'):
            records=scrap_records(name,rows,start,h,hint,role)
            if summary and norm(name)=='HURDA':
                if any(not r['voucher'] for r in records):
                    controls=combined_baseline(sheets,name,records,hint)
                    if not controls:raise InputError('Ana HURDA tablosunda fiş grubu yok; kaynak fiş bilgilerini dosyada belirtin.')
                warnings.append('Ana kaynak HURDA sayfasıdır. Diğer sayfaların miktarları ayrıca eklenmez.')
            if role=='fragmented':
                baseline=combined_baseline(sheets,name,records,hint)
                if baseline:warnings.append('Birleşik HURDA tablosu tespit edildi. Ana dosyayla kaynak miktarları doğrulanarak çift sayım önlenecek.')
        else:
            records=parse_cost_records(name,rows,start,h,hint)
    if not records: raise InputError('Dosyada kullanılabilir veri satırı bulunamadı.')
    if role=='fragmented' and any(not r['voucher'] for r in records):
        warnings.append('Parçalanmış tabloda fiş grubu bulunmuyor. Aşağıdaki Parçalanmış fiş grubu seçimi kullanılacak.')
    if role=='fragmented' and any(not r['year'] or not r['month'] for r in records):
        warnings.append('Parçalanmış tabloda bazı satırların yıl/ayı yok; arayüzdeki parçalanmış veri dönemi kullanılacak.')
    result={'role':role,'filename':Path(filename or ('input.xlsx' if hasattr(path,'read') else path)).name,'records':records,'warnings':warnings,'hint':list(hint)}
    if baseline:result['combined_baseline']=baseline
    return result

def parse_cost_records(name,rows,start,h,hint):
    records=[]
    for n,row in enumerate(rows[start+1:],start+2):
        stock=code(get(row,h,'Stok Kodu'))
        if len(stock)<6:continue
        year,month,dt=record_period(row,h,hint,f'{name}!{n}')
        if not dt:raise InputError(f'{name}!{n}: maliyet tarihi eksik/geçersiz; son tarihli değer seçilemiyor.')
        records.append(dict(code=stock,name=str(get(row,h,'Stok Adı','Stok İsmi') or ''),year=year,month=month,
            date=dt.isoformat(),source_row=n,sheet=name,
            tl=numeric(get(row,h,'Birim Malzeme Maliyet (TL)','Birim Malzeme Maliyeti (TL)'),f'{stock} TL maliyet'),
            eur=numeric(get(row,h,'Birim Malzeme Maliyet (EUR)','Birim Malzeme Maliyeti (EUR)'),f'{stock} EUR maliyet')))
    return records

def sort_error(value):
    s=str(value or '')
    return (0,int(s)) if s.isdigit() else (1,s)

def shifted(year,month,delta):
    v=year*12+month-1+delta
    return v//12,v%12+1

def resolve_mode(rows,history,year,month):
    counts=Counter(r['error'] for r in rows if r['error'])
    if not counts:return None,'Bu ayın hata kodları tamamen boş.'
    candidates={k for k,v in counts.items() if v==max(counts.values())}
    if len(candidates)==1:return next(iter(candidates)),'Ay içindeki kayıt sayısına göre mod.'
    prior={shifted(year,month,-i) for i in range(1,6)}
    old=Counter(r['error'] for r in history if (r['year'],r['month']) in prior and r['error'] in candidates)
    if old:
        winners=[k for k in candidates if old[k]==max(old.values())]
        if len(winners)==1:return winners[0],'Eşitlik, önceki 5 ayın kayıt sayısıyla çözüldü.'
    return None,'Mod eşitliği önceki 5 ayda da çözülemedi; boş kodlar belirsiz bırakıldı.'

def merge_scrap(inputs,options,year,months,groups,vouchers):
    main=[dict(r) for r in inputs['scrap']['records']]
    fragment=inputs.get('fragmented')
    if not fragment:return main,[]
    fy=integer(options.get('fragment_year',year),'Parçalanmış yıl',2000,2100);fm=integer(options.get('fragment_month',months[-1]),'Parçalanmış ay',1,12);fv=options.get('fragment_voucher','Y')
    if not 2000<=fy<=2100 or not 1<=fm<=12 or fv not in ('Y','Ç','HRD'):raise InputError('Parçalanmış veri dönemi ve fiş grubu geçersiz.')
    extra=[]
    for original in fragment['records']:
        r=dict(original);r['year']=r['year'] or fy;r['month']=r['month'] or fm;r['voucher']=r['voucher'] or fv
        extra.append(r)
    baseline=fragment.get('combined_baseline')
    if not baseline:return main+extra,[]
    def key(r):return r['code'],r['year'],r['month'],r['voucher']
    controls=defaultdict(float);actual=defaultdict(float);control_docs=defaultdict(float);actual_docs=defaultdict(float)
    for r in baseline:controls[key(r)]+=r['qty']
    for r in main:actual[key(r)]+=r['qty']
    for r in baseline:
        if r.get('voucher_no'):control_docs[(key(r),r['voucher_no'])]+=r['qty']
    for r in main:
        if key(r) in controls:actual_docs[(key(r),r.get('voucher_no'))]+=r['qty']
    relevant={(year,m) for m in months}
    relevant.update(shifted(year,m,-i) for m in months for i in range(1,6))
    for k,expected in controls.items():
        if k[0][:3] in groups and k[3] in vouchers and (k[1],k[2]) in relevant:
            if k not in actual or not math.isclose(actual[k],expected,rel_tol=1e-9,abs_tol=1e-7):
                raise InputError(f'{k[0]} ({k[1]} / {k[2]}, {k[3]}): birleşik dosyanın ana hurda miktarı {expected:g}; ana dosyada {actual.get(k,0):g}. Kaynaklar eşleşmediği için çift sayım kontrolü tamamlanamadı.')
            docs={d:q for (kk,d),q in control_docs.items() if kk==k}
            if docs:
                found={d:q for (kk,d),q in actual_docs.items() if kk==k and q!=0}
                if set(found)!=set(docs) or any(not math.isclose(found[d],q,rel_tol=1e-9,abs_tol=1e-7) for d,q in docs.items()):
                    raise InputError(f'{k[0]}: birleşik dosyanın fiş numaraları ve miktarları ana dosyayla eşleşmiyor.')
            units={r['unit'] for r in main if key(r)==k and r['unit']}
            expected_unit=next((r.get('unit','MT') for r in inputs['weights']['records'] if r['code']==k[0]),'MT') if k[0].startswith('202') else 'AD'
            baseline_units={r['unit'] for r in baseline if key(r)==k and r['unit']}
            if any(u!=expected_unit for u in units|baseline_units):raise InputError(f'{k[0]}: birleşik dosyanın kaynak hareketlerinde birimler uyuşmuyor.')
    # Replace verified originals with their combined result. Never subtract one
    # error-code bucket from a different bucket, or deduplicate legitimate extra rows.
    removed=defaultdict(float);replacement=defaultdict(float)
    for k,q in controls.items():removed[(k[0][:3],k[1],k[2],k[3])]+=q
    for r in extra:replacement[(r['code'][:3],r['year'],r['month'],r['voucher'])]+=r['qty']
    checks=[]
    for k,q in sorted(replacement.items()):
        group,y,m,v=k
        if group in groups and y==year and m in months and v in vouchers:
            before=removed.get(k,0)
            checks.append(dict(group=group,year=y,month=m,voucher=v,original_qty=before,
                combined_qty=q,net_added_qty=q-before,source=fragment['filename']))
    return [r for r in main if key(r) not in controls]+extra,checks

def analyze(inputs,options):
    if not isinstance(options,dict):raise InputError('Analiz seçimleri bir nesne olarak gönderilmelidir.')
    for field in ('no_fragment','skip_missing'):
        if field in options and not isinstance(options[field],bool):raise InputError(f'{field}: evet/hayır seçimi geçersiz.')
    if options.get('main_only') is True:inputs={role:value for role,value in inputs.items() if role!='fragmented'}
    if options.get('no_fragment') is True:inputs={**inputs,'fragmented':None}
    year=integer(options['year'],'Yıl',2000,2100);months=sorted(set(integer(m,'Ay',1,12) for m in options['months']))
    groups=list(dict.fromkeys(options['groups']));vouchers=list(dict.fromkeys(options['vouchers']))
    if not 2000<=year<=2100 or not months or any(m<1 or m>12 for m in months): raise InputError('Geçerli yıl ve ay seçin.')
    if not groups or not set(groups)<=set(GROUPS) or not vouchers or not set(vouchers)<={'Y','Ç','HRD'}: raise InputError('Fiş ve stok grubu seçin.')
    all_rows,reconciliation=merge_scrap(inputs,options,year,months,groups,vouchers)
    scoped=[r for r in all_rows if r['code'][:3] in groups and r['voucher'] in vouchers]
    current=[r for r in scoped if r['year']==year and r['month'] in months and r['qty']!=0]
    active_keys={(r['code'],year,r['month']) for r in current}
    costs=defaultdict(list)
    for r in inputs['costs']['records']:costs[(r['code'],r['year'],r['month'])].append(r)
    rates={};conflicts=[]
    for key,rr in costs.items():
        latest=max(r['date'] for r in rr);same=[r for r in rr if r['date']==latest]
        choices={(r['tl'],r['eur']) for r in same}
        if len(choices)>1 and key in active_keys:
            conflicts.append({'code':key[0],'month':key[2],'date':latest})
        rates[key]=same[-1]
    if conflicts:raise InputError('Son tarihte farklı maliyetler var. Maliyet dosyasını düzeltin: '+', '.join(f"{r['code']} ({r['month']}. ay)" for r in conflicts))
    consumption=defaultdict(float);cons_count=Counter()
    for r in inputs['consumption']['records']:
        key=(r['code'],r['year'],r['month']);consumption[key]+=r['qty'];cons_count[key]+=1
    weights={};weight_units={}
    needed_weights={r['code'] for r in current if r['code'].startswith('202')}
    for r in inputs['weights']['records']:
        if r['code'] not in needed_weights:continue
        if r['code'] in weights and weights[r['code']]!=r['grams']:raise InputError(f"{r['code']}: çelişen gramajlar var.")
        if r['code'] in weight_units and weight_units[r['code']]!=r.get('unit','MT'):raise InputError(f"{r['code']}: çelişen gramaj birimleri var.")
        weights[r['code']]=r['grams']
        weight_units[r['code']]=r.get('unit','MT')
    excluded=[];eligible=[];missing={}
    for r in current:
        key=(r['code'],year,r['month']);rate=rates.get(key)
        if not rate:
            prior=[c for c in inputs['costs']['records'] if c['code']==r['code'] and (c['year']<year or (c['year']==year and c['month']<=r['month'])) and c.get('date')]
            if prior:
                latest_date=max(c['date'] for c in prior)
                same=[c for c in prior if c['date']==latest_date]
                choices={(c['tl'],c['eur']) for c in same}
                if len(choices)>1 and key in active_keys:
                    conflicts.append({'code':key[0],'month':key[2],'date':latest_date})
                rate=same[-1]
                rates[key]=rate
        reasons=[]
        if not rate or rate['tl'] is None or rate['eur'] is None:reasons.append('TL/EUR birim malzeme maliyeti eksik')
        if key not in consumption:reasons.append('Tüketim eksik')
        elif consumption[key]<0:raise InputError(f"{r['code']}: seçilen ayın tüketimi negatif; kaynak veriyi kontrol edin.")
        if r['code'].startswith('202') and (r['code'] not in weights or weights[r['code']]<=0):reasons.append('Gramaj eksik veya geçersiz')
        expected=weight_units.get(r['code'],'MT') if r['code'].startswith('202') else 'AD'
        if r['unit'] and r['unit']!=expected and (not r['code'].startswith('202') or r['code'] in weights):raise InputError(f"{r['code']}: hurda birimi {r['unit']}; gramaj/birim tanımında {expected} bekleniyor.")
        r['unit']=r['unit'] or expected
        if reasons:
            missing[key]={'code':r['code'],'month':r['month'],'reason':'; '.join(reasons)}
            excluded.append(r)
        else:eligible.append(r)
    issues=list(missing.values())
    if issues and not options.get('skip_missing'):
        return {'requires_confirmation':True,'issues':issues,'message':'Bu stokların verileri eksik. Bu stokları hurda, maliyet ve PPM hesabından çıkararak devam etmek istiyor musunuz?'}
    warnings=[w for inp in inputs.values() if inp for w in inp['warnings']]
    eligible_keys={(r['code'],year,r['month']) for r in eligible}
    if any(n>1 for key,n in cons_count.items() if key in eligible_keys): warnings.append('Aynı stok/yıl için birden fazla tüketim satırı toplandı.')
    processed=[];modes=[]
    for month in months:
        rows=[r for r in eligible if r['month']==month]
        mode,note=resolve_mode(rows,[r for r in scoped if r['qty']>0],year,month)
        if rows:
            modes.append({'month':month,'mode':mode,'note':note,'assigned_count':sum(not r['error'] for r in rows)})
        for r in rows:
            rr=dict(r);rr['original_error']=rr['error'];rr['assigned']=not bool(rr['error']);rr['error']=rr['error'] or mode or 'Belirsiz';processed.append(rr)
    coverage=[]
    for month in months:
        source_rows=sum(r['year']==year and r['month']==month for r in all_rows)
        selected_rows=sum(r['month']==month for r in current)
        eligible_rows=sum(r['month']==month for r in processed)
        status='no_source' if not source_rows else 'no_scope' if not selected_rows else 'all_excluded' if not eligible_rows else 'partial' if eligible_rows<selected_rows else 'ok'
        coverage.append(dict(month=month,source_rows=source_rows,selected_rows=selected_rows,eligible_rows=eligible_rows,excluded_rows=selected_rows-eligible_rows,status=status))
    results=[]
    for group in groups:
        cable_in_metres=group!='202' or all(r['unit']=='MT' for r in processed if r['code'].startswith(group))
        stocks={};pivot=defaultdict(lambda:defaultdict(float));pareto=defaultdict(float);monthly=[]
        for month in months:
            rows=[r for r in processed if r['month']==month and r['code'].startswith(group)]
            quantities=defaultdict(float)
            for r in rows:
                quantities[r['code']]+=r['qty']
                measured=r['qty']*(weights[r['code']]/1000 if group=='202' and not cable_in_metres else 1)
                pivot[r['code']][r['error']]+=measured;pareto[r['error']]+=measured
            month_stocks=[]
            for stock,qty in sorted(quantities.items()):
                key=(stock,year,month);rate=rates[key];grams=weights.get(stock) if group=='202' else None
                source=next(r for r in rows if r['code']==stock)
                item={'code':stock,'name':source['name'],'month':month,'unit':source['unit'],'scrap':qty,'consumption':consumption[key],'grams':grams,'scrap_kg':qty*grams/1000 if grams is not None else None,'consumption_kg':consumption[key]*grams/1000 if grams is not None else None,'unit_tl':rate['tl'],'unit_eur':rate['eur'],'cost_tl':qty*rate['tl'],'cost_eur':qty*rate['eur'],'cost_date':rate['date']}
                month_stocks.append(item);stocks[(month,stock)]=item
            scrap=sum(i['scrap_kg'] if group=='202' else i['scrap'] for i in month_stocks)
            consumed=sum(i['consumption_kg'] if group=='202' else i['consumption'] for i in month_stocks)
            source_rows=sum(r['month']==month and r['code'].startswith(group) for r in current)
            file_rows=next(c['source_rows'] for c in coverage if c['month']==month)
            status='no_source' if not file_rows else 'no_scope' if not source_rows else 'all_excluded' if not rows else 'partial' if len(rows)<source_rows else 'ok'
            monthly.append({'month':month,'label':MONTHS[month-1],'scrap':scrap,'consumption':consumed,'ppm':scrap/consumed*1e6 if consumed>0 else None,'cost_tl':sum(i['cost_tl'] for i in month_stocks),'cost_eur':sum(i['cost_eur'] for i in month_stocks),'stock_count':len(month_stocks),'source_rows':source_rows,'eligible_rows':len(rows),'data_status':status})
        total_scrap=sum(i['scrap'] for i in monthly);total_consumption=sum(i['consumption'] for i in monthly)
        if stocks and not total_consumption:warnings.append(f'{GROUPS[group]}: tüketim toplamı sıfır; miktar ve maliyet hesaplandı, PPM hesaplanamaz.')
        error_codes=sorted(pareto,key=sort_error);cum=0;pareto_rows=[];pq=sum(pareto.values())
        for error,qty in sorted(pareto.items(),key=lambda x:(-x[1],sort_error(x[0]))):
            share=qty/pq if pq else 0;cum+=share
            pareto_rows.append({'error':error,'description':ERRORS.get(int(error),'Tanımsız hata kodu') if str(error).isdigit() else error,'qty':qty,'share':share,'cumulative':cum})
        statuses={m['data_status'] for m in monthly}
        status=('partial' if statuses-{'ok','no_scope'} else 'ok') if stocks else 'all_excluded' if 'all_excluded' in statuses else 'no_scope' if statuses-{'no_source'} else 'no_source'
        results.append({'group':group,'name':GROUPS[group],'unit':'kg' if group=='202' else 'adet','qty_unit':('metre' if cable_in_metres else 'kaynak birimi') if group=='202' else 'adet','pareto_unit':('metre' if cable_in_metres else 'kg') if group=='202' else 'adet','stocks':list(stocks.values()),'monthly':monthly,'pivot_codes':error_codes,'pivot':[{'code':stock,'errors':dict(errors),'total':sum(errors.values())} for stock,errors in sorted(pivot.items())],'pareto':pareto_rows,'scrap':total_scrap,'consumption':total_consumption,'ppm':total_scrap/total_consumption*1e6 if total_consumption else None,'cost_tl':sum(i['cost_tl'] for i in monthly),'cost_eur':sum(i['cost_eur'] for i in monthly),'data_status':status,'target':numeric(options.get('targets',{}).get(group))})
    for g in results:g['pivot_unit']=g['pareto_unit']
    return {'requires_confirmation':False,'year':year,'months':months,'vouchers':vouchers,'groups':results,'modes':modes,'issues':issues,'warnings':warnings,'coverage':coverage,'reconciliation':reconciliation,'excluded_rows':len(excluded),'detail':processed,'sources':{role:inp['filename'] for role,inp in inputs.items() if inp},'created':datetime.now().isoformat(timespec='microseconds')}
