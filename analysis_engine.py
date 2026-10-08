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
ERRORS = dict(enumerate(['Terminal Çakım Hatası','IDC Çakım Hatası','Kablo Boyu Hatalı','Taşıma Kaynaklı','Insertion Hatası','Potalama Kaynaklı','Tedarikçi Kaynaklı','Makara Sonu Firesi','Kablo Uç Açımı Hatalı','Yanlış Malzeme Kullanımı','Bıçak İzi','Ayar Firesi','İşçilik Firesi','Sipariş Fazlası Olarak Gelen Kablolar','Revizyon','Ezik / Kırık / Deformeli','Proje Kaynaklı','Operatör Hatası','Yanlış Onay','Elektriksel Test','Kilo Hatası','IDC Kesme Hurdası','Makina Kabul','Enjeksiyon','Fire','Diğer'],1))

class InputError(ValueError): pass

def norm(value):
    value = str(value or '').replace('ı','i').replace('İ','I')
    return ''.join(c for c in unicodedata.normalize('NFKD',value) if not unicodedata.combining(c)).upper().strip()

def code(value):
    if value is None: return ''
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
                    if cell.ctype==xlrd.XL_CELL_ERROR: row.append(None)
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
                return name, rows, i, headers
    raise InputError('Beklenen sütunlar bulunamadı: '+', '.join(a[0] for a in required))

def get(row,headers,*aliases):
    for alias in aliases:
        i=headers.get(norm(alias))
        if i is not None: return row[i] if i<len(row) else None
    return None

def period_hint(sheets):
    text=' '.join(str(v) for rows in sheets.values() for row in rows[:3] for v in row if v is not None)
    year=re.search(r'(?:YIL|Yıl|Yil)\s*[:=]?\s*(20\d\d)',text,re.I)
    month=re.search(r'AY\s*[:=]?\s*(\d{1,2})(?!\d)',norm(text))
    return (int(year.group(1)) if year else None, int(month.group(1)) if month else None)

def parse_file(path,role,filename=None):
    sheets=read_sheets(path,filename)
    warnings=[]; records=[]; hint=period_hint(sheets)
    if role=='consumption':
        name,rows,start,h=find_table(sheets,[['Ham Kodu','Stok Kodu'],['Yıl']])
        for n,row in enumerate(rows[start+1:],start+2):
            stock=code(get(row,h,'Ham Kodu','Stok Kodu'))
            if len(stock)<6: continue
            year=numeric(get(row,h,'Yıl'),f'{name}!{n} yıl')
            if not year: raise InputError(f'{name}!{n}: yıl eksik.')
            for month,label in enumerate(MONTHS,1):
                v=numeric(get(row,h,label),f'{stock} {label} tüketim')
                if v is not None: records.append({'code':stock,'year':int(year),'month':month,'qty':v})
    elif role=='weights':
        name,rows,start,h=find_table(sheets,[['Stok Kodu'],['Çevrim Pay-1','Gramaj']])
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
        name,rows,start,h=find_table(sheets,required, 'HURDA' if role=='fragmented' else None)
        for n,row in enumerate(rows[start+1:],start+2):
            stock=code(get(row,h,'Stok Kodu'))
            if len(stock)<6:continue
            dt=parse_date(get(row,h,'Tarih'))
            y=get(row,h,'Yıl'); m=get(row,h,'Ay')
            year=dt.year if dt else int(numeric(y) or hint[0] or 0)
            month=dt.month if dt else int(numeric(m) or hint[1] or 0)
            common={'code':stock,'name':str(get(row,h,'Stok Adı','Stok İsmi') or ''),'year':year,'month':month,'date':dt.isoformat() if dt else None,'source_row':n,'sheet':name}
            if role=='costs':
                if not dt: raise InputError(f'{name}!{n}: maliyet tarihi eksik/geçersiz; son tarihli değer seçilemiyor.')
                common.update(tl=numeric(get(row,h,'Birim Malzeme Maliyet (TL)','Birim Malzeme Maliyeti (TL)'),f'{stock} TL maliyet'),eur=numeric(get(row,h,'Birim Malzeme Maliyet (EUR)','Birim Malzeme Maliyeti (EUR)'),f'{stock} EUR maliyet'))
            else:
                q=numeric(get(row,h,'Miktar','Toplam Miktar'),f'{name}!{n} miktar')
                if q is None:raise InputError(f'{name}!{n}: hurda miktarı boş.')
                if q<0:raise InputError(f'{name}!{n}: negatif hurda miktarı var; iade/düzeltme kaydı önce ayrılmalı.')
                err=code(get(row,h,'Hata Kodu','EKALAN1')) or None
                common.update(qty=q,error=err,voucher=voucher(get(row,h,'Fiş No')),unit=norm(get(row,h,'Birim')) or None,source=role)
                if role=='scrap' and (not year or not month): raise InputError(f'{name}!{n}: hurda tarihi eksik/geçersiz.')
            records.append(common)
    if not records: raise InputError('Dosyada kullanılabilir veri satırı bulunamadı.')
    if role=='fragmented' and any(not r['voucher'] for r in records):
        warnings.append('Parçalanmış tabloda fiş grubu bulunmuyor. Aşağıdaki Parçalanmış fiş grubu seçimi kullanılacak.')
    return {'role':role,'filename':Path(filename or ('input.xlsx' if hasattr(path,'read') else path)).name,'records':records,'warnings':warnings,'hint':list(hint)}

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

def analyze(inputs,options):
    year=int(options['year']);months=sorted(set(int(m) for m in options['months']))
    groups=list(dict.fromkeys(options['groups']));vouchers=list(dict.fromkeys(options['vouchers']))
    if not 2000<=year<=2100 or not months or any(m<1 or m>12 for m in months): raise InputError('Geçerli yıl ve ay seçin.')
    if not groups or not set(groups)<=set(GROUPS) or not vouchers or not set(vouchers)<={'Y','Ç','HRD'}: raise InputError('Fiş ve stok grubu seçin.')
    all_rows=[dict(r) for r in inputs['scrap']['records']]
    if inputs.get('fragmented'):
        fy=int(options.get('fragment_year') or year);fm=int(options.get('fragment_month') or months[-1]);fv=options.get('fragment_voucher','Y')
        for original in inputs['fragmented']['records']:
            r=dict(original);r['year']=r['year'] or fy;r['month']=r['month'] or fm;r['voucher']=r['voucher'] or fv
            all_rows.append(r)
    scoped=[r for r in all_rows if r['code'][:3] in groups and r['voucher'] in vouchers]
    current=[r for r in scoped if r['year']==year and r['month'] in months and r['qty']!=0]
    costs=defaultdict(list)
    for r in inputs['costs']['records']:costs[(r['code'],r['year'],r['month'])].append(r)
    rates={};conflicts=[]
    for key,rr in costs.items():
        latest=max(r['date'] for r in rr);same=[r for r in rr if r['date']==latest]
        choices={(r['tl'],r['eur']) for r in same}
        if len(choices)>1 and key in {(r['code'],year,r['month']) for r in current}:
            conflicts.append({'code':key[0],'month':key[2],'date':latest})
        rates[key]=same[-1]
    if conflicts:raise InputError('Son tarihte farklı maliyetler var. Maliyet dosyasını düzeltin: '+', '.join(f"{r['code']} ({r['month']}. ay)" for r in conflicts))
    consumption=defaultdict(float);cons_count=Counter()
    for r in inputs['consumption']['records']:
        key=(r['code'],r['year'],r['month']);consumption[key]+=r['qty'];cons_count[key]+=1
    weights={};weight_units={}
    for r in inputs['weights']['records']:
        if r['code'] in weights and weights[r['code']]!=r['grams']:raise InputError(f"{r['code']}: çelişen gramajlar var.")
        weights[r['code']]=r['grams']
        weight_units[r['code']]=r.get('unit','MT')
    excluded=[];eligible=[];missing={}
    for r in current:
        key=(r['code'],year,r['month']);rate=rates.get(key)
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
    if any(n>1 for key,n in cons_count.items() if key in {(r['code'],year,r['month']) for r in eligible}): warnings.append('Aynı stok/yıl için birden fazla tüketim satırı toplandı.')
    processed=[];modes=[]
    for month in months:
        rows=[r for r in eligible if r['month']==month]
        mode,note=resolve_mode(rows,scoped,year,month)
        if rows:
            modes.append({'month':month,'mode':mode,'note':note,'assigned_count':sum(not r['error'] for r in rows)})
        for r in rows:
            rr=dict(r);rr['original_error']=rr['error'];rr['assigned']=not bool(rr['error']);rr['error']=rr['error'] or mode or 'Belirsiz';processed.append(rr)
    results=[]
    for group in groups:
        cable_in_metres=group!='202' or all(r['unit']=='MT' for r in processed if r['code'].startswith(group))
        stocks={};pivot=defaultdict(lambda:defaultdict(float));pareto=defaultdict(float);monthly=[]
        for month in months:
            rows=[r for r in processed if r['month']==month and r['code'].startswith(group)]
            quantities=defaultdict(float)
            for r in rows:
                quantities[r['code']]+=r['qty'];pivot[r['code']][r['error']]+=r['qty'];pareto[r['error']]+=r['qty']*(weights[r['code']]/1000 if group=='202' and not cable_in_metres else 1)
            month_stocks=[]
            for stock,qty in sorted(quantities.items()):
                key=(stock,year,month);rate=rates[key];grams=weights.get(stock) if group=='202' else None
                source=next(r for r in rows if r['code']==stock)
                item={'code':stock,'name':source['name'],'month':month,'unit':source['unit'],'scrap':qty,'consumption':consumption[key],'grams':grams,'scrap_kg':qty*grams/1000 if grams is not None else None,'consumption_kg':consumption[key]*grams/1000 if grams is not None else None,'unit_tl':rate['tl'],'unit_eur':rate['eur'],'cost_tl':qty*rate['tl'],'cost_eur':qty*rate['eur'],'cost_date':rate['date']}
                month_stocks.append(item);stocks[(month,stock)]=item
            scrap=sum(i['scrap_kg'] if group=='202' else i['scrap'] for i in month_stocks)
            consumed=sum(i['consumption_kg'] if group=='202' else i['consumption'] for i in month_stocks)
            monthly.append({'month':month,'label':MONTHS[month-1],'scrap':scrap,'consumption':consumed,'ppm':scrap/consumed*1e6 if consumed>0 else None,'cost_tl':sum(i['cost_tl'] for i in month_stocks),'cost_eur':sum(i['cost_eur'] for i in month_stocks),'stock_count':len(month_stocks)})
        total_scrap=sum(i['scrap'] for i in monthly);total_consumption=sum(i['consumption'] for i in monthly)
        if stocks and not total_consumption:warnings.append(f'{GROUPS[group]}: tüketim toplamı sıfır; miktar ve maliyet hesaplandı, PPM hesaplanamaz.')
        error_codes=sorted(pareto,key=sort_error);cum=0;pareto_rows=[];pq=sum(pareto.values())
        for error,qty in sorted(pareto.items(),key=lambda x:(-x[1],sort_error(x[0]))):
            share=qty/pq if pq else 0;cum+=share
            pareto_rows.append({'error':error,'description':ERRORS.get(int(error),'Tanımsız hata kodu') if str(error).isdigit() else error,'qty':qty,'share':share,'cumulative':cum})
        results.append({'group':group,'name':GROUPS[group],'unit':'kg' if group=='202' else 'adet','qty_unit':('metre' if cable_in_metres else 'kaynak birimi') if group=='202' else 'adet','pareto_unit':('metre' if cable_in_metres else 'kg') if group=='202' else 'adet','stocks':list(stocks.values()),'monthly':monthly,'pivot_codes':error_codes,'pivot':[{'code':stock,'errors':dict(errors),'total':sum(errors.values())} for stock,errors in sorted(pivot.items())],'pareto':pareto_rows,'scrap':total_scrap,'consumption':total_consumption,'ppm':total_scrap/total_consumption*1e6 if total_consumption else None,'cost_tl':sum(i['cost_tl'] for i in monthly),'cost_eur':sum(i['cost_eur'] for i in monthly),'target':numeric(options.get('targets',{}).get(group))})
    return {'requires_confirmation':False,'year':year,'months':months,'vouchers':vouchers,'groups':results,'modes':modes,'issues':issues,'warnings':warnings,'excluded_rows':len(excluded),'detail':processed,'sources':{role:inp['filename'] for role,inp in inputs.items() if inp},'created':datetime.now().isoformat(timespec='seconds')}
