import io,zipfile
from copy import deepcopy
import openpyxl
from analysis_engine import analyze
from test_analysis import fixture,opts
from web_excel import export_zip

def analysis_book(result,prefix='KABLO'):
    with zipfile.ZipFile(io.BytesIO(export_zip(result))) as z:
        name=next(n for n in z.namelist() if n.startswith(prefix) and 'PARETO' not in n)
        return openpyxl.load_workbook(io.BytesIO(z.read(name)),data_only=True)

def test_unrepresented_month_does_not_export_false_zero_totals():
    o=opts();o['months']=[11]
    r=analyze(fixture(),o)
    assert r['coverage'][0]['status']=='no_source'
    assert all(g['data_status']=='no_source' for g in r['groups'])
    w=analysis_book(r)
    for row in (2,3):
        for col in ('B','C','D','F','G'):assert w['GRAFİK VERİSİ'][f'{col}{row}'].value is None
    assert w['Dönem Kontrolü']['E2'].value=='Kaynak dosyada bu dönem yok'
    w.close()

def test_month_present_but_selected_voucher_absent_is_distinct():
    o=opts();o['vouchers']=['Ç']
    r=analyze(fixture(),o)
    assert r['coverage'][0]['source_rows']>0
    assert r['coverage'][0]['status']=='no_scope'
    assert all(g['data_status']=='no_scope' for g in r['groups'])

def test_excluding_every_stock_is_not_reported_as_zero_scrap():
    f=fixture();f['costs']['records']=[r for r in f['costs']['records'] if not r['code'].startswith('202')]
    o=opts(skip_missing=True);o['groups']=['202']
    r=analyze(f,o)
    assert r['coverage'][0]['status']=='all_excluded'
    assert r['groups'][0]['data_status']=='all_excluded'
    assert r['excluded_rows']==3
    w=analysis_book(r)
    assert w['GRAFİK VERİSİ']['C2'].value is None and w['GRAFİK VERİSİ']['F2'].value is None
    assert w['GRAFİK VERİSİ']['H2'].value=='Tüm hurda kayıtları eksik veri nedeniyle çıkarıldı'
    w.close()

def test_zero_material_cost_and_zero_consumption_remain_real_zero_values():
    f=fixture()
    for r in f['costs']['records']:r['tl']=0;r['eur']=0
    for r in f['consumption']['records']:r['qty']=0
    r=analyze(f,opts());g=r['groups'][0]
    assert g['data_status']=='ok' and g['scrap']==.1 and g['consumption']==0 and g['ppm'] is None
    w=analysis_book(r)
    assert w['GRAFİK VERİSİ']['B2'].value==0 and w['GRAFİK VERİSİ']['C2'].value==.1
    assert w['GRAFİK VERİSİ']['F2'].value==0 and w['GRAFİK VERİSİ']['G2'].value==0
    w.close()

def test_partial_month_range_keeps_known_totals_and_marks_unknown_period():
    o=opts();o['months']=[9,11]
    r=analyze(fixture(),o);g=r['groups'][0]
    assert g['data_status']=='partial'
    assert [m['data_status'] for m in g['monthly']]==['ok','no_source']
    assert g['ppm']==25000 and g['cost_tl']==50
    w=analysis_book(r)
    assert w['GRAFİK VERİSİ']['C3'].value is None
    assert w['GRAFİK VERİSİ']['F4'].value==50
    assert w['Dönem Kontrolü']['E3'].value=='Kaynak dosyada bu dönem yok'
    w.close()

def test_no_fragment_option_is_honored_even_if_uploaded_state_contains_file():
    f=fixture();before=deepcopy(f)
    o=opts();o['no_fragment']=True
    r=analyze(f,o)
    assert r['groups'][0]['stocks'][0]['scrap']==20
    assert r['groups'][0]['ppm']==20000
    assert not any(d['source']=='fragmented' for d in r['detail'])
    assert 'fragmented' not in r['sources']
    assert f==before
