import threading
from copy import deepcopy
import pytest
import analysis_engine as engine
from analysis_engine import analyze,parse_file,InputError
from test_analysis import fixture,opts,row

@pytest.mark.parametrize('field,value',[('year',2026.5),('months',[9.5]),('months',[True])])
def test_fractional_or_boolean_periods_are_not_silently_changed(field,value):
    o=opts();o[field]=value
    with pytest.raises(InputError):analyze(fixture(),o)

@pytest.mark.parametrize('field',['no_fragment','skip_missing'])
def test_string_false_cannot_authorize_excluding_missing_data(field):
    o=opts();o[field]='false'
    with pytest.raises(InputError,match='evet/hayır'):analyze(fixture(),o)

def test_conflicting_weights_of_an_unselected_stock_do_not_block_valid_analysis():
    f=fixture();f['weights']['records'] += [dict(code='202999999',grams=2),dict(code='202999999',grams=3)]
    assert analyze(f,opts())['groups'][0]['ppm']==25000

def test_duplicate_month_headers_cannot_choose_the_last_consumption_column(monkeypatch):
    monkeypatch.setattr(engine,'read_sheets',lambda *a,**k:{'A':[['Ham Kodu','Yıl','Eylül','Eylül'],['206000100',2026,1000,2000]]})
    with pytest.raises(InputError,match='aynı başlık'):parse_file('bad.xlsx','consumption')

def test_excel_error_in_stock_identifier_cannot_silently_disappear(monkeypatch):
    monkeypatch.setattr(engine,'read_sheets',lambda *a,**k:{'A':[['Stok Kodu','Miktar','Tarih'],['#N/A',20,'2026-09-01']]})
    with pytest.raises(InputError,match='Excel'):parse_file('bad.xlsx','scrap')

def test_zero_quantity_history_cannot_change_assignment_for_positive_scrap():
    f=fixture();f['fragmented']=None
    f['scrap']['records']=[row(error='16'),row(error='18'),row(error=None),row(qty=0,month=8,error='18')]
    r=analyze(f,opts())
    assert r['modes'][0]['mode'] is None
    assert r['groups'][0]['pivot'][0]['errors']['Belirsiz']==10

def test_chart_keeps_unavailable_month_gap_and_explains_zero_denominator():
    from charts import figure
    f=fixture();o=opts();o['months']=[9,11]
    r=analyze(f,o);fig=figure(r,r['groups'][0],'ppm')
    assert list(fig.axes[0].lines[0].get_xdata())==['Eylül','Kasım']
    fig.clear()
    f['consumption']['records'][0]['qty']=0
    for c in f['consumption']['records']:c['qty']=0
    r=analyze(f,opts());fig=figure(r,r['groups'][0],'ppm')
    assert any('tüketim sıfır' in t.get_text() for t in fig.axes[0].texts)
    fig.clear()

@pytest.mark.parametrize('change',['upload','reset','new_analysis'])
def test_concurrent_change_cannot_publish_result_from_old_inputs(monkeypatch,change):
    import app as web
    a=web.app.test_client();a.get('/api/state')
    with a.session_transaction() as s:sid=s['id'];web.STATE[sid]['inputs']=fixture()
    other=web.app.test_client()
    with other.session_transaction() as s:s['id']=sid
    entered=threading.Event();resume=threading.Event();response=[]
    real=web.analyze
    def delayed(inputs,options):
        if threading.current_thread().name=='old-analysis':
            entered.set();assert resume.wait(10)
        return real(inputs,options)
    monkeypatch.setattr(web,'analyze',delayed)
    worker=threading.Thread(name='old-analysis',target=lambda:response.append(a.post('/api/analyze',json=opts())))
    worker.start();assert entered.wait(10)
    try:
        if change=='upload':
            monkeypatch.setattr(web,'parse_file',lambda *a,**k:deepcopy(fixture()['weights']))
            import io
            assert other.post('/api/upload/weights',data={'file':(io.BytesIO(b'test'),'test.xlsx')}).status_code==200
        elif change=='reset':assert other.post('/api/reset').status_code==200
        else:assert other.post('/api/analyze',json=opts()).status_code==200
    finally:resume.set();worker.join(10)
    assert not worker.is_alive()
    assert response[0].status_code==400
    assert 'değişti' in response[0].json['error']
    if change!='new_analysis':assert other.get('/api/state').json['has_result'] is False

def test_rejected_analysis_options_clear_previous_exportable_result():
    from app import app,STATE
    a=app.test_client();a.get('/api/state')
    with a.session_transaction() as s:STATE[s['id']]['inputs']=fixture()
    assert a.post('/api/analyze',json=opts()).status_code==200
    invalid=opts();invalid['months']=[]
    assert a.post('/api/analyze',json=invalid).status_code==400
    assert a.get('/api/state').json['has_result'] is False

def mixed_fixture():
    f=fixture();f['scrap']['records'].append(dict(row('202999999',2),unit='AD'))
    f['weights']['records'].append(dict(code='202999999',grams=10,unit='AD'))
    f['costs']['records'].append(dict(code='202999999',year=2026,month=9,date='2026-09-20',tl=3,eur=.06))
    f['consumption']['records'].append(dict(code='202999999',year=2026,month=9,qty=100))
    return f

def test_mixed_cable_units_pivot_and_pareto_share_common_kg_and_preserve_raw_qty():
    r=analyze(mixed_fixture(),opts());g=r['groups'][0]
    assert g['pivot_unit']=='kg'
    assert sum(p['total'] for p in g['pivot'])==pytest.approx(.12)
    assert sum(p['qty'] for p in g['pareto'])==pytest.approx(.12)
    assert sum(p['qty'] for p in r['detail'] if p['code'].startswith('202'))==27
    from web_excel import export_zip
    import io,zipfile,openpyxl
    with zipfile.ZipFile(io.BytesIO(export_zip(r))) as z:
        w=openpyxl.load_workbook(io.BytesIO(z.read('KABLO HURDA ANALİZİ.xlsx')),data_only=True)
        assert sum(row[4] for row in list(w['Hurda Detayı'].values)[1:])==pytest.approx(.12)
        assert sum(row[8] for row in list(w['Hurda Detayı'].values)[1:])==27
        assert w['Hata Kodu Pivot']['B5'].value==pytest.approx(.12)
        cache=w['Hata Kodu Pivot']._pivots[0].cache
        assert sum(row._fields[4].v for row in cache.records.r)==pytest.approx(.12)
        w.close()

def test_xls_numeric_error_is_preserved_and_not_added_as_missing_consumption(monkeypatch):
    import xlrd
    from types import SimpleNamespace
    rows=[['Ham Kodu','Yıl','Eylül'],['206000100',2026,7]]
    sheet=SimpleNamespace(name='A',nrows=2,ncols=3,cell=lambda r,c:SimpleNamespace(ctype=xlrd.XL_CELL_ERROR if (r,c)==(1,2) else xlrd.XL_CELL_TEXT,value=rows[r][c]))
    book=SimpleNamespace(sheets=lambda:[sheet],release_resources=lambda:None)
    monkeypatch.setattr(xlrd,'open_workbook',lambda **k:book)
    import io
    with pytest.raises(InputError,match='sayısal olmayan'):parse_file(io.BytesIO(b'test'),'consumption',filename='bad.xls')
