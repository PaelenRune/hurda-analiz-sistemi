from copy import deepcopy
import pytest
from analysis_engine import analyze, voucher, resolve_mode, InputError, read_sheets
from pathlib import Path

def row(stock='202100003-71',qty=10,month=9,error='16',voucher='Y',source='scrap'):
    return dict(code=stock,name='Test malzemesi',qty=qty,month=month,year=2026,error=error,voucher=voucher,source=source,unit='MT' if stock.startswith('202') else 'AD',date='2026-09-01T00:00:00',source_row=4,sheet='Sheet1')

def fixture():
    rows=[row(),row(error=None),row('206000100',20),row('204000100',30)]
    def inp(records):return dict(records=records,warnings=[],filename='test.xlsx',hint=[None,None])
    costs=[];cons=[]
    for stock in {r['code'] for r in rows}:
        costs.append(dict(code=stock,year=2026,month=9,date='2026-09-20T00:00:00',tl=2,eur=.04))
        cons.append(dict(code=stock,year=2026,month=9,qty=1000))
    return {'scrap':inp(rows),'consumption':inp(cons),'costs':inp(costs),'weights':inp([dict(code='202100003-71',grams=4)]),'fragmented':inp([row(qty=5,source='fragmented')])}

def opts(**kw):return dict(year=2026,months=[9],groups=['202','204','206'],vouchers=['Y'],**kw)

def test_additive_fragment_and_units():
    r=analyze(fixture(),opts());g=r['groups'][0]
    assert g['stocks'][0]['scrap']==25
    assert g['scrap']==.1
    assert g['consumption']==4
    assert g['ppm']==25000
    assert g['cost_tl']==50
    assert g['pivot'][0]['errors']=={'16':25}

def test_missing_cost_gate_and_denominator_exclusion():
    f=fixture();f['scrap']['records'].append(row('202100004-71',100))
    f['consumption']['records'].append(dict(code='202100004-71',year=2026,month=9,qty=900000))
    f['weights']['records'].append(dict(code='202100004-71',grams=4))
    assert analyze(f,opts())['requires_confirmation']
    r=analyze(f,opts(skip_missing=True))
    assert r['groups'][0]['ppm']==25000
    assert r['excluded_rows']==1
    assert r['issues'][0]['code']=='202100004-71'

def test_latest_cost_and_real_zero():
    f=fixture();f['costs']['records'].append(dict(code='202100003-71',year=2026,month=9,date='2026-09-30T00:00:00',tl=0,eur=0))
    assert analyze(f,opts())['groups'][0]['cost_tl']==0

def test_conflicting_latest_cost_does_not_pick_arbitrarily():
    f=fixture();f['costs']['records'].append(dict(code='202100003-71',year=2026,month=9,date='2026-09-20T00:00:00',tl=9,eur=.04))
    with pytest.raises(InputError,match='Son tarihte'):analyze(f,opts())

def test_mode_counts_records_not_quantity():
    rows=[row(qty=1000,error='18'),row(qty=1,error='16'),row(qty=1,error='16')]
    assert resolve_mode(rows,[],2026,9)[0]=='16'

def test_mode_five_month_window_across_year():
    current=[row(error='16'),row(error='18')]
    hist=[dict(row(error='18'),year=2025,month=12),dict(row(error='16'),year=2025,month=6)]
    assert resolve_mode(current,hist,2026,1)[0]=='18'
    assert resolve_mode(current,[],2026,1)[0] is None

def test_all_errors_empty_stays_unknown():
    mode,note=resolve_mode([row(error=None)],[],2026,9)
    assert mode is None and 'tamamen boş' in note

def test_exact_voucher_prefixes():
    assert [voucher(v) for v in ['Y0001','Ç0001','HRD0001','H0001','C0001']]==['Y','Ç','HRD',None,None]

def test_multiple_months_ppm_is_weighted_ratio():
    f=fixture();f['scrap']['records'].append(row(qty=10,month=8));f['costs']['records'].append(dict(code='202100003-71',year=2026,month=8,date='2026-08-30T00:00:00',tl=2,eur=.04));f['consumption']['records'].append(dict(code='202100003-71',year=2026,month=8,qty=100))
    o=opts();o['months']=[8,9];g=analyze(f,o)['groups'][0]
    assert g['ppm']==pytest.approx(.14/4.4*1e6)

@pytest.mark.skipif(not (Path(__file__).parents[1]/'template_files'/'KABLO HURDA ANALİZİ.xls').exists(),reason='Private reference workbook is not distributed')
def test_legacy_reader_preserves_sample():
    rows=read_sheets(Path(__file__).parents[1]/'template_files'/'KABLO HURDA ANALİZİ.xls')['EYLÜL']
    assert len(rows)>1
    assert any(isinstance(v,(int,float)) for row in rows[1:] for v in row)

def test_cost_exclusion_does_not_mutate_source():
    f=fixture();before=deepcopy(f);analyze(f,opts());assert f==before

def test_ad_to_grams_cable_uses_weight_without_adding_units():
    f=fixture();f['weights']['records'][0]['unit']='AD'
    for r in f['scrap']['records']+f['fragmented']['records']:
        if r['code'].startswith('202'):r['unit']='AD'
    g=analyze(f,opts())['groups'][0]
    assert g['stocks'][0]['unit']=='AD'
    assert g['pareto_unit']=='kg'
    assert g['pareto'][0]['qty']==pytest.approx(.1)
    assert g['ppm']==25000

def test_real_zero_consumption_is_not_missing_data():
    f=fixture();next(r for r in f['consumption']['records'] if r['code']=='206000100')['qty']=0
    r=analyze(f,opts());g=next(g for g in r['groups'] if g['group']=='206')
    assert not r['issues']
    assert g['scrap']==20 and g['cost_tl']==40 and g['ppm'] is None

def test_http_session_isolation_and_reset():
    from app import app,STATE
    a=app.test_client();b=app.test_client()
    a.get('/api/state');b.get('/api/state')
    with a.session_transaction() as s:STATE[s['id']]['inputs']=fixture()
    response=a.post('/api/analyze',json=opts())
    assert response.status_code==200
    assert a.get('/api/chart/202/ppm.png').data.startswith(b'\x89PNG')
    assert a.get('/api/charts/export?format=pdf').data.startswith(b'%PDF')
    assert b.get('/api/state').json['files']=={}
    a.post('/api/reset');assert a.get('/api/state').json['files']=={}
