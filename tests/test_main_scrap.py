import io
from copy import deepcopy
import pytest
import analysis_engine as engine
from analysis_engine import parse_file,analyze,InputError
from test_analysis import fixture,opts

def main_sheets():
    return {'Sheet1':[['YIL 2026 AY 9'],[],['Stok Kodu','Miktar','Fiş No'],['206000100',20,'Y001']],
            'HURDA':[['Stok Kodu','Toplam Miktar','EKALAN1'],['206000100',25,'18']],
            'IDC':[['Satır Etiketleri','Genel Toplam'],['206000100',99999]]}

def test_main_hurda_summary_is_the_only_quantity_source(monkeypatch):
    monkeypatch.setattr(engine,'read_sheets',lambda *a,**k:main_sheets())
    source=parse_file('any-name.xlsx','scrap')
    assert len(source['records'])==1 and source['records'][0]['qty']==25
    assert source['records'][0]['voucher']=='Y' and source['hint']==[2026,9]
    assert 'combined_baseline' not in source
    f=fixture();f['scrap']=source;o=opts();o['main_only']=True
    r=analyze(f,o);g=next(g for g in r['groups'] if g['group']=='206')
    assert g['scrap']==25 and g['cost_tl']==50 and g['ppm']==25000
    assert r['reconciliation']==[] and 'fragmented' not in r['sources']

def test_invalid_hurda_summary_does_not_fall_back_to_original_transactions(monkeypatch):
    sheets=main_sheets();sheets['HURDA']=[['Yanlış tablo'],['206000100']]
    monkeypatch.setattr(engine,'read_sheets',lambda *a,**k:sheets)
    with pytest.raises(InputError,match='Beklenen sütunlar'):parse_file('bad.xlsx','scrap')

def test_summary_without_supported_voucher_evidence_blocks(monkeypatch):
    sheets=main_sheets();del sheets['Sheet1'];sheets['HURDA'].insert(0,['YIL 2026 AY 9'])
    monkeypatch.setattr(engine,'read_sheets',lambda *a,**k:sheets)
    with pytest.raises(InputError,match='fiş grubu yok'):parse_file('bad.xlsx','scrap')

def test_four_input_api_ignores_any_legacy_fragmented_file_and_overrides_client_flag():
    from app import app,STATE
    client=app.test_client();client.get('/api/state')
    f=fixture();f['fragmented']['records'][0]['qty']=999999
    before=deepcopy(f)
    with client.session_transaction() as s:STATE[s['id']]['inputs']=f
    response=client.post('/api/analyze',json={**opts(),'main_only':False})
    assert response.status_code==200
    assert response.json['groups'][0]['stocks'][0]['scrap']==20
    assert 'fragmented' not in response.json['sources']
    assert 'fragmented' not in client.get('/api/state').json['files']
    assert f==before
    assert client.post('/api/upload/fragmented',data={'file':(io.BytesIO(b'test'),'old.xlsx')}).status_code==400
    html=client.get('/').data.decode()
    assert 'Dört kaynağı' in html and 'id="no-fragment"' not in html and 'id="fragment-year"' not in html
