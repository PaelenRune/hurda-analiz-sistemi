from copy import deepcopy
import io
import zipfile
import openpyxl
import pytest
import analysis_engine as engine
from analysis_engine import analyze,parse_file,InputError
from test_analysis import fixture,opts,row
from web_excel import export_zip


def combined_fixture(monkeypatch):
    sheets={
        'Sheet1':[['YIL 2026 AY 9'],[],['Stok Kodu','Toplam Miktar','Fiş No','Ay','EKALAN1'],
                  ['206000100',20,'Y001',9,'16'],['527000100',3,'Y002',9,'18']],
        'HURDA':[['Stok Kodu','Toplam Miktar','EKALAN1'],['206000100',25,'18']],
    }
    monkeypatch.setattr(engine,'read_sheets',lambda *a,**kw:sheets)
    f=fixture();f['fragmented']=parse_file('combined.xlsx','fragmented')
    next(r for r in f['scrap']['records'] if r['code']=='206000100')['voucher_no']='Y001'
    return f


def test_verified_combined_table_replaces_original_once_and_preserves_final_errors(monkeypatch):
    f=combined_fixture(monkeypatch);before=deepcopy(f)
    r=analyze(f,opts());g=next(g for g in r['groups'] if g['group']=='206')
    assert g['scrap']==25 and g['cost_tl']==50 and g['ppm']==25000
    assert g['pivot'][0]['errors']=={'18':25}
    assert r['reconciliation'][0]['original_qty']==20
    assert r['reconciliation'][0]['combined_qty']==25
    assert r['reconciliation'][0]['net_added_qty']==5
    assert f==before


def test_combined_mismatched_main_blocks_even_if_missing_costs_are_skipped(monkeypatch):
    f=combined_fixture(monkeypatch)
    next(r for r in f['scrap']['records'] if r['code']=='206000100')['qty']=21
    with pytest.raises(InputError,match='Kaynaklar eşleşmediği'):analyze(f,opts(skip_missing=True))


def test_same_stock_total_from_different_voucher_cannot_replace_original(monkeypatch):
    f=combined_fixture(monkeypatch)
    next(r for r in f['scrap']['records'] if r['code']=='206000100')['voucher_no']='Y_OTHER'
    with pytest.raises(InputError,match='fiş numaraları'):analyze(f,opts())


def test_combined_replacement_keeps_other_vouchers_and_other_months(monkeypatch):
    f=combined_fixture(monkeypatch)
    f['scrap']['records'] += [dict(row('206000100',4),voucher='Ç'),row('206000100',7,month=8)]
    f['costs']['records'].append(dict(code='206000100',year=2026,month=8,date='2026-08-20',tl=2,eur=.04))
    f['consumption']['records'].append(dict(code='206000100',year=2026,month=8,qty=1000))
    o=opts();o.update(months=[8,9],vouchers=['Y','Ç'])
    r=analyze(f,o);g=next(g for g in r['groups'] if g['group']=='206')
    assert g['scrap']==36 and g['consumption']==2000
    assert [m['scrap'] for m in g['monthly']]==[7,29]
    assert len(r['reconciliation'])==1


def test_combined_missing_or_changed_units_do_not_bypass_validation(monkeypatch):
    f=combined_fixture(monkeypatch)
    next(r for r in f['scrap']['records'] if r['code']=='206000100')['unit']='MT'
    with pytest.raises(InputError,match='birimler uyuşmuyor'):analyze(f,opts())


def test_equal_quantity_extra_without_embedded_baseline_is_still_additive(monkeypatch):
    sheets={'HURDA':[['YIL 2026 AY 9'],[],['Stok Kodu','Toplam Miktar','EKALAN1'],['206000100',20,'16']]}
    monkeypatch.setattr(engine,'read_sheets',lambda *a,**kw:sheets)
    f=fixture();f['fragmented']=parse_file('extra.xlsx','fragmented')
    assert 'combined_baseline' not in f['fragmented']
    r=analyze(f,opts());assert next(g for g in r['groups'] if g['group']=='206')['scrap']==40
    assert r['reconciliation']==[]


def test_explicit_combined_voucher_disambiguates_other_original_vouchers(monkeypatch):
    sheets={'Sheet1':[['YIL 2026 AY 9'],[],['Stok Kodu','Miktar','Fiş No'],['206000100',20,'Y001'],['206000100',1,'Ç001']],
            'HURDA':[['Stok Kodu','Miktar','Fiş No'],['206000100',25,'Y_FINAL']]}
    monkeypatch.setattr(engine,'read_sheets',lambda *a,**kw:sheets)
    f=fixture();f['fragmented']=parse_file('combined.xlsx','fragmented')
    next(r for r in f['scrap']['records'] if r['code']=='206000100')['voucher_no']='Y001'
    f['scrap']['records'].append(dict(row('206000100',1),voucher='Ç',voucher_no='Ç001'))
    o=opts();o['vouchers']=['Y','Ç']
    assert next(g for g in analyze(f,o)['groups'] if g['group']=='206')['scrap']==26


def test_combined_smaller_than_embedded_original_is_not_silently_added(monkeypatch):
    sheets={'HURDA':[['YIL 2026 AY 9'],[],['Stok Kodu','Miktar'],['206000100',10]],
            'Sheet1':[['Stok Kodu','Miktar','Fiş No','Tarih'],['206000100',20,'Y001','2026-09-01']]}
    monkeypatch.setattr(engine,'read_sheets',lambda *a,**kw:sheets)
    with pytest.raises(InputError,match='kaynak hareketleriyle uyuşmuyor'):parse_file('bad.xlsx','fragmented')


@pytest.mark.parametrize('date_value,month', [('31.09.2026',9),('not a date',9),(None,9.5),(None,13)])
def test_bad_date_or_period_cannot_fall_back_silently(monkeypatch,date_value,month):
    rows=[['Stok Kodu','Miktar','Tarih','Yıl','Ay'],['206000100',20,date_value,2026,month]]
    monkeypatch.setattr(engine,'read_sheets',lambda *a,**kw:{'Sheet1':rows})
    with pytest.raises(InputError):parse_file('bad.xlsx','scrap')


def test_conflicting_period_hints_and_duplicate_date_headers_block(monkeypatch):
    monkeypatch.setattr(engine,'read_sheets',lambda *a,**kw:{'A':[['YIL 2026 AY 8']],'B':[['YIL 2026 AY 9']]})
    with pytest.raises(InputError,match='birden fazla yıl/ay'):parse_file('bad.xlsx','fragmented')
    monkeypatch.setattr(engine,'read_sheets',lambda *a,**kw:{'A':[['Stok Kodu','Miktar','Tarih','Tarih'],['206000100',20,'2026-09-01','2026-08-01']]})
    with pytest.raises(InputError,match='aynı başlık'):parse_file('bad.xlsx','scrap')


def test_multiple_main_source_tables_do_not_silently_drop_a_month(monkeypatch):
    sheets={name:[['Stok Kodu','Miktar','Tarih'],['206000100',20,date]] for name,date in [('August','2026-08-01'),('September','2026-09-01')]}
    monkeypatch.setattr(engine,'read_sheets',lambda *a,**kw:sheets)
    with pytest.raises(InputError,match='Birden fazla kaynak tablosu'):parse_file('bad.xlsx','scrap')


def test_conflicting_weight_units_with_equal_grams_block():
    f=fixture();f['weights']['records'].append(dict(code='202100003-71',grams=4,unit='AD'))
    with pytest.raises(InputError,match='çelişen gramaj birimleri'):analyze(f,opts())


def test_combined_control_values_are_exported_without_changing_pivot_totals(monkeypatch):
    r=analyze(combined_fixture(monkeypatch),opts())
    with zipfile.ZipFile(io.BytesIO(export_zip(r))) as z:
        for name in z.namelist():
            if not name.startswith('IDC'):continue
            w=openpyxl.load_workbook(io.BytesIO(z.read(name)),data_only=True)
            assert list(w['Kaynak Kontrolü'].values)[1][3:6]==(20,25,5)
            if 'PARETO' not in name:
                assert w['GRAFİK VERİSİ']['C2'].value==25
                assert w['GRAFİK VERİSİ']['D2'].value==25000
                assert w['Hata Kodu Pivot']._pivots[0].cache.recordCount==1
            w.close()
