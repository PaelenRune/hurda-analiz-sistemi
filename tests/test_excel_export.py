"""Integration check using Microsoft Excel, including expanding template tables."""
import importlib.util
import sys
from pathlib import Path
import openpyxl
import pytest
from analysis_engine import analyze
from excel_export import export_excel
from test_analysis import fixture,opts,row

@pytest.mark.skipif(sys.platform!='win32' or not importlib.util.find_spec('win32com') or not (Path(__file__).parents[1]/'template_files').exists(),reason='Private templates and Windows Excel required')
def test_native_excel_export_growth_and_chart_bindings(tmp_path):
    f=fixture()
    for i in range(60):
        stock=f'20699{i:04}'
        f['scrap']['records'].append(row(stock,qty=i+1,error=str(i%26+1)))
        f['costs']['records'].append(dict(code=stock,year=2026,month=9,date='2026-09-30T00:00:00',tl=2,eur=.04))
        f['consumption']['records'].append(dict(code=stock,year=2026,month=9,qty=10000))
    options=opts();options['months']=[9,12]
    r=analyze(f,options)
    out=export_excel(r,tmp_path)
    assert len(out['files'])==6
    for name in out['files']:
        wb=openpyxl.load_workbook(name,data_only=False)
        values=openpyxl.load_workbook(name,data_only=True)
        try:
            for sheet in values:
                for rr in sheet:
                    assert not any(c.data_type=='e' for c in rr),f'Excel formula error in {Path(name).name} / {sheet.title}'
            if 'PARETO' not in name:
                assert wb['Hata Kodu Pivot']._pivots
                chart=next(s._charts[0] for s in wb if s._charts)
                assert all(series.val.numRef.f for chart_ in [chart]+chart._charts[1:] for series in chart_.ser)
                assert values['GRAFİK VERİSİ']['B4'].value>0
                assert wb['EYLÜL']['C2'].number_format.startswith('#,##0.')
            else:
                assert wb['EYLÜL']._charts
                assert wb['EYLÜL']['E10'].number_format=='0.00%'
        finally:values.close();wb.close()
