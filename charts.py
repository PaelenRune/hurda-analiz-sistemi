"""The same plot is displayed in the UI and exported to PDF/PNG."""
import io
import threading
import zipfile
import matplotlib
matplotlib.use('Agg')
from matplotlib.figure import Figure
from matplotlib.backends.backend_pdf import PdfPages
from analysis_engine import DATA_LABELS

PLOT_LOCK=threading.Lock()

def figure(result,group,kind):
    fig=Figure(figsize=(11.7,6.2),layout='constrained',facecolor='white')
    ax=fig.subplots();color='#2454cf'
    period=f"{result['year']} · {', '.join(result['vouchers'])}"
    if kind=='ppm':
        rr=[r for r in group['monthly'] if r['ppm'] is not None]
        ax.plot([r['label'] for r in group['monthly']],[r['ppm'] if r['ppm'] is not None else float('nan') for r in group['monthly']],color=color,marker='o',linewidth=2.3,label='Gerçekleşen PPM')
        if group['target'] is not None:ax.axhline(group['target'],color='#d68a29',linestyle='--',label=f"Hedef: {group['target']:g}")
        ax.set_ylabel('PPM');ax.legend(loc='upper left') if rr else None
        if not rr:ax.text(.5,.5,'PPM hesaplanamaz: tüketim sıfır.' if group['stocks'] and not group['consumption'] else DATA_LABELS.get(group.get('data_status'),'Analize uygun veri yok'),ha='center',transform=ax.transAxes)
        title=f"{group['name']} — aylık PPM"
    else:
        rows=group['pareto'];x=list(range(len(rows)))
        ax.bar(x,[r['share']*100 for r in rows],color=color,width=.6,label='Hurda payı')
        ax.set_xticks(x,[str(r['error']) for r in rows]);ax.set_xlabel('Hata kodu');ax.set_ylabel('Hurda payı (%)');ax.set_ylim(0,105)
        right=ax.twinx();right.plot(x,[r['cumulative']*100 for r in rows],color='#d68a29',marker='o',label='Kümülatif pay');right.set_ylim(0,105);right.set_ylabel('Kümülatif pay (%)')
        if not rows:ax.text(.5,.5,DATA_LABELS.get(group.get('data_status'),'Analize uygun veri yok'),ha='center',transform=ax.transAxes)
        title=f"{group['name']} — hata nedenleri Pareto ({group.get('pareto_unit',group['qty_unit'])})"
    ax.set_title(f'{title}\n{period}',loc='left',fontsize=15,pad=22)
    ax.grid(axis='y',alpha=.14);ax.spines[['top','right']].set_visible(False)
    if result['issues']:fig.text(.01,.005,'Eksik verili stoklar analiz dışında tutulmuştur.',fontsize=9,color='#9b5116')
    return fig

def png(result,group,kind):
    with PLOT_LOCK:
        fig=figure(result,group,kind);buffer=io.BytesIO();fig.savefig(buffer,format='png',dpi=160);fig.clear();return buffer.getvalue()

def export_charts(result,fmt,selection='all'):
    selected=[(g,k) for g in result['groups'] for k in ('ppm','pareto') if selection in ('all',f"{g['group']}-{k}")]
    if not selected:raise ValueError('Grafik seçimi geçersiz.')
    buffer=io.BytesIO()
    if fmt=='pdf':
        with PLOT_LOCK, PdfPages(buffer) as pdf:
            for group,kind in selected:
                fig=figure(result,group,kind);pdf.savefig(fig);fig.clear()
        return buffer.getvalue(),'application/pdf','Analiz_Grafikleri.pdf'
    if fmt!='png':raise ValueError('PDF veya PNG seçin.')
    if len(selected)==1:return png(result,*selected[0]),'image/png',f'{selection}.png'
    with zipfile.ZipFile(buffer,'w',zipfile.ZIP_DEFLATED) as archive:
        for group,kind in selected:archive.writestr(f"{group['group']}_{kind}.png",png(result,group,kind))
    return buffer.getvalue(),'application/zip','Analiz_Grafikleri_PNG.zip'
