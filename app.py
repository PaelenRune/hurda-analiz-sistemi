from __future__ import annotations
import io
import os
import hashlib
import secrets
import threading
import time
from pathlib import Path
from flask import Flask, jsonify, request, session, send_file, render_template
from flask import redirect, url_for
from werkzeug.security import check_password_hash
from flask.wrappers import Request
from analysis_engine import parse_file, analyze, InputError, MONTHS
from charts import png, export_charts
from excel_export import export_excel

app=Flask(__name__)
class MemoryRequest(Request):
    def _get_file_stream(self, total_content_length, content_type, filename=None, content_length=None):
        return io.BytesIO()
app.request_class=MemoryRequest
WEB_MODE=os.environ.get('APP_MODE','local')=='web'
app.secret_key=os.environ.get('SECRET_KEY') or secrets.token_bytes(32)
app.config.update(WEB_MODE=WEB_MODE,MAX_CONTENT_LENGTH=80*1024*1024,
                  MAX_FORM_MEMORY_SIZE=80*1024*1024,SESSION_COOKIE_HTTPONLY=True,
                  SESSION_COOKIE_SAMESITE='Strict',SESSION_COOKIE_SECURE=WEB_MODE,
                  PASSWORD_HASH=os.environ.get('APP_PASSWORD_HASH',''),
                  SESSION_TTL_SECONDS=1800 if WEB_MODE else 7200)
if WEB_MODE and (not os.environ.get('SECRET_KEY') or not app.config['PASSWORD_HASH']):
    raise RuntimeError('Web ortamında SECRET_KEY ve APP_PASSWORD_HASH zorunludur.')
STATE={};LOCK=threading.RLock()
LOGIN_ATTEMPTS={}
ROLES={'scrap','consumption','weights','costs','fragmented'}

def csrf_token():
    if not app.config['WEB_MODE']:return ''
    return session.setdefault('csrf_token',secrets.token_urlsafe(32))

@app.context_processor
def csrf_context():return {'csrf_token':csrf_token}

def state():
    with LOCK:
        now=time.monotonic()
        for key in list(STATE):
            if now-STATE[key]['touched']>app.config['SESSION_TTL_SECONDS']:del STATE[key]
        sid=session.setdefault('id',secrets.token_urlsafe(24))
        value=STATE.setdefault(sid,{'inputs':{},'result':None,'touched':now})
        value['touched']=now
        return value

@app.before_request
def local_requests():
    if not app.config['WEB_MODE'] and request.host.split(':')[0] not in ('127.0.0.1','localhost'):return jsonify(error='Yerel bağlantı gerekli.'),403
    if request.method not in ('GET','HEAD','OPTIONS'):
        if app.config['WEB_MODE']:
            expected=session.get('csrf_token','')
            supplied=request.form.get('csrf_token','') if request.endpoint=='login' else request.headers.get('X-CSRF-Token','')
            if not expected or not supplied or not secrets.compare_digest(expected.encode(),supplied.encode()):
                error='Sayfa oturumu geçersiz veya süresi dolmuş. Sayfayı yenileyip tekrar deneyin.'
                if request.endpoint=='login':return render_template('login.html',error=error),403
                return jsonify(error=error),403
        else:
            origin=request.headers.get('Origin')
            from urllib.parse import urlsplit
            if origin and urlsplit(origin).netloc!=request.host:return jsonify(error='Geçersiz istek kaynağı.'),403
    if app.config['WEB_MODE'] and request.endpoint not in ('login','health','static') and not session.get('authenticated'):
        if request.path.startswith('/api/'):return jsonify(error='Oturum açmanız gerekiyor.'),401
        return redirect(url_for('login'))

@app.after_request
def private_response(response):
    response.headers['Cache-Control']='no-store, private'
    response.headers['Pragma']='no-cache'
    response.headers['X-Content-Type-Options']='nosniff'
    response.headers['Referrer-Policy']='no-referrer'
    response.headers['Content-Security-Policy']="default-src 'self'; script-src 'self'; style-src 'self' 'unsafe-inline'; img-src 'self' blob: data:; frame-ancestors 'none'; form-action 'self'; base-uri 'self'"
    return response

@app.get('/health')
def health():return jsonify(ok=True)

@app.route('/login',methods=['GET','POST'])
def login():
    if not app.config['WEB_MODE']:return redirect('/')
    error=None;status_code=200
    if request.method=='POST':
        peer=hashlib.sha256((request.remote_addr or '').encode()).hexdigest()
        now=time.monotonic()
        with LOCK:
            for key in list(LOGIN_ATTEMPTS):
                LOGIN_ATTEMPTS[key]=[t for t in LOGIN_ATTEMPTS[key] if now-t<600]
                if not LOGIN_ATTEMPTS[key]:del LOGIN_ATTEMPTS[key]
            attempts=LOGIN_ATTEMPTS.setdefault(peer,[])
            blocked=len(attempts)>=10
            if not blocked:attempts.append(now)
        if blocked:error='Çok fazla deneme. On dakika sonra tekrar deneyin.';status_code=429
        elif check_password_hash(app.config['PASSWORD_HASH'],request.form.get('password','')):
            with LOCK:STATE.pop(session.get('id'),None);LOGIN_ATTEMPTS.pop(peer,None)
            session.clear();session['authenticated']=True
            return redirect('/')
        else:error='Şifre hatalı.';status_code=401
    return render_template('login.html',error=error),status_code

@app.post('/api/logout')
def logout():
    with LOCK:STATE.pop(session.get('id'),None)
    session.clear()
    return jsonify(ok=True)

def expire_sessions():
    while True:
        time.sleep(30)
        with LOCK:
            now=time.monotonic()
            for key in list(STATE):
                if now-STATE[key]['touched']>app.config['SESSION_TTL_SECONDS']:del STATE[key]
threading.Thread(target=expire_sessions,daemon=True).start()

@app.errorhandler(InputError)
def invalid(exc):return jsonify(error=str(exc)),400

@app.errorhandler(413)
def too_large(exc):return jsonify(error='Dosya boyutu 80 MB sınırını aşıyor.'),413

@app.errorhandler(Exception)
def unexpected(exc):
    from werkzeug.exceptions import HTTPException
    if isinstance(exc,HTTPException):return jsonify(error=exc.description),exc.code
    app.logger.error('İşlem tamamlanamadı: %s',type(exc).__name__)
    return jsonify(error='İşlem tamamlanamadı. Girdi biçimini kontrol edip tekrar deneyin.'),500

@app.get('/')
def index():return render_template('index.html',web_mode=app.config['WEB_MODE'])

@app.get('/api/state')
def status():
    s=state()
    return jsonify(files={k:{'filename':v['filename'],'count':len(v['records']),'warnings':v['warnings'],'hint':v['hint']} for k,v in s['inputs'].items()},has_result=bool(s['result']))

@app.post('/api/upload/<role>')
def upload(role):
    if role not in ROLES:raise InputError('Geçersiz girdi türü.')
    file=request.files.get('file')
    if not file or not file.filename:raise InputError('Excel dosyası seçin.')
    suffix=Path(file.filename).suffix.lower()
    if suffix not in ('.xls','.xlsx'):raise InputError('XLS veya XLSX dosyası yükleyin.')
    parsed=parse_file(file.stream,role,filename='input'+suffix)
    parsed['filename']=Path(file.filename.replace('\\','/')).name
    with LOCK:s=state();s['inputs'][role]=parsed;s['result']=None
    periods=sorted({(r['year'],r['month']) for r in parsed['records'] if r.get('year') and r.get('month')})
    return jsonify(filename=parsed['filename'],count=len(parsed['records']),warnings=parsed['warnings'],hint=parsed['hint'],periods=periods)

@app.delete('/api/upload/<role>')
def remove(role):
    with LOCK:s=state();s['inputs'].pop(role,None);s['result']=None
    return jsonify(ok=True)

@app.post('/api/analyze')
def run():
    options=request.get_json() or {}
    s=state();inputs=dict(s['inputs'])
    if not {'scrap','consumption','weights','costs'}<=inputs.keys():raise InputError('Hurdalar, Tüketim, Gramaj ve Maliyet dosyalarını yükleyin.')
    if not inputs.get('fragmented') and not options.get('no_fragment'):raise InputError('Parçalanmış hurda dosyasını yükleyin veya bu dönem için parçalanmış hurda olmadığını işaretleyin.')
    with LOCK:s['result']=None
    try:result=analyze(inputs,options)
    except (KeyError,TypeError,ValueError) as exc:
        if isinstance(exc,InputError):raise
        raise InputError('Filtre seçimleri geçersiz: '+str(exc))
    if not result['requires_confirmation']:
        with LOCK:s['result']=result
    return jsonify(result)

def result():
    value=state()['result']
    if not value:raise InputError('Önce analizi çalıştırın.')
    return value

@app.get('/api/chart/<group>/<kind>.png')
def chart(group,kind):
    r=result();g=next((g for g in r['groups'] if g['group']==group),None)
    if g is None or kind not in ('ppm','pareto'):raise InputError('Grafik bulunamadı.')
    return send_file(io.BytesIO(png(r,g,kind)),mimetype='image/png',max_age=0)

@app.get('/api/charts/export')
def charts_export():
    data,mime,name=export_charts(result(),request.args.get('format','pdf'),request.args.get('selection','all'))
    return send_file(io.BytesIO(data),mimetype=mime,as_attachment=True,download_name=name)

@app.post('/api/export/excel')
def excel_export():
    if app.config['WEB_MODE']:
        from web_excel import export_zip
        return send_file(io.BytesIO(export_zip(result())),mimetype='application/zip',as_attachment=True,download_name='Hurda_Analiz_Excel.zip')
    return jsonify(export_excel(result()))

@app.post('/api/reset')
def reset():
    with LOCK:STATE.pop(session.get('id'),None)
    authenticated=session.get('authenticated')
    token=session.get('csrf_token')
    session.clear()
    if authenticated:session['authenticated']=True
    if token:session['csrf_token']=token
    return jsonify(ok=True)

if __name__=='__main__':
    import argparse
    parser=argparse.ArgumentParser();parser.add_argument('--port',type=int,default=8770);args=parser.parse_args()
    app.run(host='127.0.0.1',port=args.port,debug=False,threaded=True)
