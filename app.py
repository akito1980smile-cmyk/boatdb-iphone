import csv, io, os, re, shutil, subprocess, tempfile, zipfile
from datetime import date, timedelta
from pathlib import Path
from urllib.parse import quote
import requests
from fastapi import FastAPI, HTTPException, Query
from fastapi.responses import HTMLResponse, StreamingResponse, JSONResponse
from fastapi.templating import Jinja2Templates
from fastapi import Request
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

app = FastAPI(title='競艇公式データベース')
templates = Jinja2Templates(directory=str(Path(__file__).parent/'templates'))
VENUES=['桐生','戸田','江戸川','平和島','多摩川','浜名湖','蒲郡','常滑','津','三国','びわこ','住之江','尼崎','鳴門','丸亀','児島','宮島','徳山','下関','若松','芦屋','福岡','唐津','大村']
BASE='https://www1.mbrace.or.jp/od2'
MAX_DAYS=31
VENUE_CODES={f'{i:02d}':name for i,name in enumerate(VENUES,1)}

def source_url(day, kind):
    if kind not in ('K','B'): raise ValueError('invalid type')
    return f'{BASE}/{kind}/{day:%Y%m}/{kind.lower()}{day:%y%m%d}.lzh'

def fetch(day, kind):
    url=source_url(day,kind)
    try:
        r=requests.get(url,timeout=25,headers={'User-Agent':'BoatDB-PersonalResearch/1.0'})
        if r.status_code==404: return None
        r.raise_for_status()
        if len(r.content)>8_000_000: raise ValueError('ファイルサイズ上限を超えました')
        return r.content
    except requests.RequestException as exc: raise HTTPException(502,f'公式サイトへの接続に失敗: {exc}')

def extract_lzh(data):
    # Requires lhasa (`lha`) installed on the server; extraction is isolated to a temp directory.
    if not shutil.which('lha'): raise HTTPException(503,'サーバーに lhasa (lha) がインストールされていません')
    with tempfile.TemporaryDirectory() as tmp:
        archive=Path(tmp)/'source.lzh'; archive.write_bytes(data)
        out=Path(tmp)/'out'; out.mkdir()
        result=subprocess.run(['lha','xw='+str(out),str(archive)],capture_output=True,timeout=20)
        if result.returncode: raise HTTPException(422,'公式LZHファイルを解凍できませんでした')
        files=[p for p in out.rglob('*') if p.is_file() and not p.is_symlink()]
        if not files: raise HTTPException(422,'解凍したファイルが空です')
        lines=[]
        for p in files:
            if p.stat().st_size>15_000_000: continue
            raw=p.read_bytes()
            for enc in ('cp932','shift_jis','utf-8'):
                try: text=raw.decode(enc);break
                except UnicodeDecodeError: continue
            else: text=raw.decode('cp932',errors='replace')
            lines.extend((p.name,n,line) for n,line in enumerate(text.splitlines(),1) if line.strip())
        return lines

def records(start,end,kind):
    day=start
    while day<=end:
        data=fetch(day,kind)
        if data:
            for filename,n,line in extract_lzh(data):
                # Preserve official fixed-width text losslessly rather than falsely assigning fields.
                yield [day.isoformat(),kind,filename,n,line]
        day+=timedelta(days=1)

def bounds(start,end):
    if start>end: raise HTTPException(400,'開始日が終了日より後です')
    if end>date.today(): raise HTTPException(400,'未来の日付は取得できません')
    if (end-start).days>=MAX_DAYS: raise HTTPException(400,f'一度に取得できるのは最大{MAX_DAYS}日です')

@app.get('/',response_class=HTMLResponse)
def home(request:Request):return templates.TemplateResponse(request=request,name='index.html',context={'venues':VENUES})

@app.get('/api/status')
def status():return {'source':'BOAT RACE公式ダウンロード','venues':24,'max_days':MAX_DAYS,'extractor_ready':bool(shutil.which('lha')),'mode':'official-original-text'}

@app.get('/api/export')
def export(start:date,end:date,kind:str=Query('K',pattern='^(K|B)$'),format:str=Query('xlsx',pattern='^(xlsx|csv)$'),venue:int=Query(0,ge=0,le=24)):
    bounds(start,end)
    # Official source is day-level nationwide. Venue filtering of fixed-width data is NOT
    # offered until a verified parser exists; reject rather than silently return wrong data.
    if venue: raise HTTPException(422,'場別の抽出は公式固定長レイアウトの解析が必要です。現在は全国24場一括出力を選択してください。')
    headers=['日付','データ種別','公式ファイル名','行番号','公式原文']
    rows=list(records(start,end,kind))
    if not rows: raise HTTPException(404,'指定期間に取得できる公式データがありません')
    filename=f'boatrace_{kind}_{start:%Y%m%d}_{end:%Y%m%d}.{format}'
    if format=='csv':
        out=io.StringIO(newline='');writer=csv.writer(out);writer.writerow(headers);writer.writerows(rows)
        payload=out.getvalue().encode('utf-8-sig');media='text/csv; charset=utf-8'
    else:
        wb=Workbook();ws=wb.active;ws.title='公式データ';ws.append(headers)
        for row in rows: ws.append([str(x)[:32767] for x in row])
        ws.freeze_panes='A2';ws.auto_filter.ref=ws.dimensions
        for c in ws[1]: c.font=Font(color='FFFFFF',bold=True);c.fill=PatternFill('solid',fgColor='163C66')
        for col,width in {'A':15,'B':15,'C':24,'D':12,'E':95}.items():ws.column_dimensions[col].width=width
        info=wb.create_sheet('説明');info.append(['項目','内容']);info.append(['出典','BOAT RACEオフィシャルウェブサイト']);info.append(['形式','公式配布ファイルの行単位原文。項目解析済みデータではありません']);info.append(['場別抽出','未実装（固定長レイアウトの検証が必要）'])
        out=io.BytesIO();wb.save(out);payload=out.getvalue();media='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    return StreamingResponse(io.BytesIO(payload),media_type=media,headers={'Content-Disposition':f"attachment; filename={quote(filename)}"})


# Conservative line-level index: never guess fixed-width column offsets.
def classify_line(line):
    import unicodedata
    normalized=unicodedata.normalize('NFKC',line)
    venue=next((v for v in VENUES if v in normalized), '')
    race_match=re.search(r'(?<!\d)(1[0-2]|[1-9])\s*R(?![A-Za-z])', normalized,re.I)
    racer_ids=re.findall(r'(?<!\d)([2-5]\d{3})(?!\d)',normalized)
    return {'競艇場候補':venue,'レース番号候補':int(race_match.group(1)) if race_match else '',
            '選手登録番号候補':','.join(dict.fromkeys(racer_ids[:12]))}

def indexed_rows(start,end,kind,keyword='',venue=0):
    if len(keyword)>80: raise HTTPException(400,'検索語が長すぎます')
    for day,k,filename,n,line in records(start,end,kind):
        fields=classify_line(line)
        # Only return confirmed textual venue mentions, not inferred sections.
        if venue and fields['競艇場候補']!=VENUES[venue-1]: continue
        if keyword and keyword.casefold() not in line.casefold(): continue
        yield [day,k,filename,n,fields['競艇場候補'],fields['レース番号候補'],fields['選手登録番号候補'],line]

@app.get('/api/search')
def search(start:date,end:date,kind:str=Query('K',pattern='^(K|B)$'),keyword:str='',venue:int=Query(0,ge=0,le=24),limit:int=Query(100,ge=1,le=500)):
    bounds(start,end)
    from itertools import islice
    rows=list(islice(indexed_rows(start,end,kind,keyword,venue),limit+1))
    return {'columns':['日付','種別','ファイル名','行番号','競艇場候補','レース番号候補','選手登録番号候補','公式原文'],
            'rows':rows[:limit], 'truncated':len(rows)>limit,
            'notice':'候補列は文字列検出による参考値です。公式固定長仕様に基づくレース・選手単位の確定解析ではありません。'}

@app.get('/api/export-index')
def export_index(start:date,end:date,kind:str=Query('K',pattern='^(K|B)$'),format:str=Query('xlsx',pattern='^(xlsx|csv)$'),keyword:str='',venue:int=Query(0,ge=0,le=24)):
    bounds(start,end)
    headers=['日付','種別','ファイル名','行番号','競艇場候補','レース番号候補','選手登録番号候補','公式原文']
    rows=list(indexed_rows(start,end,kind,keyword,venue))
    if not rows: raise HTTPException(404,'条件に一致する公式原文がありません')
    filename=f'boatrace_index_{start:%Y%m%d}_{end:%Y%m%d}.{format}'
    if format=='csv':
        out=io.StringIO(newline='');w=csv.writer(out);w.writerow(headers);w.writerows(rows)
        payload=out.getvalue().encode('utf-8-sig');media='text/csv; charset=utf-8'
    else:
        wb=Workbook();ws=wb.active;ws.title='検索結果';ws.append(headers)
        for row in rows[:1048575]:ws.append([str(x)[:32767] for x in row])
        ws.freeze_panes='A2';ws.auto_filter.ref=ws.dimensions
        for c in ws[1]:c.font=Font(color='FFFFFF',bold=True);c.fill=PatternFill('solid',fgColor='163C66')
        for col,width in {'A':15,'B':10,'C':23,'D':10,'E':17,'F':18,'G':26,'H':90}.items():ws.column_dimensions[col].width=width
        info=wb.create_sheet('注意事項')
        for row in [['出典','BOAT RACE公式配布ファイル'],['解析状況','文字列による候補抽出。着順・勝率・平均ST等の確定解析は未対応'],['場別検索','場名が同一行に明記される行のみ対象。全レースを網羅しません'],['利用','公式サイトの利用条件を確認してください']]:info.append(row)
        out=io.BytesIO();wb.save(out);payload=out.getvalue();media='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    return StreamingResponse(io.BytesIO(payload),media_type=media,headers={'Content-Disposition':f"attachment; filename={quote(filename)}"})


# v3: structured extraction from official competition-result files
from parser_v3 import parse_lines, RESULT_HEADERS, PAY_HEADERS, REJECT_HEADERS

@app.get('/api/export-v3')
def export_v3(start:date,end:date,format:str=Query('xlsx',pattern='^(xlsx|csv)$'),venue:int=Query(0,ge=0,le=24)):
    bounds(start,end)
    results,payouts,rejected=parse_lines(records(start,end,'K'))
    if venue:
        selected=VENUES[venue-1]
        results=[r for r in results if r['競艇場']==selected]
        payouts=[r for r in payouts if r['競艇場']==selected]
        rejected=[r for r in rejected if r['競艇場']==selected]
    if not results and not payouts:
        raise HTTPException(404,'対象期間の解析可能な公式成績データがありません')
    if format=='csv':
        # CSV is one flat runner table with matching payout attached.
        pmap={(p['日付'],p['競艇場'],p['R']):p for p in payouts}
        headers=RESULT_HEADERS[:-1]+['三連単組番','三連単払戻金','原文']
        out=io.StringIO(newline='');w=csv.writer(out);w.writerow(headers)
        for r in results:
            p=pmap.get((r['日付'],r['競艇場'],r['R']),{})
            w.writerow([r.get(k,'') for k in RESULT_HEADERS[:-1]]+[p.get('三連単組番',''),p.get('三連単払戻金',''),r['原文']])
        payload=out.getvalue().encode('utf-8-sig');media='text/csv; charset=utf-8'
    else:
        wb=Workbook();ws=wb.active;ws.title='着順・ST';ws.append(RESULT_HEADERS)
        for r in results:ws.append([r.get(k,'') for k in RESULT_HEADERS])
        ps=wb.create_sheet('三連単払戻');ps.append(PAY_HEADERS)
        for r in payouts:ps.append([r.get(k,'') for k in PAY_HEADERS])
        rs=wb.create_sheet('要確認');rs.append(REJECT_HEADERS)
        for r in rejected:rs.append([r.get(k,'') for k in REJECT_HEADERS])
        info=wb.create_sheet('説明');info.append(['項目','説明'])
        for row in [('出典','BOAT RACE公式競走成績（Kファイル）'),('解析','通常着順行のみ厳格抽出。特殊着順は要確認へ'),('ST','F/L等の記号を含む可能性があるため文字列として保存'),('三連単','払戻金は100円あたりの公式表示値'),('注意','抽出結果は公式原文と照合してください'),('期間','最大31日／回。過去全期間の自動保存は未実装')]:info.append(row)
        for sheet in wb:
            sheet.freeze_panes='A2';sheet.auto_filter.ref=sheet.dimensions
            for c in sheet[1]:c.font=Font(color='FFFFFF',bold=True);c.fill=PatternFill('solid',fgColor='163C66')
            for col in sheet.columns:
                letter=col[0].column_letter
                sheet.column_dimensions[letter].width=min(65,max(13,max(len(str(c.value or '')) for c in list(col)[:100])+3))
        out=io.BytesIO();wb.save(out);payload=out.getvalue();media='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    filename=f'boatrace_v3_{start:%Y%m%d}_{end:%Y%m%d}.{format}'
    return StreamingResponse(io.BytesIO(payload),media_type=media,headers={'Content-Disposition':f'attachment; filename="{filename}"'})


# v3 structured export. Unrecognized fixed-width rows are preserved in a review sheet.
from parser_v3 import parse_lines, RESULT_HEADERS, PAY_HEADERS, REJECT_HEADERS

@app.get('/api/v3/preview')
def v3_preview(start:date,end:date,venue:int=Query(0,ge=0,le=24)):
    bounds(start,end)
    results,payouts,rejected=parse_lines(records(start,end,'K'))
    if venue:
        selected=VENUES[venue-1]
        results=[x for x in results if x['競艇場']==selected]
        payouts=[x for x in payouts if x['競艇場']==selected]
        rejected=[x for x in rejected if x['競艇場']==selected]
    return {'results':results[:100],'payouts':payouts[:100], 'counts':{'results':len(results),'payouts':len(payouts),'rejected':len(rejected)},'notice':'未解析の着順行はExcelの要確認シートに記録します。正確性は公式原文との照合が必要です。'}

@app.get('/api/v3/export')
def v3_export(start:date,end:date,venue:int=Query(0,ge=0,le=24),format:str=Query('xlsx',pattern='^(xlsx|csv)$')):
    bounds(start,end)
    results,payouts,rejected=parse_lines(records(start,end,'K'))
    if venue:
        selected=VENUES[venue-1]
        results=[x for x in results if x['競艇場']==selected]
        payouts=[x for x in payouts if x['競艇場']==selected]
        rejected=[x for x in rejected if x['競艇場']==selected]
    if not results and not payouts and not rejected:
        raise HTTPException(404,'解析可能なデータがありません。期間・配布状況を確認してください。')
    filename=f'boatrace_v3_{start:%Y%m%d}_{end:%Y%m%d}.{format}'
    if format=='csv':
        # CSV supports one table; payout fields are joined to each confirmed finish row.
        headers=RESULT_HEADERS+[x for x in PAY_HEADERS if x not in ('日付','競艇場','R','原文')]
        lookup={(x['日付'],x['競艇場'],x['R']):x for x in payouts}
        out=io.StringIO(newline='');w=csv.writer(out);w.writerow(headers)
        for r in results:
            pay=lookup.get((r['日付'],r['競艇場'],r['R']),{})
            w.writerow([r.get(k,'') for k in RESULT_HEADERS]+[pay.get(k,'') for k in headers[len(RESULT_HEADERS):]])
        payload=out.getvalue().encode('utf-8-sig');media='text/csv; charset=utf-8'
    else:
        wb=Workbook();wb.remove(wb.active)
        for title,headers,items in [('着順・ST',RESULT_HEADERS,results),('払戻金',PAY_HEADERS,payouts),('要確認',REJECT_HEADERS,rejected)]:
            ws=wb.create_sheet(title);ws.append(headers)
            for r in items:ws.append([str(r.get(k,''))[:32767] if isinstance(r.get(k,''),str) else r.get(k,'') for k in headers])
            ws.freeze_panes='A2';ws.auto_filter.ref=ws.dimensions
            for c in ws[1]:c.font=Font(color='FFFFFF',bold=True);c.fill=PatternFill('solid',fgColor='163C66')
            for col in ws.columns:
                letter=col[0].column_letter
                ws.column_dimensions[letter].width=min(65,max(12,len(str(col[0].value))*2+3))
        ws=wb.create_sheet('説明');ws.append(['項目','内容']);ws.append(['出典','BOAT RACE公式配布の競走成績(K)']);ws.append(['注意','抽出結果は原文との照合が必要。認識できない行は要確認に保存。']);ws.append(['CSV','着順に払戻情報を結合した単一表。未認識行は含まれません。']);ws.append(['件数',f'着順 {len(results)} / 払戻 {len(payouts)} / 要確認 {len(rejected)}'])
        out=io.BytesIO();wb.save(out);payload=out.getvalue();media='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    return StreamingResponse(io.BytesIO(payload),media_type=media,headers={'Content-Disposition':f'attachment; filename={filename}'})
