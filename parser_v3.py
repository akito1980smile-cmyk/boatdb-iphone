"""Conservative parser for BOAT RACE official K fixed-width text.
Only emit confirmed finish rows; keep original line for audit.
"""
import re
import unicodedata

VENUE_NAMES = ['桐生','戸田','江戸川','平和島','多摩川','浜名湖','蒲郡','常滑','津','三国','びわこ','住之江','尼崎','鳴門','丸亀','児島','宮島','徳山','下関','若松','芦屋','福岡','唐津','大村']

def norm(s):
    return unicodedata.normalize('NFKC',s).replace('\u3000',' ')

def compact(s):
    return re.sub(r'\s+','',norm(s))

# A finish line begins with two-digit placing, boat number, four-digit racer id.
FINISH = re.compile(r'^\s*(0[1-6])\s+([1-6])\s+(\d{4})\s+(.+?)\s+(\d{1,2})\s+(\d{1,2})\s+(\d\.\d{2})\s+([1-6])\s+([FL]?\.?\d{1,2}\.\d{2}|[FL]\.\d{2}|[FL])(?:\s+(\d\.\d{2}\.\d))?\s*$')
# Rows without a recorded race time (typically fifth/sixth place) are valid.
# Some archives also omit the exhibition time; keep it blank rather than inventing it.
FINISH_NO_EXHIBITION = re.compile(r'^\s*(0[1-6])\s+([1-6])\s+(\d{4})\s+(.+?)\s+(\d{1,2})\s+(\d{1,2})\s+([1-6])\s+([FL]?\.?\d{1,2}\.\d{2}|[FL]\.\d{2}|[FL])(?:\s+(\d\.\d{2}\.\d))?\s*$')
# Official K results may mark non-finishers using a short status in place of 01..06.
# Do not assign an artificial placing to such entries.
NON_FINISH = re.compile(r'^\s*(?:欠|失|転|落|沈|妨|不|エ|Ｆ|Ｌ|F|L)\s+[1-6]\s+\d{4}\b')

# payout summary row (only within a confirmed venue's payout section)
PAY = re.compile(r'^\s*(1[0-2]|[1-9])R\s+([1-6]-[1-6]-[1-6])\s+([\d,]+)\b')
RACE = re.compile(r'^\s*(1[0-2]|[1-9])R\s+.+?H\s*(\d{3,4})m\b',re.I)

def parse_lines(lines):
    """lines: iterable (day, kind, filename, line_number, original_text)."""
    results=[]; payouts=[]; rejected=[]
    current_file=None; venue=''; race=None; distance=None; in_payout=False
    for day,kind,filename,number,raw in lines:
        if kind!='K':continue
        key=(str(day),filename)
        if key!=current_file:
            current_file=key;venue='';race=None;distance=None;in_payout=False
        clean=norm(raw); tight=compact(raw)
        if '[成績]' in clean:
            label=tight.split('[成績]')[0]
            matches=[v for v in VENUE_NAMES if v==label]
            if matches:venue=matches[0];race=None
        if 'ボートレース' in tight:
            label=tight.split('ボートレース',1)[1]
            matches=[v for v in VENUE_NAMES if label.startswith(v)]
            if matches:venue=matches[0]
        if '[払戻金]' in clean:in_payout=True
        m=RACE.match(clean)
        if m:
            race=int(m.group(1));distance=int(m.group(2));in_payout=False
        if in_payout and venue:
            p=PAY.match(clean)
            if p:
                # Official summary has four ordered combination / amount pairs.
                tokens=clean[p.end():].split()
                pairs=[(p.group(2),p.group(3))]
                for i in range(0,min(len(tokens)-1,6),2):
                    if re.fullmatch(r'[1-6](?:-[1-6]){1,2}',tokens[i]) and re.fullmatch(r'[\d,]+',tokens[i+1]):
                        pairs.append((tokens[i],tokens[i+1]))
                row={'日付':str(day),'競艇場':venue,'R':int(p.group(1)),'原文':raw}
                for title,(combo,amount) in zip(('三連単','三連複','二連単','二連複'),pairs):
                    row[title+'組番']=combo;row[title+'払戻金']=int(amount.replace(',',''))
                payouts.append(row)
        if not race or not venue:continue
        # Only consider rows whose prefix resembles an official finish entry.
        if not re.match(r'^\s*(?:0[1-6])\s+[1-6]\s+\d{4}\b',clean):continue
        f=FINISH.match(clean)
        if f:
            rank,boat,racer,name,motor,hull,exhibition,course,st,rt=f.groups()
        else:
            f=FINISH_NO_EXHIBITION.match(clean)
            if f:
                rank,boat,racer,name,motor,hull,course,st,rt=f.groups()
                exhibition=None
            else:
                rejected.append({'日付':str(day),'競艇場':venue,'R':race,'行番号':number,'原文':raw,'理由':'着順行の書式と一致しない（要確認）'})
                continue
        results.append({'日付':str(day),'競艇場':venue,'R':race,'距離m':distance,'着順':int(rank),'艇番':int(boat),'登録番号':racer,'選手名':compact(name),'モーター':int(motor),'ボート':int(hull),'展示タイム':float(exhibition) if exhibition is not None else '', '進入コース':int(course),'ST':st,'レースタイム':rt or '', '原文':raw})
    return results,payouts,rejected

RESULT_HEADERS=['日付','競艇場','R','距離m','着順','艇番','登録番号','選手名','モーター','ボート','展示タイム','進入コース','ST','レースタイム','原文']
PAY_HEADERS=['日付','競艇場','R','三連単組番','三連単払戻金','三連複組番','三連複払戻金','二連単組番','二連単払戻金','二連複組番','二連複払戻金','原文']
REJECT_HEADERS=['日付','競艇場','R','行番号','原文','理由']
