"""Läser ett befintligt samordningsmötesprotokoll (.docx, Sydbyggarnas gamla mall)
och bygger startdata för Mötesprotokoll-verktyget.

Användning:  python3 -I import_docx.py <protokoll.docx> <projektnr> <utkatalog>

Regler för tolkningen (gamla mallen -> nya strukturen):
- Rubrikrader i punkttabellen blir avsnitt.
- Varje tabellrad blir en punkt med ett fast nummer (avsnittskod + löpnummer).
- Stycken som börjar med fetstilt ÅÅMMDD startar en ny notering för det mötet.
- Text i kolumnen Beslut/svar blir noteringar av typen "Svar/beslut".
- Kursiv rad eller KLART = punkten är klar (klarmarkerad vid sista noteringens möte).
- "Gäller samtliga"/"Gäller installatörer" på öppna punkter = stående punkt.
- Alla rader "Genomgång av tidplaner" slås ihop till EN punkt med en notering per möte.
- Genomstruken text behålls, markerad ~~så här~~.
"""
import sys, re, json, os, docx

src, projnr, out = sys.argv[1], sys.argv[2], sys.argv[3]
d = docx.Document(src)

def ptext(p):
    s = []
    for r in p.runs:
        t = r.text.replace('\t', ' ')
        if not t:
            continue
        if r.font.strike:
            t = '~~' + t + '~~'
        s.append(t)
    return re.sub(r'~~~~', '', ''.join(s))

def uniq_cells(row):
    out, seen = [], set()
    for c in row.cells:
        if id(c._tc) in seen:
            continue
        seen.add(id(c._tc)); out.append(c)
    return out

# ---------- Närvarolista ----------
people = []
ini_count = {}
def parse_person(role, para, group):
    raw = ptext(para).strip()
    if not raw:
        return
    bold = any(r.bold for r in para.runs if r.text.strip())
    struck = all(r.font.strike for r in para.runs if r.text.strip())
    clean = raw.replace('~~', '')
    note = ''
    em = re.search(r'\((Ersätter[^()]*(?:\([^)]*\))?[^()]*)\)', clean)
    if em:
        note = em.group(1).strip()
        clean = (clean[:em.start()] + clean[em.end():]).strip()
    mi = re.search(r'\(([^)]+)\)', clean)
    ini = mi.group(1).strip() if mi else ''
    if mi:
        name = clean[:mi.start()].strip(' ,')
        company = clean[mi.end():].strip(' ,')
    else:
        name, _, company = clean.partition(',')
        name, company = name.strip(), company.strip()
    company = re.sub(r'\s+', ' ', company).strip(' ,')
    if not ini:
        ini = ''.join(w[0] for w in name.split()[:2]).upper()
    company = company.replace('Entreprenad BA', 'Entreprenad AB')
    pid = re.sub(r'[^a-z0-9]+', '-', name.lower().translate(str.maketrans('åäöéü', 'aaoeu'))).strip('-')
    people.append({
        'id': pid, 'name': name, 'initials': ini, 'company': company,
        'role': role.replace('~~', '').strip(), 'group': group,
        'email': '', 'phone': '', 'active': not struck, 'note': note,
        'defaultDist': 'kallad' if not struck else 'ingen',
        '_present': bold and not struck,
    })

att = d.tables[1]
for ri, group in ((0, 'Sydbyggarna'), (1, 'Underentreprenör')):
    cells = uniq_cells(att.rows[ri])
    roles = [ptext(p) for p in cells[0].paragraphs]
    for i, p in enumerate(cells[1].paragraphs):
        parse_person(roles[i] if i < len(roles) else '', p, group)
if any(p['name'].startswith('Million') for p in people):
    pass

# ---------- Mötesdatum ----------
full = '\n'.join(c.text for t in d.tables[2:] for row in t.rows for c in uniq_cells(row))
dates = sorted(set(re.findall(r'\b(2[56][01]\d[0-3]\d)\b', full)))
counts = {x: full.count(x) for x in dates}
# Datum som bara förekommer en gång och ligger nära ett annat mötesdatum är skrivfel.
def ymd(s): return (int('20' + s[:2]), int(s[2:4]), int(s[4:6]))
import datetime as dt
def todate(s): y, m, dd = ymd(s); return dt.date(y, m, dd)
meet = []
for x in dates:
    if counts[x] <= 2 and meet and (todate(x) - todate(meet[-1])).days < 7:
        continue
    meet.append(x)
# skrivfel som ligger före ett riktigt datum (260810 -> 260812)
cleaned = []
for x in meet:
    if cleaned and (todate(x) - todate(cleaned[-1])).days < 7:
        if counts[cleaned[-1]] < counts[x]:
            cleaned[-1] = x
        continue
    cleaned.append(x)
meet = cleaned
def meeting_of(s):
    best = min(meet, key=lambda m: abs((todate(m) - todate(s)).days))
    return meet.index(best) + 1

# ---------- Punkter ----------
SECTIONS = [
    ('G', 'Generella punkter – arbetsmiljö, ordning och besiktningar', r'Generella'),
    ('T', 'Tider – alla', r'Tider'),
    ('B', 'Frågor – Bygg', r'BYGG'),
    ('E', 'Frågor – El, solceller', r'El, Solceller|El'),
    ('R', 'Frågor – Rör', r'RÖR'),
    ('V', 'Frågor – Vent & Styr', r'VENT'),
    ('M', 'Frågor – Mark', r'Mark'),
    ('U', 'Frågor – Övriga entreprenörer och leverantörer', r'ÖVRIGA'),
    ('O', 'Övrigt', r'Övrigt'),
]
def section_for(text):
    for code, title, rx in SECTIONS:
        if re.search(rx, text):
            return code
    return None

DATE_RX = re.compile(r'^\s*(2[56]\d{4})\s*[:\-–]?\s*(.*)$', re.S)
def split_entries(cell, kind):
    entries, cur = [], None
    for p in cell.paragraphs:
        txt = ptext(p).rstrip()
        if not txt.strip():
            if cur and cur['text'] and not cur['text'].endswith('\n'):
                cur['text'] += '\n'
            continue
        lead = ''.join(r.text for r in p.runs[:4])
        firstbold = next((r for r in p.runs if r.text.strip()), None)
        m = DATE_RX.match(txt.replace('~~', ''))
        if m and firstbold is not None and firstbold.bold:
            cur = {'date': m.group(1), 'kind': kind, 'text': m.group(2).strip()}
            entries.append(cur)
        else:
            if cur is None:
                cur = {'date': None, 'kind': kind, 'text': ''}
                entries.append(cur)
            cur['text'] = (cur['text'] + '\n' + txt.strip()).strip('\n') if cur['text'] else txt.strip()
    for e in entries:
        e['text'] = re.sub(r'\n{3,}', '\n\n', e['text']).strip()
        e['text'] = re.sub(r'^Svar:\s*', '', e['text'])
    return [e for e in entries if e['text']]

items, tidplan = [], None
cur_sec = 'G'
tbl = d.tables[2]
seq = {}
for ri, row in enumerate(tbl.rows[1:], 1):
    cells = uniq_cells(row)
    joined = ' '.join(c.text for c in cells).strip()
    if not joined:
        continue
    if len(cells) == 1:
        sec = section_for(cells[0].text.split('Generella')[-1] if 'Förklaringar' in cells[0].text else cells[0].text)
        if 'Förklaringar' in cells[0].text:
            sec = 'G'
        if sec:
            cur_sec = sec
        continue
    rc = row.cells
    q, a = rc[1], rc[2]
    if q._tc is a._tc:
        a = None
    runs = [r for c in cells for p in c.paragraphs for r in p.runs if r.text.strip()]
    italic = bool(runs) and sum(1 for r in runs if r.italic) / len(runs) > 0.8
    qe = split_entries(q, 'notering')
    ae = split_entries(a, 'svar') if a is not None else []
    atext = a.text.strip() if a is not None else ''
    klart = italic or atext.upper().startswith('KLART')
    ae = [e for e in ae if e['text'].strip().upper() != 'KLART']
    standing = (not klart) and bool(re.match(r'Gäller samtliga', atext)) and not re.search(r'saknas|Status', q.text[:200])
    if cur_sec == 'G' and not klart and not any(re.match(r'2[56]\d{4}', p.text.strip()) for p in q.paragraphs[:1]):
        standing = True
    ae = [e for e in ae if not re.match(r'^Gäller (samtliga|installatörer)\.?\s*$', e['text'])]
    filler = lambda e: re.fullmatch(r'(Vakant|26x+:?|2[56]\d{4}:?|)', e['text'].strip(' :'))
    qe = [e for e in qe if not filler(e)]
    ae = [e for e in ae if not filler(e)]
    if not qe and not ae:
        continue
    dated = [e['date'] for e in qe + ae if e['date']]
    first = min(dated) if dated else None
    last = max(dated) if dated else None
    for e in qe:
        if e['date'] is None:
            e['date'] = (meet[0] if cur_sec == 'G' and not (qe and qe[0]['date']) and (not first or e is qe[0]) else first) or meet[0]
    for e in ae:
        if e['date'] is None:
            e['date'] = last or (qe[0]['date'] if qe else meet[0])
    allE = qe + ae
    for e in allE:
        e['meeting'] = meeting_of(e['date'])
    allE.sort(key=lambda e: (e['meeting'], 0 if e['kind'] == 'notering' else 1))
    if any(re.search(r'Genomgång\s+av\s+tidplan', e['text']) for e in qe):
        if tidplan is None:
            tidplan = {'section': 'T', 'title': 'Tidplan – lägesrapport per entreprenör', 'kind': 'lägesrapport',
                       'standing': True, 'latestOnly': True, 'status': 'öppen', 'entries': []}
            items.append(tidplan)
        for e in allE:
            e['text'] = re.sub(r'^-?\s*Genomgång\s+av\s+tidplaner/tidplan\s*', '', e['text']).strip()
        tidplan['entries'].extend(allE)
        continue
    title = re.sub(r'~~.*?~~', '', allE[0]['text']).replace('**', '')
    title = re.split(r'(?<=[.?!])\s|\n', title.strip())[0].strip(' .:')
    if len(title) > 78:
        title = title[:75].rsplit(' ', 1)[0] + ' …'
    it = {
        'section': cur_sec, 'title': title or '(utan rubrik)',
        'kind': 'regel' if standing else ('fråga' if '?' in allE[0]['text'][:300] else 'åtgärd'),
        'standing': standing, 'latestOnly': False,
        'status': 'klar' if klart else 'öppen', 'entries': allE,
    }
    if klart:
        it['closedMeeting'] = max(e['meeting'] for e in allE)
    items.append(it)

# Ansvarig: initialer + "kollar/återkommer/undersöker/samordnar"
known = {p['initials'] for p in people}
trade = {'Rör': 'Rör', 'RÖR': 'Rör', 'VS': 'Rör', 'El': 'El', 'EL': 'El', 'Vent': 'Vent', 'VENT': 'Vent', 'Styr': 'Styr', 'STYR': 'Styr', 'Bygg': 'Bygg', 'Mark': 'Mark'}
for it in items:
    if it['status'] != 'öppen' or it['standing']:
        continue
    text = ' '.join(e['text'] for e in it['entries'])
    m = re.findall(r'\b([A-ZÅÄÖ][A-Za-zÅÄÖåäö]{0,3})\s+(?:kollar|återkommer|undersöker|samordnar|meddelar|skickar|ställer|tar fram|bokar)', text)
    for w in reversed(m):
        if w in known or w in trade:
            it['owner'] = trade.get(w, w)
            break

# Nummer per avsnitt i den ordning punkterna skapades
for it in items:
    it['createdMeeting'] = min(e['meeting'] for e in it['entries'])
order = {c: i for i, (c, _, _) in enumerate(SECTIONS)}
for code, _, _ in SECTIONS:
    sec_items = [it for it in items if it['section'] == code]
    sec_items.sort(key=lambda it: it['createdMeeting'])
    for n, it in enumerate(sec_items, 1):
        it['nr'] = f'{code}.{n:03d}'
        it['seq'] = n

pk = 'p' + projnr
os.makedirs(out, exist_ok=True)
docs = {}
docs[f'projects/{pk}'] = {
    'number': projnr, 'name': 'Ljungbyhedsskolan Etapp 2-3', 'client': 'Klippans kommun',
    'place': 'Nybrogatan 9, 264 53 Ljungbyhed (på arbetsplatsen)',
    'siteManager': 'johan-dahl', 'meetingType': 'Samordningsmöte', 'meetingCode': 'SM',
    'keepClosed': 2,
    'approvalText': 'Protokollet anses godkänt om inga skriftliga synpunkter har kommit till protokollföraren inom 5 arbetsdagar.',
    'sections': [{'code': c, 'title': t} for c, t, _ in SECTIONS],
    'created': '2025-08-06',
}
for i, m in enumerate(meet, 1):
    y, mo, dd = ymd(m)
    md = {'project': pk, 'number': i, 'date': f'{y:04d}-{mo:02d}-{dd:02d}', 'time': '10:00',
          'place': 'Platskontoret', 'writtenBy': 'johan-dahl', 'status': 'fastställt', 'attendance': {}}
    if i == len(meet):
        md['status'] = 'utkast'
        md['attendance'] = {f'{pk}--{p["id"]}': ('närvarande' if p['_present'] else ('frånvarande' if p['active'] else 'ingen'))
                            for p in people}
    docs[f'meetings/{pk}-m{i:03d}'] = md
for p in people:
    pres = p.pop('_present')
    docs[f'people/{pk}--{p["id"]}'] = {**{k: v for k, v in p.items() if k != 'id'}, 'project': pk}
# närvaro SM27 (fet = närvarande)
present_ids = set()
for p in people:
    pass
for i, it in enumerate(items):
    e2 = []
    for e in it['entries']:
        e2.append({'m': f'{pk}-m{e["meeting"]:03d}', 'kind': e['kind'], 'text': e['text'], 'date': e['date']})
    body = {'project': pk, 'nr': it['nr'], 'seq': it['seq'], 'section': it['section'], 'title': it['title'],
            'kind': it['kind'], 'standing': it['standing'], 'latestOnly': it['latestOnly'], 'status': it['status'],
            'owner': it.get('owner', ''), 'due': '', 'created': f'{pk}-m{it["createdMeeting"]:03d}',
            'closed': f'{pk}-m{it["closedMeeting"]:03d}' if it.get('closedMeeting') else '', 'entries': e2}
    docs[f'items/{pk}-{it["nr"].replace(".", "")}'] = body
json.dump({'meetings': meet, 'people': len(people), 'items': len(items)}, sys.stdout, ensure_ascii=False)
print()
with open(os.path.join(out, 'seed.json'), 'w') as f:
    json.dump(docs, f, ensure_ascii=False, indent=1)
# En fil per dokument, för inläsning i delad databas (ArtifactData batch med file_path)
dd = os.path.join(out, 'docs'); os.makedirs(dd, exist_ok=True)
for path, body in docs.items():
    coll, did = path.split('/')
    with open(os.path.join(dd, coll + '__' + did + '.json'), 'w') as f:
        json.dump(body, f, ensure_ascii=False)
