#!/usr/bin/env python3
"""
Byggir staff_overview.json úr Vinnuyfirlit_Margir.xlsx (tímaskráningar-útflutningur).
Hvert blað = einn starfsmaður (stundum fleiri en eitt blað per mann ef deildaskipti
áttu sér stað á tímabilinu - þau eru sameinuð).

Aldur er reiknaður út frá kennitölu en kennitalan sjálf er ALDREI vistuð né birt.
"""
import openpyxl, re, json, datetime, sys

NAME_ALIASES = {
    'Nanna Katrín Snorradóttir': 'Nanna Kristín Snorradóttir',
    'Magnús Orri Sigmundsson': 'Magnús Orri Sigumundsson',
}

HEADER_RE = re.compile(r'^(.*?)\s*,,\s*kt\.\s*(\d{10})\s*,\s*(\d{2}\.\d{2}\.\d{4})\s*-\s*(\d{2}\.\d{2}\.\d{4})')
DAY_RE = re.compile(r'^(\d{2})\.(\d{2})\s*([a-záðéíóúýþæö]?)$', re.IGNORECASE)

# Aðgerð-kóðar (staðfest af notanda)
SICK_CODES = {
    'VFL': 'stadfest',      # Staðfest veikindi (tilkynnt Heilsuvernd + vaktstjóra)
    'VEB': 'barn',          # Veikindi barns (tilkynnt Heilsuvernd + vaktstjóra)
    'VEIKOST': 'ostadfest',  # Óstaðfest veikindi (beðið eftir Heilsuvernd)
}
# íslenska Ó í VEIKÓST veldur usmiku - notum bæði stafsetningar
def norm_code(c):
    if not c:
        return c
    return str(c).strip().upper().replace('Ó', 'O')

def build_name_to_initials(staff_xlsx):
    wb = openpyxl.load_workbook(staff_xlsx, data_only=True)
    ws = wb['Staff']
    m = {}
    for r in range(2, ws.max_row + 1):
        ini = ws.cell(row=r, column=1).value
        full = ws.cell(row=r, column=2).value
        if ini and full:
            m[str(full).strip()] = str(ini).strip()
    return m

def kennitala_age(kt, reference_date):
    dd, mm, yy = int(kt[0:2]), int(kt[2:4]), int(kt[4:6])
    for century in (1900, 2000):
        try:
            byear = century + yy
            bdate = datetime.date(byear, mm, dd)
        except ValueError:
            continue
        age = reference_date.year - bdate.year - ((reference_date.month, reference_date.day) < (bdate.month, bdate.day))
        if 14 <= age <= 80:
            return age, bdate
    return None, None

def parse_time_range(s):
    if not s or '-' not in str(s):
        return None, None
    a, b = str(s).split('-', 1)
    def to_min(t):
        t = t.strip()
        if ':' not in t:
            return None
        h, m = t.split(':')
        return int(h) * 60 + int(m)
    return to_min(a), to_min(b)

def parse_sheet(ws, year_hint_start, year_hint_end):
    """Skilar lista af dag-fikserslum úr einu blaði (frá röð 4 og niður úr)."""
    rows = []
    r = 4
    max_r = ws.max_row
    while r <= max_r:
        dag_raw = ws.cell(row=r, column=2).value
        if dag_raw:
            m = DAY_RE.match(str(dag_raw).strip())
            if m:
                dd, mm, wd = m.groups()
                # ákveða ár: ef mánuður passar við upphafsmánuð tímabilsins, nota það ár, annars lokaár
                month = int(mm)
                year = year_hint_start.year if month == year_hint_start.month else year_hint_end.year
                try:
                    date = datetime.date(year, month, int(dd))
                except ValueError:
                    r += 1
                    continue
                eining = ws.cell(row=r, column=3).value
                adgerd = ws.cell(row=r, column=4).value
                skipulag = ws.cell(row=r, column=5).value
                vidvera = ws.cell(row=r, column=6).value
                ath = ws.cell(row=r, column=7).value
                rows.append({
                    'date': date.isoformat(),
                    'weekday_letter': wd,
                    'eining': str(eining).strip() if eining else '',
                    'adgerd': str(adgerd).strip() if adgerd else '',
                    'skipulag': str(skipulag).strip() if skipulag else '',
                    'vidvera': str(vidvera).strip() if vidvera else '',
                    'ath': str(ath).strip() if ath else '',
                })
        r += 1
    return rows

def build_overview(vinnuyfirlit_xlsx, staff_xlsx, name_map, phone_map):
    name_to_initials = build_name_to_initials(staff_xlsx)
    wb = openpyxl.load_workbook(vinnuyfirlit_xlsx, data_only=True)

    by_kt = {}  # kennitala -> {fullname, initials, rows: [...], period}
    skipped_sheets = []

    for sheetname in wb.sheetnames:
        ws = wb[sheetname]
        header = ws.cell(row=2, column=2).value
        if not header:
            skipped_sheets.append(sheetname)
            continue
        m = HEADER_RE.match(str(header))
        if not m:
            skipped_sheets.append(sheetname)
            continue
        fullname, kt, d1, d2 = m.groups()
        fullname = fullname.strip()
        lookup_name = NAME_ALIASES.get(fullname, fullname)
        initials = name_to_initials.get(lookup_name)
        if not initials:
            skipped_sheets.append(sheetname + ' (nafn fannst ekki: ' + fullname + ')')
            continue
        start = datetime.datetime.strptime(d1, '%d.%m.%Y').date()
        end = datetime.datetime.strptime(d2, '%d.%m.%Y').date()

        rows = parse_sheet(ws, start, end)
        if not rows:
            continue

        entry = by_kt.setdefault(kt, {
            'fullname': fullname, 'initials': initials,
            'period_start': start.isoformat(), 'period_end': end.isoformat(),
            'rows_by_date': {}
        })
        for row in rows:
            entry['rows_by_date'][row['date']] = row  # síðasta blað "vinnur" ef skörun

    employees = []
    for kt, entry in by_kt.items():
        ini = entry['initials']
        age, _ = kennitala_age(kt, datetime.date.today())
        rows = sorted(entry['rows_by_date'].values(), key=lambda r: r['date'])

        worked_days = 0
        late_count = 0
        late_minutes_total = 0
        sick = {'stadfest': 0, 'barn': 0, 'ostadfest': 0}
        unpaid = 0
        training = 0
        extra = 0
        closed = 0
        other = 0
        depts = {}
        day_records = []

        for row in rows:
            code = norm_code(row['adgerd'])
            dept = row['eining']
            if dept:
                # Hreinsa "Skeifan " forskeyti og geyma stystu skýru mynd
                dept_clean = dept.replace('Skeifan', '').strip() or dept
                depts[dept_clean] = depts.get(dept_clean, 0) + 1

            is_late = False
            late_min = 0
            if code == 'VINNA':
                worked_days += 1
                s_start, s_end = parse_time_range(row['skipulag'])
                v_start, v_end = parse_time_range(row['vidvera'])
                if s_start is not None and v_start is not None and v_start > s_start + 5:
                    is_late = True
                    late_min = v_start - s_start
                    late_count += 1
                    late_minutes_total += late_min
            elif code in SICK_CODES:
                sick[SICK_CODES[code]] += 1
            elif code == 'LAUL':
                unpaid += 1
            elif code == 'NAMSK':
                training += 1
            elif code == 'AUKA':
                extra += 1
            elif code == 'LOKAÐ':
                closed += 1
            elif code:
                other += 1

            day_records.append({
                'date': row['date'], 'adgerd': row['adgerd'], 'dept': dept,
                'skipulag': row['skipulag'], 'vidvera': row['vidvera'],
                'late': is_late, 'lateMin': late_min,
            })

        top_dept = max(depts.items(), key=lambda x: x[1])[0] if depts else ''
        total_sick = sick['stadfest'] + sick['barn'] + sick['ostadfest']

        employees.append({
            'initials': ini,
            'name': name_map.get(ini, ini),
            'phone': phone_map.get(ini, ''),
            'dept': top_dept,
            'age': age,
            'periodStart': entry['period_start'],
            'periodEnd': entry['period_end'],
            'workedDays': worked_days,
            'lateCount': late_count,
            'lateMinutesTotal': late_minutes_total,
            'sick': sick,
            'sickTotal': total_sick,
            'unpaid': unpaid,
            'training': training,
            'extra': extra,
            'closed': closed,
            'other': other,
            'days': day_records,
        })

    employees.sort(key=lambda e: e['name'])
    return employees, skipped_sheets

if __name__ == '__main__':
    vinnu = sys.argv[1] if len(sys.argv) > 1 else 'Vinnuyfirlit_Margir.xlsx'
    staff = sys.argv[2] if len(sys.argv) > 2 else 'Vaktaplan_Skeifan.xlsx'
    namemap_path = sys.argv[3] if len(sys.argv) > 3 else 'name_map.json'
    out = sys.argv[4] if len(sys.argv) > 4 else 'staff_overview.json'
    nm = json.load(open(namemap_path, encoding='utf-8'))
    employees, skipped = build_overview(vinnu, staff, nm['names'], nm['phones'])
    json.dump(employees, open(out, 'w', encoding='utf-8'), ensure_ascii=False, separators=(',', ':'))
    print(f'{len(employees)} starfsmenn skrifaðir í {out}')
    print(f'Sleppt: {len(skipped)} blöðum')
    for s in skipped:
        print('  -', s)
