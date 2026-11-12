import io
import re
import pandas as pd

REQUIRED = ['room', 'subject', 'course_id', 'course_time']
INVALID = {'', 'NA', 'NAN', 'NONE', 'NO ROOM', 'ONLINE', 'IN PERSON', 'TBA'}
DAYS = {'M':'MON', 'T':'TUES', 'W':'WED', 'TH':'THU', 'R':'THU', 'F':'FRI', 'SA':'SAT', 'SU':'SUN', 'MON':'MON', 'TUE':'TUES', 'TUES':'TUES', 'WED':'WED', 'THU':'THU', 'FRI':'FRI', 'SAT':'SAT', 'SUN':'SUN'}


def normalize(df):
    df = df.copy()
    df.columns = [str(c).strip().lower().replace(' ', '_') for c in df.columns]
    missing = [c for c in REQUIRED if c not in df.columns]
    if missing:
        raise ValueError('Missing required columns: ' + ', '.join(missing) + '. Required: ' + ', '.join(REQUIRED))
    for c in REQUIRED:
        df[c] = df[c].fillna('').astype(str).str.strip()
    df['subject'] = df['subject'].str.upper()
    return df


def read_upload(upload):
    raw = upload.getvalue()
    if upload.name.lower().endswith('.csv'):
        try:
            df = pd.read_csv(io.BytesIO(raw), dtype=str, encoding='utf-8-sig')
        except UnicodeDecodeError:
            df = pd.read_csv(io.BytesIO(raw), dtype=str, encoding='cp1252')
    elif upload.name.lower().endswith(('.xlsx', '.xls')):
        df = pd.read_excel(io.BytesIO(raw), dtype=str)
    else:
        raise ValueError('Please upload CSV, XLSX, or XLS.')
    return normalize(df)


def split_rooms(value):
    return [p.strip().upper() for p in re.split(r'\s*[,;/]\s*', str(value)) if p.strip().upper() not in INVALID]


def clock_minutes(s):
    s = str(s).strip().lower().replace(' ', '')
    m = re.fullmatch(r'(\d{1,2}):(\d{2})(am|pm)', s)
    if not m:
        return None
    h, minute, ap = int(m[1]), int(m[2]), m[3]
    if not 1 <= h <= 12 or minute > 59:
        return None
    return (h % 12 + (12 if ap == 'pm' else 0)) * 60 + minute


def display_time(minutes):
    h, m = divmod(minutes, 60)
    return f'{(h % 12) or 12}:{m:02d} {"PM" if h >= 12 else "AM"}'


def parse_meetings(course_time):
    """Parse ClassFinder format: '01/31 - 05/20; M 5:30pm-8:30pm @ SCC 123; Th 02/15 9:00am-10:00am @ SCC 123'."""
    result = []
    for part in str(course_time).split(';'):
        part = part.strip()
        match = re.search(r'\b(Mon|Tues|Tue|Wed|Thu|Th|Fri|Sat|Sun|Sa|Su|M|T|W|R|F)\s+(?:(\d{1,2}/\d{1,2})\s+)?(\d{1,2}:\d{2}\s*[ap]m)\s*-\s*(\d{1,2}:\d{2}\s*[ap]m)\s*@\s*(.+)$', part, re.I)
        if not match:
            continue
        day = DAYS.get(match[1].upper())
        start, end = clock_minutes(match[3]), clock_minutes(match[4])
        if day and start is not None and end is not None and end > start:
            for room in split_rooms(match[5]):
                result.append({'day':day, 'date':match[2], 'start':start, 'end':end, 'room':room})
    return result


def events(df, prefix=None, seed_subjects=None):
    df = normalize(df)
    if seed_subjects:
        seeds = df[df.subject.isin(seed_subjects)]
        rooms = {r for value in seeds.room for r in split_rooms(value) if not prefix or r.startswith(prefix)}
        for _, row in seeds.iterrows():
            for meeting in parse_meetings(row.course_time):
                if not prefix or meeting['room'].startswith(prefix):
                    rooms.add(meeting['room'])
    else:
        rooms = None
    result = []
    for _, row in df.iterrows():
        course = f'{row.subject} {row.course_id}'.strip()
        for meeting in parse_meetings(row.course_time):
            if prefix and not meeting['room'].startswith(prefix):
                continue
            if rooms is not None and meeting['room'] not in rooms:
                continue
            result.append({**meeting, 'course':course})
    return sorted(result, key=lambda e: (e['room'], e['day'], e['start'], e['course']))
