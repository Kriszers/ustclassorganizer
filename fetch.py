"""UST ClassFinder subject-by-subject fetcher. Site markup may change over time."""
import re
import time
import requests
import pandas as pd
from bs4 import BeautifulSoup
from .data import normalize

BASE = 'https://classes.aws.stthomas.edu'
TERMS = {'J-Term':10, 'Spring':20, 'Summer':30, 'Fall':40}


def clean(s):
    return re.sub(r'\s+', ' ', str(s or '')).strip()


def get_subjects(session, year, term, school='ALL', level='ALL'):
    response = session.get(BASE + '/json/getSubjectList', params={'year':year,'term':term,'schoolCode':school,'levelCode':level,'hideInvisible':'true'}, timeout=35)
    response.raise_for_status()
    data = response.json()
    if isinstance(data, dict):
        data = data.get('subjects', data.get('data', []))
    codes = []
    for item in data:
        if isinstance(item, str):
            codes.append(item.strip())
            continue
        if not isinstance(item, dict):
            continue
        value = next((str(item[k]).strip() for k in ('subjectCode','code','subject','displaySubjectCode') if item.get(k)), '')
        if not value:
            m = re.search(r'\(([A-Z0-9]+)\)\s*$', str(item.get('name', item.get('subjectName',''))))
            value = m[1] if m else ''
        if value:
            codes.append(value)
    return sorted(set(codes))


def fetch_html(session, year, term, subject, school='ALL', level='ALL'):
    base = {'year':year,'term':term,'schoolCode':school,'levelCode':level}
    for extra in ({'selectedSubjects':subject}, {'displaySubjectCode':subject}, {'selectedSubjects':subject,'displaySubjectCode':subject}, {'displaySubjectcode':subject}, {'subject':subject}):
        r = session.get(BASE + '/index.htm', params={**base, **extra}, timeout=35)
        r.raise_for_status()
        if 'CRN:' in r.text:
            return r.text
    return ''


def course_cards(soup):
    cards = soup.select('div.courseListing, div.course-listing, section.course, li.course')
    if cards:
        return cards
    found, seen = [], set()
    for span in soup.find_all('span', class_=re.compile(r'course(Open|Waitlist|Closed)')):
        node = span
        for _ in range(7):
            node = node.parent
            if not node:
                break
            if re.search(r'\bCRN\s*:\s*\d+', node.get_text(' ', strip=True), re.I):
                if id(node) not in seen:
                    found.append(node)
                    seen.add(id(node))
                break
    return found


def calendar_text(card):
    cal = card.select_one('table.courseCalendar')
    if cal is None:
        return 'NA'
    dates = ''
    for td in cal.select('td'):
        m = re.search(r'Dates:\s*(\d{2}/\d{2}\s*-\s*\d{2}/\d{2})', clean(td.get_text(' ')))
        if m:
            dates = m[1]
            break
    days = [clean(td.get_text(' ')) for td in cal.select('tr td.day')]
    if len(days) != 7:
        days = ['M','T','W','Th','F','Sa','Su']
    entries = []
    row = next((tr for tr in cal.select('tr') if tr.select_one('td.time, td.noTime')), None)
    if row:
        for day, td in zip(days, row.select('td')):
            if 'time' not in td.get('class', []):
                continue
            for p in td.find_all('p'):
                date_match = re.search(r'(\d{2}/\d{2})\s*:', clean(p.get_text(' ')))
                loc = p.select_one('.calendarLocation')
                location = clean(loc.get_text(' ')) if loc else 'NA'
                times = re.findall(r'\d{1,2}:\d{2}\s*[ap]m', p.get_text(' ', strip=True), re.I)
                if len(times) >= 2:
                    entries.append(f'{day} {date_match[1] + " " if date_match else ""}{times[0]}-{times[1]} @ {location}')
    note = cal.select_one('td.asynchNotice')
    parts = [dates, *entries, clean(note.get_text(' ')) if note else '']
    return '; '.join(x for x in parts if x) or 'NA'


def parse_subject_page(html):
    soup = BeautifulSoup(html, 'lxml')
    rows = []
    for card in course_cards(soup):
        title = card.find('span', class_=re.compile(r'course(Open|Waitlist|Closed)'))
        course_id = clean(title.get_text(' ')) if title else 'NA'
        sub = card.select_one('p.courseInfoHighlight a')
        subject = clean(sub.get_text(' ')) if sub else (course_id.split('-')[0] if '-' in course_id else 'NA')
        crn_match = re.search(r'\bCRN\s*:\s*(\d+)', card.get_text(' ', strip=True), re.I)
        location = card.select_one('span.locationHover')
        room = clean(location.get_text(' ')) if location else ''
        if not room and location:
            room = re.sub(r'^.*?Room\s+', '', clean(location.get('title','')), flags=re.I)
        rows.append({'_crn':crn_match[1] if crn_match else '', 'room':room or 'NA','subject':subject,'course_id':course_id,'course_time':calendar_text(card)})
    return rows


def fetch_courses(year=2027, term=20, school='ALL', level='ALL', progress=None, delay=0.15):
    session = requests.Session()
    session.headers.update({'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/123.0 Safari/537.36','Referer':BASE + '/index.htm'})
    subjects = get_subjects(session, year, term, school, level)
    if not subjects:
        raise RuntimeError('ClassFinder returned no subjects. Check semester availability or the website.')
    rows, failed = [], []
    for i, subject in enumerate(subjects, 1):
        try:
            html = fetch_html(session, year, term, subject, school, level)
            if not html:
                failed.append(subject)
            else:
                parsed = parse_subject_page(html)
                if not parsed:
                    failed.append(subject)
                rows.extend(parsed)
        except requests.RequestException:
            failed.append(subject)
        if progress:
            progress(i, len(subjects), subject)
        time.sleep(delay)
    if not rows:
        raise RuntimeError('No course sections could be parsed. ClassFinder may have changed its HTML.')
    df = pd.DataFrame(rows)
    df = df[df['_crn'].ne('')].drop_duplicates(subset=['_crn']).drop(columns=['_crn'])
    if df.empty:
        raise RuntimeError('No valid CRNs found in ClassFinder response.')
    return normalize(df).sort_values(['subject','course_id']).reset_index(drop=True), failed
