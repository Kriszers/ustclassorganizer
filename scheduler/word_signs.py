"""OWS/OSS/FDC Word signs based on the user's original Colab sign generator.

Adaptations: input is a DataFrame (no Colab upload), outputs are BytesIO-backed
separate Word files by building for the Streamlit download workflow.
"""
from collections import defaultdict
from io import BytesIO
import re
from datetime import datetime
from docx import Document
from docx.shared import Pt, Inches
from .data import normalize, split_rooms, parse_meetings, display_time

BUILDINGS = ('OWS', 'OSS', 'FDC')
TARGET_SUBJECTS = {'ETLS', 'SEIS'}
DAY_ORDER = {'MON':1, 'TUES':2, 'THU':3, 'WED':4, 'FRI':5, 'SAT':6, 'SUN':7}
DAY_CODES = {'MON':'M','TUES':'T','WED':'W','THU':'TH','FRI':'F','SAT':'SA','SUN':'SU'}
DAY_DISPLAY = {'M':'M','T':'T','W':'W','TH':'TR','F':'F','SA':'SA','SU':'SU'}
HIDDEN_COMMON_RANGES = {'09/09 to 12/21','09/09 to 12/22'}


def format_day_pattern(days):
    """Match the original script's day-pattern display, including T/TH => TR."""
    order = {'M':1,'T':2,'TH':3,'W':4,'F':5,'SA':6,'SU':7}
    ordered = sorted(set(days),key=lambda d:order.get(d,99))
    if ordered == ['M','W','F']:
        return 'MWF'
    if ordered == ['M','W']:
        return 'MW'
    if ordered in (['T','TH'],['T'],['TH']):
        return 'TR'
    return ''.join(DAY_DISPLAY.get(d,d) for d in ordered)


def format_day_for_evening(day):
    return 'TR' if day == 'TH' else day


def extract_date_range(raw):
    match = re.search(r'(\d{2}/\d{2})\s*-\s*(\d{2}/\d{2})',str(raw))
    return f'{match.group(1)} to {match.group(2)}' if match else None


def format_time(mins):
    return display_time(mins)


def raw_room_selection(df, building):
    """Find rooms named in SEIS/ETLS course rows, then include all subjects."""
    seeds = df[df['subject'].isin(TARGET_SUBJECTS)]
    return sorted({room for text in seeds['room'] for room in split_rooms(text)
                   if room.startswith(building)})


def build_room_schedule_sections(room, rows):
    """Reconstruct the original grouped plain-text sign layout."""
    daytime_groups = defaultdict(list)
    evening_groups = defaultdict(list)
    special_lines = []
    for _, row in rows.iterrows():
        course = f"{row['subject']} {row['course_id']}"
        date_range = extract_date_range(row['course_time'])
        meetings = [m for m in parse_meetings(row['course_time']) if m['room'] == room]
        regular_by_time = defaultdict(list)
        for m in meetings:
            day = DAY_CODES.get(m['day'],m['day'])
            if m['date']:
                special_lines.append({
                    'line': f"{format_day_for_evening(day)} {m['date']} {format_time(m['start'])} - {format_time(m['end'])}    {course}",
                    'sort': (m['date'],DAY_ORDER.get(m['day'],99),m['start'],course)})
            else:
                regular_by_time[(m['start'],m['end'])].append(day)
        for (start_m,end_m),days in regular_by_time.items():
            start,end = format_time(start_m),format_time(end_m)
            pattern = format_day_pattern(days)
            if start_m >= 17*60:
                evening_key = (start_m,end_m)
                for day in sorted(days,key=lambda d:{'M':1,'T':2,'TH':3,'W':4,'F':5,'SA':6,'SU':7}.get(d,99)):
                    evening_groups[evening_key].append({
                        'line': f"{format_day_for_evening(day)}    {start} - {end}    {course}",
                        'sort': ({'M':1,'T':2,'TH':3,'W':4,'F':5,'SA':6,'SU':7}.get(day,99),course)})
            else:
                line = f'{pattern} {start} - {end}    {course}'
                if date_range and date_range not in HIDDEN_COMMON_RANGES:
                    line = f'{date_range}    {line}'
                daytime_groups[pattern].append({'line':line,'sort':(start_m,course)})
    sections = []
    preferred = ['MWF','MW','TR','F','SA']
    for pattern in preferred + sorted(k for k in daytime_groups if k not in preferred):
        if pattern in daytime_groups:
            sections.append([pattern]+[x['line'] for x in sorted(daytime_groups[pattern],key=lambda x:x['sort'])])
    for key in sorted(evening_groups):
        start_m,end_m=key
        sections.append([f'EVENINGS {format_time(start_m)} - {format_time(end_m)}'] +
                        [x['line'] for x in sorted(evening_groups[key],key=lambda x:x['sort'])])
    if special_lines:
        sections.append(['SPECIAL DATES'] + [x['line'] for x in sorted(special_lines,key=lambda x:x['sort'])])
    return sections


def add_text_line(doc, text, bold=False, font_size=11, space_after=0):
    para = doc.add_paragraph()
    para.paragraph_format.space_after = Pt(space_after)
    para.paragraph_format.space_before = Pt(0)
    para.paragraph_format.line_spacing = 1.0
    run = para.add_run(text)
    run.font.name = 'Arial'
    run.font.size = Pt(font_size)
    run.bold = bool(bold)
    return para


def add_blank_line(doc):
    para = doc.add_paragraph()
    para.paragraph_format.space_after = Pt(0)
    para.paragraph_format.space_before = Pt(0)
    para.paragraph_format.line_spacing = 1.0


def make_word_signs(df, term_label='Spring 2027'):
    """Generate one Word file for each building; each eligible room starts a page."""
    df = normalize(df)
    results = {}
    for building in BUILDINGS:
        target_rooms = raw_room_selection(df,building)
        doc = Document()
        page = doc.sections[0]
        page.top_margin = Inches(0.7)
        page.bottom_margin = Inches(0.7)
        page.left_margin = Inches(0.8)
        page.right_margin = Inches(0.8)
        added = 0
        for room in target_rooms:
            rows = df[df['room'].apply(lambda text: room in split_rooms(text))]
            groups = build_room_schedule_sections(room, rows)
            if not groups:
                continue
            if added:
                doc.add_page_break()
            added += 1
            add_text_line(doc,room,bold=True,font_size=12,space_after=6)
            for group_index, group in enumerate(groups):
                for line_index,line in enumerate(group):
                    add_text_line(doc,line,bold=(line_index==0),font_size=11)
                if group_index < len(groups)-1:
                    add_blank_line(doc)
        if not added:
            add_text_line(doc,f'{building} — {term_label}',bold=True,font_size=12,space_after=6)
            add_text_line(doc,'No eligible classroom meetings found.')
        output = BytesIO()
        doc.save(output)
        results[building] = (output.getvalue(),added)
    return results
