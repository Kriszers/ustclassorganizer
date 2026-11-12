"""OWS / OSS / FDC classroom signs matching the supplied Fall 2026 sample.

Compatible with the existing app.py contract: make_word_signs(df, term_label)
returns {building: (docx_bytes, classroom_count)}. SCC generation is untouched.
"""
from collections import defaultdict
from io import BytesIO
import re
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_TAB_ALIGNMENT, WD_TAB_LEADER
from .data import normalize, split_rooms, parse_meetings, display_time

BUILDINGS = ('OWS', 'OSS', 'FDC')
TARGET_SUBJECTS = {'ETLS', 'SEIS'}
DAY_ORDER = {'M': 1, 'T': 2, 'TH': 3, 'W': 4, 'F': 5, 'SA': 6, 'SU': 7}
DAY_CODES = {'MON': 'M', 'TUES': 'T', 'WED': 'W', 'THU': 'TH', 'FRI': 'F', 'SAT': 'SA', 'SUN': 'SU'}
DAY_DISPLAY = {'M':'M', 'T':'T', 'W':'W', 'TH':'TR', 'F':'F', 'SA':'SA', 'SU':'SU'}
HIDDEN_COMMON_RANGES = {'09/09 to 12/21', '09/09 to 12/22'}


def format_day_pattern(days):
    ordered = sorted(set(days), key=lambda d: DAY_ORDER.get(d, 99))
    if ordered == ['M', 'W', 'F']:
        return 'MWF'
    if ordered == ['M', 'W']:
        return 'MW'
    if ordered in (['T', 'TH'], ['T'], ['TH']):
        return 'TR'
    return ''.join(DAY_DISPLAY.get(d, d) for d in ordered)


def format_day_for_evening(day):
    return 'TR' if day == 'TH' else day


def extract_date_range(raw):
    match = re.search(r'(\d{2}/\d{2})\s*-\s*(\d{2}/\d{2})', str(raw))
    return f'{match.group(1)} to {match.group(2)}' if match else None


def raw_room_selection(df, building):
    seeds = df[df['subject'].isin(TARGET_SUBJECTS)]
    return sorted({room for value in seeds['room'] for room in split_rooms(value)
                   if room.startswith(building)})


def build_room_schedule_sections(room, rows):
    """Return groups, each with a heading and rows: (days, times, course, date_range)."""
    daytime = defaultdict(list)
    evening = defaultdict(list)
    special = []
    for _, row in rows.iterrows():
        course = f"{row['subject']} {row['course_id']}"
        date_range = extract_date_range(row['course_time'])
        meetings = [m for m in parse_meetings(row['course_time']) if m['room'] == room]
        regular_by_time = defaultdict(list)
        for m in meetings:
            day = DAY_CODES.get(m['day'], m['day'])
            if m['date']:
                special.append((m['date'], DAY_ORDER.get(day, 99), m['start'],
                                f"{format_day_for_evening(day)} {m['date']}",
                                f"{display_time(m['start'])} - {display_time(m['end'])}", course))
            else:
                regular_by_time[(m['start'], m['end'])].append(day)
        for (start_m, end_m), days in regular_by_time.items():
            times = f'{display_time(start_m)} - {display_time(end_m)}'
            if start_m >= 17 * 60:
                for day in sorted(days, key=lambda d: DAY_ORDER.get(d, 99)):
                    evening[(start_m, end_m)].append((DAY_ORDER.get(day, 99), course,
                                                       format_day_for_evening(day), times, course))
            else:
                pattern = format_day_pattern(days)
                shown_range = date_range if date_range and date_range not in HIDDEN_COMMON_RANGES else ''
                daytime[pattern].append((start_m, course, pattern, times, course, shown_range))
    sections = []
    preferred = ['MWF', 'MW', 'TR', 'F', 'SA']
    for pattern in preferred + sorted(k for k in daytime if k not in preferred):
        if pattern in daytime:
            entries = sorted(daytime[pattern], key=lambda x: x[:2])
            sections.append((pattern, [(d, t, c, dr) for _, _, d, t, c, dr in entries]))
    for start_end in sorted(evening):
        start_m, end_m = start_end
        heading = f'EVENINGS {display_time(start_m)} - {display_time(end_m)}'
        entries = sorted(evening[start_end], key=lambda x: x[:2])
        sections.append((heading, [(d, t, c, '') for _, _, d, t, c in entries]))
    if special:
        entries = sorted(special, key=lambda x: x[:3])
        sections.append(('SPECIAL DATES', [(d, t, c, '') for _, _, _, d, t, c in entries]))
    return sections


def configure_document(doc):
    section = doc.sections[0]
    # The supplied reference is compact and uses ample surrounding whitespace.
    section.top_margin = Inches(0.70)
    section.bottom_margin = Inches(0.70)
    section.left_margin = Inches(0.80)
    section.right_margin = Inches(0.80)
    normal = doc.styles['Normal']
    normal.font.name = 'Arial'
    normal.font.size = Pt(9)
    normal.paragraph_format.space_after = Pt(0)
    normal.paragraph_format.space_before = Pt(0)
    normal.paragraph_format.line_spacing = 1.0


def add_line(doc, text='', bold=False, after=0):
    para = doc.add_paragraph(style='Normal')
    para.paragraph_format.space_after = Pt(after)
    run = para.add_run(text)
    run.font.name = 'Arial'
    run.font.size = Pt(9)
    run.bold = bold
    return para


def add_meeting(doc, day, times, course):
    para = add_line(doc)
    # Tab stops keep all three fields in reliable, readable columns.
    tabs = para.paragraph_format.tab_stops
    tabs.add_tab_stop(Inches(0.43), WD_TAB_ALIGNMENT.LEFT, WD_TAB_LEADER.SPACES)
    tabs.add_tab_stop(Inches(1.75), WD_TAB_ALIGNMENT.LEFT, WD_TAB_LEADER.SPACES)
    para.add_run(day + '\t' + times + '\t' + course)
    return para


def make_word_signs(df, term_label='Spring 2027'):
    df = normalize(df)
    results = {}
    for building in BUILDINGS:
        doc = Document()
        configure_document(doc)
        added = 0
        for room in raw_room_selection(df, building):
            rows = df[df['room'].apply(lambda value: room in split_rooms(value))]
            sections = build_room_schedule_sections(room, rows)
            if not sections:
                continue
            if added:
                doc.add_page_break()
            added += 1
            add_line(doc, room, bold=True, after=5)
            for section_index, (heading, items) in enumerate(sections):
                add_line(doc, heading, bold=True)
                last_date = None
                for day, times, course, date_range in items:
                    if date_range and date_range != last_date:
                        # In the reference, uncommon course-date ranges occupy a separate line.
                        add_line(doc, date_range)
                    add_meeting(doc, day, times, course)
                    last_date = date_range
                if section_index < len(sections)-1:
                    add_line(doc)
        if not added:
            add_line(doc, f'{building} — {term_label}', bold=True, after=5)
            add_line(doc, 'No eligible classroom meetings found.')
        buffer = BytesIO()
        doc.save(buffer)
        results[building] = (buffer.getvalue(), added)
    return results
