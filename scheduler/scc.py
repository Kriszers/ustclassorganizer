"""SCC printable schedule: one easy-to-read Excel worksheet per classroom."""
from collections import defaultdict
from io import BytesIO
import re

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins

from .data import display_time, events

DAYS = ('MON', 'TUES', 'WED', 'THU', 'FRI', 'SAT')
PALETTE = ('DDEBF7', 'DFF0E1', 'FFF2CE', 'FCE3D7', 'EAE3F7', 'D9EEF1', 'F9E0E8', 'E7EBDD')
INK = '183153'
HEADER = '183153'
LIGHT = 'EAF0F7'
GRID = 'CBD5DF'


def safe_title(room, taken):
    base = re.sub(r'[\\/*?:\[\]]', '_', room).strip(" '")[:31] or 'Room'
    title = base
    index = 2
    while title.casefold() in taken:
        suffix = f' ({index})'
        title = base[:31 - len(suffix)] + suffix
        index += 1
    taken.add(title.casefold())
    return title


def event_label(e):
    course = re.sub(r'(?<=\d)-(?=\d)', ' - ', e['course'])
    time = f"{display_time(e['start'])} – {display_time(e['end'])}"
    return f"{course}\n{time}" + (f"\nDate: {e['date']}" if e.get('date') else '')


def make_scc(df, term_label='Spring 2027'):
    by_room = defaultdict(list)
    for item in events(df, prefix='SCC', seed_subjects={'ETLS', 'SEIS', 'ENGR'}):
        if item['day'] in DAYS:
            by_room[item['room']].append(item)
    wb = Workbook()
    blank = wb.active
    used = set()
    stroke = Side(style='hair', color=GRID)
    for room, meetings in sorted(by_room.items()):
        ws = wb.create_sheet(safe_title(room, used))
        ws.sheet_view.showGridLines = False
        ws.sheet_view.zoomScale = 85
        ws.freeze_panes = 'B5'
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.sheet_properties.outlinePr.summaryBelow = True
        ws.page_setup.orientation = 'landscape'
        ws.page_setup.paperSize = ws.PAPERSIZE_A3
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 1
        ws.print_options.horizontalCentered = True
        ws.page_margins = PageMargins(left=.20, right=.20, top=.35, bottom=.35, header=.15, footer=.15)
        ws.oddFooter.center.text = f'{room}  |  {term_label}  |  Page &P of &N'
        ws.oddFooter.center.size = 9
        ws.column_dimensions['A'].width = 19
        for col in range(2, 8):
            ws.column_dimensions[get_column_letter(col)].width = 29

        ws.merge_cells('A1:G2')
        title = ws['A1']
        title.value = f'{room}  |  CLASSROOM SCHEDULE'
        title.font = Font(name='Aptos Display', size=19, bold=True, color='FFFFFF')
        title.fill = PatternFill('solid', fgColor=HEADER)
        title.alignment = Alignment(vertical='center', indent=1)
        for row in (1, 2):
            ws.row_dimensions[row].height = 24
        ws.merge_cells('A3:G3')
        sub = ws['A3']
        sub.value = f'{term_label}    •    Monday–Saturday    •    30-minute time grid'
        sub.fill = PatternFill('solid', fgColor=LIGHT)
        sub.font = Font(name='Aptos', size=11, bold=True, color=INK)
        sub.alignment = Alignment(vertical='center', indent=2)
        ws.row_dimensions[3].height = 27

        for col, text in enumerate(('TIME', *DAYS), 1):
            cell = ws.cell(4, col, text)
            cell.fill = PatternFill('solid', fgColor='36577F')
            cell.font = Font(name='Aptos', size=11, color='FFFFFF', bold=True)
            cell.alignment = Alignment(horizontal='center', vertical='center')
        ws.row_dimensions[4].height = 28

        # A fixed grid across room tabs makes comparisons and printing easier.
        # Match original SCC schedule: 8:00 AM through the 9:30 PM slot.
        # Expand only when the live data includes meetings outside that grid.
        earliest = min(8 * 60, min(e['start'] for e in meetings) // 30 * 30)
        latest = max(22 * 60, max((e['end'] + 29) // 30 * 30 for e in meetings))
        slots = list(range(earliest, latest, 30))
        first = 5
        for index, t in enumerate(slots):
            r = first + index
            ws.row_dimensions[r].height = 27
            timecell = ws.cell(r, 1, f'{display_time(t)} – {display_time(t + 30)}')
            timecell.font = Font(name='Aptos', size=9, bold=(t % 60 == 0), color=INK)
            timecell.fill = PatternFill('solid', fgColor=LIGHT if index % 2 == 0 else 'F7F9FC')
            timecell.alignment = Alignment(horizontal='center', vertical='center')
            for col in range(1, 8):
                cell = ws.cell(r, col)
                cell.border = Border(bottom=stroke, right=stroke)
                if col > 1:
                    cell.fill = PatternFill('solid', fgColor='FFFFFF' if index % 2 else 'FAFBFD')
                    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

        # Partition each day into intervals with the same concurrent courses.
        # Merge a block only when its content is identical throughout the interval.
        for day_idx, day in enumerate(DAYS, start=2):
            day_events = [e for e in meetings if e['day'] == day]
            signatures = []
            for minute in slots:
                active = [e for e in day_events if e['start'] < minute + 30 and e['end'] > minute]
                signature = tuple(sorted({(e['course'], e['start'], e['end'], e.get('date') or '') for e in active}))
                signatures.append(signature)
            start = 0
            while start < len(slots):
                signature = signatures[start]
                end = start + 1
                while end < len(slots) and signatures[end] == signature:
                    end += 1
                if signature:
                    top = first + start
                    bottom = first + end - 1
                    if bottom > top:
                        ws.merge_cells(start_row=top, start_column=day_idx, end_row=bottom, end_column=day_idx)
                    cell = ws.cell(top, day_idx)
                    labels = []
                    for course, beginning, finish, date in signature:
                        labels.append(event_label({'course': course, 'start': beginning, 'end': finish, 'date': date}))
                    cell.value = '\n──────────\n'.join(labels)
                    shade = sum(ord(ch) for ch in signature[0][0]) % len(PALETTE)
                    cell.fill = PatternFill('solid', fgColor=PALETTE[shade] if len(signature) == 1 else 'F6D6D6')
                    cell.font = Font(name='Aptos', size=10 if len(signature) == 1 else 9, color=INK, bold=True)
                    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True, indent=0)
                    cell.border = Border(left=Side(style='medium', color='6682A4'),
                                         right=stroke, top=stroke, bottom=stroke)
                    # More room for two simultaneous classes, including very short overlaps.
                    if len(signature) > 1:
                        per_row = max(27, (63 * len(signature)) / (end - start))
                        for rr in range(top, bottom + 1):
                            ws.row_dimensions[rr].height = max(ws.row_dimensions[rr].height, per_row)
                start = end

        ws.print_title_rows = '1:4'
        ws.print_area = f'A1:G{first + len(slots) - 1}'
    if by_room:
        wb.remove(blank)
        wb.active = 0
    else:
        blank.title = 'No SCC Rooms'
        blank['A1'] = 'No SCC classroom meetings found in the supplied data.'
    output = BytesIO()
    wb.save(output)
    return output.getvalue(), len(by_room)
