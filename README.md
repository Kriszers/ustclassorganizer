# UST Class Organizer — Spring 2027

Streamlit website for fetching UST ClassFinder data and creating SCC Excel schedules and OWS/OSS/FDC Word signs.

## Windows setup (Python 3.13 + VS Code)

1. Download the ZIP and **Extract All**. Do not copy files out individually.
2. Open VS Code > File > Open Folder > `ust_classroom_scheduler` (the folder containing `app.py`).
3. Open Terminal > New Terminal (PowerShell) and run:

```powershell
py -3.13 -m venv .venv
.\.venv\Scripts\python.exe -m pip install --upgrade pip
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe -m streamlit run app.py
```

4. Open the local URL shown (usually `http://localhost:8501`).
5. Choose a semester (Spring 2027 is the default) and click **Generate Classroom Schedules**. The website will fetch ClassFinder data, then automatically generate SCC Excel and OWS/OSS/FDC Word files.
6. Download individual schedules or the combined ZIP. If ClassFinder fails, expand **Backup: Upload CSV / Excel (optional)** and generate from a file instead.

### Upload columns

Headers (case-insensitive): `room`, `subject`, `course_id`, `course_time`.

Example `course_time`:

`01/31 - 05/20; M 5:30pm-8:30pm @ SCC 123; W 5:30pm-8:30pm @ SCC 123`

Special-date meeting example: `Th 02/15 9:00am-10:00am @ OWS 200`.

The meeting room after `@` is used for schedule placement. Online/TBA meetings are excluded. SCC uses all subjects meeting in SCC rooms. Word signs discover OWS/OSS/FDC rooms with ETLS or SEIS classes, then include other courses in those same rooms.

### Troubleshooting

- **Null bytes / SyntaxError**: Extract the whole ZIP first. Do not copy code from a formatted chat or open Python files in Word.
- **`streamlit` not recognized**: Use the exact `.venv\Scripts\python.exe -m streamlit run app.py` command above.
- **ClassFinder fetch fails**: The external website may change or block automated requests. Use the backup upload section with a previously exported CSV/Excel file instead.
- **Empty signs**: Verify `room` and `course_time` values and matching room codes; review the data preview.
- **For coworkers**: `localhost` is only accessible on your laptop. Deploy the Streamlit app on an approved server/cloud service to share a URL; check university data/security policy first.

## Testing

```powershell
.\.venv\Scripts\python.exe -m unittest discover -s tests -v
```

## Notes

This is a reconstructed application using the original workflow. Output layout may differ from the legacy Colab scripts. Verify a sample schedule against ClassFinder before using signs in production. The ClassFinder parser depends on UST's current HTML structure and has not been live-verified.

## SCC Excel workbook layout

Each SCC classroom now has its **own worksheet/tab** (for example `SCC 101`). The timetable has Monday–Saturday columns, 30-minute time intervals, full-length merged course blocks, clear color highlighting, freeze panes, and print-ready landscape settings. Overlapping classes are displayed together in a pink block rather than hidden.

To upgrade an existing one-click installation, replace **`scheduler/scc.py`** (and `app.py` if you want the updated website name), then refresh Streamlit.

## Word signs: original grouping style

The OWS, OSS and FDC Word generators now follow the uploaded original sign script: only ETLS/SEIS courses determine eligible rooms, then all subjects are listed for those rooms. Entries are grouped into daytime day-pattern sections, evening time sections and special-date sections. Each room begins a separate page; the three buildings remain separate downloadable documents. SCC generator is unchanged.
