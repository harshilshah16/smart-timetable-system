# smart-timetable-system

# Timetable backend

This API generates one timetable containing `SY A`, `SY B`, `TY A`, `TY B`, `LY A`, and `LY B` divisions from an Excel workbook.

## Run

```powershell
cd backend
py -m venv .venv
.\.venv\Scripts\Activate.ps1
py -m pip install -r requirements.txt
py anti.py
```

The API is available at `http://localhost:8000`. Interactive API documentation is at `http://localhost:8000/docs`.

## Excel format

Use a sheet named `Courses` with the first row containing these required columns. Download a working template from `GET /api/timetable/sample.xlsx`. CSV uploads use the same header row and do not need a sheet name.

| Division | Batch | Course Type | Course Code | Course Name | Faculty ID | Room | Sessions Per Week | Duration |
| --- | --- | --- | --- | --- | --- | --- | --- | --- |
| SY A | B1 | Lab | L301 | Data Structures Lab | DS-B1 | Lab | 1 | 2 |

Optional columns are `Preferred Day` and `Preferred Period`. `Room` can be `Any`, a named room number, a comma-separated list of room numbers, or `Lab`. A value of `Lab` allocates one available lab from the workbook. Add a `Rooms` sheet to define the lab inventory; no lab names are built into the system:

| Room No | Type |
| --- | --- |
| 502A | Lab |
| 502B | Lab |
| 503A | Lab |
| 503B | Lab |

All four rows can be allocated in the same period to four different batches. For a full rotation, include one row for each lab subject and batch, as shown in the downloadable sample. The four lab subjects rotate across B1-B4 over four days, and each lab occupies the current period plus the next period. `Duration` may be left blank for lab rows because labs default to two periods. The assigned unique room number and batch are printed across the two-period timetable cell, for example `B1 DS / 502A`. Each course row is expanded into the requested number of weekly sessions, repeated sessions of the same subject are separated by at least one teaching period, and faculty and room clashes are checked globally across divisions. An optional `Settings` sheet can contain `Days`, `Periods`, `College`, `Department`, `Academic Year`, and `Semester` rows. By default the periods start at 08:00 and use only these breaks: 09:55-10:05 and 12:00-12:30. The PDF uses these values to create the college-style header and timetable grid; it does not use the example image's subjects or faculty.

## Endpoints

- `GET /health`
- `GET /api/timetable/sample.xlsx` downloads a complete example workbook.
- `POST /api/timetable/generate` with multipart field `file`; accepts `.xlsx`, `.xlsm`, or `.csv` and returns JSON grouped by division.
- `POST /api/timetable/generate.xlsx` with multipart field `file`; downloads the generated timetable as Excel.
- `POST /api/timetable/generate.pdf` with multipart field `file`; downloads the generated timetable as PDF.

The generator prevents clashes for the same division, faculty member, or named room. Sessions that cannot be placed are returned in `conflicts` instead of being silently dropped.
