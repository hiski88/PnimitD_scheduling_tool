import calendar
import json
import re
from io import BytesIO
from datetime import date, datetime, time, timedelta, timezone
from pathlib import Path
from typing import Dict, List, Optional, Tuple
from urllib.parse import urlencode

from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

import pandas as pd
import requests
import streamlit as st

try:
    from google_auth_oauthlib.flow import Flow
    from google.oauth2.credentials import Credentials
    from googleapiclient.discovery import build
except Exception:
    Flow = None
    Credentials = None
    build = None


# =========================
# App constants
# =========================

SCOPES = ["https://www.googleapis.com/auth/calendar.readonly"]
DATA_DIR = Path("data")
DATA_DIR.mkdir(exist_ok=True)
OUTPUT_FILE = DATA_DIR / "availability_submissions.jsonl"

st.set_page_config(
    page_title="מע׳ לתכנון תורנויות- פנימית ד׳",
    page_icon="🗓️",
    layout="wide",
)


# =========================
# Date helpers
# =========================

def add_months(year: int, month: int, delta: int) -> Tuple[int, int]:
    """Return year, month after adding delta months."""
    zero_based = (year * 12 + (month - 1)) + delta
    return zero_based // 12, zero_based % 12 + 1


def default_next_month() -> Tuple[int, int]:
    today = date.today()
    return add_months(today.year, today.month, 1)


def month_dates(year: int, month: int) -> List[date]:
    last_day = calendar.monthrange(year, month)[1]
    return [date(year, month, day) for day in range(1, last_day + 1)]


def iso_month_range(year: int, month: int) -> Tuple[str, str]:
    days = month_dates(year, month)
    start_dt = datetime.combine(days[0], time.min).astimezone()
    end_dt = datetime.combine(days[-1] + timedelta(days=1), time.min).astimezone()
    return start_dt.isoformat(), end_dt.isoformat()


def hebrew_weekday(d: date) -> str:
    names = {
        0: "שני",
        1: "שלישי",
        2: "רביעי",
        3: "חמישי",
        4: "שישי",
        5: "שבת",
        6: "ראשון",
    }
    return names[d.weekday()]


def month_title(year: int, month: int) -> str:
    hebrew_months = {
        1: "ינואר",
        2: "פברואר",
        3: "מרץ",
        4: "אפריל",
        5: "מאי",
        6: "יוני",
        7: "יולי",
        8: "אוגוסט",
        9: "ספטמבר",
        10: "אוקטובר",
        11: "נובמבר",
        12: "דצמבר",
    }
    return f"{hebrew_months[month]} {year}"


# =========================
# Holiday provider
# MVP: display-only Jewish/Israeli holidays.
# This is intentionally separate from the scheduling engine.
# =========================

@st.cache_data(ttl=60 * 60 * 24)
def fetch_jewish_holidays_from_hebcal(year: int) -> Dict[str, List[str]]:
    """
    Fetch Jewish holidays for Israel from Hebcal.
    Display-only; user decides what becomes unavailable/vacation.
    """
    url = "https://www.hebcal.com/hebcal"
    params = {
        "v": "1",
        "cfg": "json",
        "maj": "on",
        "min": "on",
        "mod": "on",
        "nx": "on",
        "year": year,
        "month": "x",
        "geo": "none",
        "c": "off",
        "lg": "he",
        "i": "on",
    }

    try:
        res = requests.get(url, params=params, timeout=10)
        res.raise_for_status()
        payload = res.json()
    except Exception:
        return {}

    holidays: Dict[str, List[str]] = {}
    for item in payload.get("items", []):
        category = item.get("category", "")
        if category not in {"holiday", "hebdate", "roshchodesh", "fast"}:
            continue
        iso = item.get("date")
        title = item.get("title")
        if not iso or not title:
            continue
        # Normalize date part only.
        iso_date = iso[:10]
        holidays.setdefault(iso_date, []).append(title)
    return holidays


def holiday_map_for_month(year: int, month: int) -> Dict[str, str]:
    # For January/December month views, fetch current year only is enough for visible dates.
    raw = fetch_jewish_holidays_from_hebcal(year)
    out = {}
    for d in month_dates(year, month):
        values = raw.get(d.isoformat(), [])
        if d.weekday() == 5:
            values = ["שבת"] + values
        out[d.isoformat()] = " | ".join(values)
    return out


# =========================
# Calendar integrations
# =========================

def parse_google_secret_client_config() -> Optional[dict]:
    """
    Expected Streamlit secrets format:

    [google_oauth]
    client_id = "..."
    client_secret = "..."
    redirect_uri = "https://your-app.streamlit.app"
    """
    try:
        cfg = st.secrets["google_oauth"]
        return {
            "web": {
                "client_id": cfg["client_id"],
                "client_secret": cfg["client_secret"],
                "auth_uri": "https://accounts.google.com/o/oauth2/auth",
                "token_uri": "https://oauth2.googleapis.com/token",
                "redirect_uris": [cfg["redirect_uri"]],
            }
        }
    except Exception:
        return None


def google_is_configured() -> bool:
    return Flow is not None and build is not None and parse_google_secret_client_config() is not None


def start_google_oauth() -> Optional[str]:
    client_config = parse_google_secret_client_config()
    if not client_config or Flow is None:
        return None

    redirect_uri = st.secrets["google_oauth"]["redirect_uri"]
    flow = Flow.from_client_config(
        client_config,
        scopes=SCOPES,
        redirect_uri=redirect_uri,
    )
    auth_url, state = flow.authorization_url(
        access_type="offline",
        include_granted_scopes="true",
        prompt="consent",
    )
    st.session_state["google_oauth_state"] = state
    return auth_url


def finish_google_oauth(code: str) -> bool:
    client_config = parse_google_secret_client_config()
    if not client_config or Flow is None:
        return False

    redirect_uri = st.secrets["google_oauth"]["redirect_uri"]
    flow = Flow.from_client_config(
        client_config,
        scopes=SCOPES,
        redirect_uri=redirect_uri,
        state=st.session_state.get("google_oauth_state"),
    )
    flow.fetch_token(code=code)
    creds = flow.credentials
    st.session_state["google_credentials"] = {
        "token": creds.token,
        "refresh_token": creds.refresh_token,
        "token_uri": creds.token_uri,
        "client_id": creds.client_id,
        "client_secret": creds.client_secret,
        "scopes": creds.scopes,
    }
    return True


def list_google_calendars() -> List[dict]:
    if "google_credentials" not in st.session_state or Credentials is None or build is None:
        return []

    creds = Credentials(**st.session_state["google_credentials"])
    service = build("calendar", "v3", credentials=creds)
    result = service.calendarList().list().execute()
    return result.get("items", [])


def read_google_events(calendar_ids: List[str], year: int, month: int) -> Dict[str, List[str]]:
    if "google_credentials" not in st.session_state or Credentials is None or build is None:
        return {}

    time_min, time_max = iso_month_range(year, month)
    creds = Credentials(**st.session_state["google_credentials"])
    service = build("calendar", "v3", credentials=creds)

    events_by_date: Dict[str, List[str]] = {}
    for cal_id in calendar_ids:
        result = service.events().list(
            calendarId=cal_id,
            timeMin=time_min,
            timeMax=time_max,
            singleEvents=True,
            orderBy="startTime",
            maxResults=2500,
        ).execute()

        for event in result.get("items", []):
            start = event.get("start", {})
            raw_start = start.get("date") or start.get("dateTime")
            if not raw_start:
                continue

            iso_date = raw_start[:10]
            title = event.get("summary", "אירוע ללא כותרת")
            events_by_date.setdefault(iso_date, []).append(title)

    return events_by_date


def events_to_cell(events_by_date: Dict[str, List[str]], d: date) -> str:
    events = events_by_date.get(d.isoformat(), [])
    if not events:
        return ""
    # Keep the table readable.
    return " | ".join(events[:4]) + (" ..." if len(events) > 4 else "")


# =========================
# Availability table
# =========================

def build_month_dataframe(year: int, month: int, events_by_date: Dict[str, List[str]]) -> pd.DataFrame:
    holidays = holiday_map_for_month(year, month)
    rows = []

    for d in month_dates(year, month):
        rows.append({
            "date": d.isoformat(),
            "תאריך": d.strftime("%d/%m/%Y"),
            "יום": hebrew_weekday(d),
            "חגים / שבתות": holidays.get(d.isoformat(), ""),
            "אירועים מהיומן": events_to_cell(events_by_date, d),
            "חסום לתורנות": False,
            "יום חופש": False,
            "העדפה רכה": "",
            "הערה": "",
        })

    return pd.DataFrame(rows)



def summarize_submission(df: pd.DataFrame, person_id: str, year: int, month: int) -> List[dict]:
    constraints = []

    for _, row in df.iterrows():
        if bool(row.get("חסום לתורנות")):
            constraints.append({
                "person_id": person_id,
                "date": row["date"],
                "day_in_month": int(str(row["date"])[8:10]),
                "type": "unavailable_for_shift",
                "strength": "hard",
                "source": "manual",
                "note": str(row.get("הערה", "") or ""),
            })

        if bool(row.get("יום חופש")):
            constraints.append({
                "person_id": person_id,
                "date": row["date"],
                "day_in_month": int(str(row["date"])[8:10]),
                "type": "vacation_request",
                "strength": "hard",
                "source": "manual",
                "note": str(row.get("הערה", "") or ""),
            })

        preference = str(row.get("העדפה רכה", "") or "").strip()
        if preference:
            constraints.append({
                "person_id": person_id,
                "date": row["date"],
                "day_in_month": int(str(row["date"])[8:10]),
                "type": "preference",
                "strength": "soft",
                "source": "manual",
                "note": preference,
            })

    return constraints


def build_copyable_submission_text(df: pd.DataFrame, employee_name: str) -> str:
    blocked_days = []
    vacation_days = []

    for _, row in df.iterrows():
        day_num = int(str(row["date"])[8:10])
        if bool(row.get("חסום לתורנות")):
            blocked_days.append(day_num)
        if bool(row.get("יום חופש")):
            vacation_days.append(day_num)

    blocked = ",".join(str(x) for x in blocked_days)
    vacations = ",".join(str(x) for x in vacation_days)

    return f"{employee_name}\nחסימות- {blocked}\nחופשים- {vacations}"


def persist_submission(payload: dict) -> None:
    with OUTPUT_FILE.open("a", encoding="utf-8") as f:
        f.write(json.dumps(payload, ensure_ascii=False) + "\n")


# =========================
# UI
# =========================

def init_state():
    if "selected_year" not in st.session_state or "selected_month" not in st.session_state:
        y, m = default_next_month()
        st.session_state["selected_year"] = y
        st.session_state["selected_month"] = m

    if "events_by_date" not in st.session_state:
        st.session_state["events_by_date"] = {}


def move_month(delta: int):
    y, m = add_months(st.session_state["selected_year"], st.session_state["selected_month"], delta)
    st.session_state["selected_year"] = y
    st.session_state["selected_month"] = m


def render_shift_planning():
    init_state()

    st.header("תכנון תורנויות")
    st.caption("הזנת זמינות, חופשות וחסימות עבור החודש המוצג.")

    with st.sidebar:
        st.subheader("הגדרות תכנון")
        employee_name = st.text_input("שם העובד", value="יאיר")
        person_id = st.text_input("מזהה פנימי", value="demo_user_001")

        st.divider()
        st.subheader("חיבור יומנים אישיים")

        calendar_mode = st.radio(
            "מקור אירועים",
            ["ללא חיבור", "Google Calendar"],
            index=0,
            help="ב-MVP האירועים מוצגים כהקשר בלבד ולא הופכים אוטומטית לחסימות.",
        )

        selected_calendar_ids: List[str] = []

        if calendar_mode == "Google Calendar":
            if not google_is_configured():
                st.warning("Google Calendar לא מוגדר עדיין. הוסף secrets ב-Streamlit Cloud או בקובץ מקומי.")
                st.code(
                    """[google_oauth]
client_id = "YOUR_CLIENT_ID"
client_secret = "YOUR_CLIENT_SECRET"
redirect_uri = "https://YOUR_APP.streamlit.app"
""",
                    language="toml",
                )
            else:
                query_params = st.query_params
                code = query_params.get("code")

                if code and "google_credentials" not in st.session_state:
                    try:
                        if finish_google_oauth(code):
                            st.success("החיבור ליומן הושלם.")
                    except Exception as e:
                        st.error(f"שגיאה בהשלמת החיבור: {e}")

                if "google_credentials" not in st.session_state:
                    auth_url = start_google_oauth()
                    if auth_url:
                        st.link_button("התחבר ל-Google Calendar", auth_url)
                else:
                    st.success("מחובר ל-Google Calendar")
                    try:
                        calendars = list_google_calendars()
                        options = {
                            f"{c.get('summary', c.get('id'))} ({c.get('id')})": c["id"]
                            for c in calendars
                        }
                        selected_labels = st.multiselect(
                            "בחר יומנים לקריאה",
                            list(options.keys()),
                            default=list(options.keys())[:1],
                        )
                        selected_calendar_ids = [options[label] for label in selected_labels]

                        if st.button("טען אירועים לחודש המוצג"):
                            st.session_state["events_by_date"] = read_google_events(
                                selected_calendar_ids,
                                st.session_state["selected_year"],
                                st.session_state["selected_month"],
                            )
                            st.success("האירועים נטענו.")
                    except Exception as e:
                        st.error(f"שגיאה בקריאת יומנים: {e}")

        st.divider()
        st.subheader("שמירה")
        st.caption(f"קובץ פלט מקומי: `{OUTPUT_FILE}`")

    # Month navigation
    col_prev, col_title, col_next, col_reset = st.columns([1, 3, 1, 1])
    with col_prev:
        if st.button("← חודש קודם", use_container_width=True):
            move_month(-1)
            st.rerun()

    with col_title:
        st.markdown(
            f"<h2 style='text-align:center'>{month_title(st.session_state['selected_year'], st.session_state['selected_month'])}</h2>",
            unsafe_allow_html=True,
        )

    with col_next:
        if st.button("חודש הבא →", use_container_width=True):
            move_month(1)
            st.rerun()

    with col_reset:
        if st.button("חזור לחודש הבא", use_container_width=True):
            y, m = default_next_month()
            st.session_state["selected_year"] = y
            st.session_state["selected_month"] = m
            st.rerun()

    y = st.session_state["selected_year"]
    m = st.session_state["selected_month"]

    df = build_month_dataframe(y, m, st.session_state["events_by_date"])

    st.info("ברירת המחדל היא החודש הבא. ניתן לנוע קדימה ואחורה ללא מגבלה כדי לתכנן חצי שנה ומעלה.")

    edited_df = st.data_editor(
        df,
        key=f"availability_editor_{y}_{m}",
        hide_index=True,
        use_container_width=True,
        disabled=["date", "תאריך", "יום", "חגים / שבתות", "אירועים מהיומן"],
        column_config={
            "date": None,
            "חסום לתורנות": st.column_config.CheckboxColumn("חסום לתורנות"),
            "יום חופש": st.column_config.CheckboxColumn("יום חופש"),
            "העדפה רכה": st.column_config.TextColumn("העדפה רכה", help="לדוגמה: עדיף לא לילה / עדיף לא ערב"),
            "הערה": st.column_config.TextColumn("הערה"),
        },
    )

    st.divider()

    col_finish, col_preview = st.columns([1, 3])
    constraints = summarize_submission(edited_df, person_id, y, m)
    submission_text = build_copyable_submission_text(edited_df, employee_name)

    with col_finish:
        submitted = st.button("סיום / שמור", type="primary", use_container_width=True)

    with col_preview:
        st.subheader("סיכום בחירות")
        if constraints:
            summary_df = pd.DataFrame(constraints)
            st.dataframe(summary_df, hide_index=True, use_container_width=True)
        else:
            st.caption("לא נבחרו חסימות, חופשות או העדפות בחודש זה.")

    if submitted:
        payload = {
            "submitted_at": datetime.now(timezone.utc).isoformat(),
            "display_month": f"{y}-{m:02d}",
            "person_id": person_id,
            "employee_name": employee_name,
            "copyable_text": submission_text,
            "constraints": constraints,
        }
        persist_submission(payload)
        st.success(f"נשמרו {len(constraints)} אילוצים/בקשות עבור {month_title(y, m)}.")

        st.subheader("פלט להעתקה")
        st.caption("העתק את הטקסט הבא ושלח אותו לריכוז. המבנה כולל שם, חסימות וחופשים לפי מספרי הימים בחודש.")
        st.text_area(
            "טקסט להעתקה",
            value=submission_text,
            height=110,
            label_visibility="collapsed",
        )

        st.download_button(
            "הורד JSON",
            data=json.dumps(payload, ensure_ascii=False, indent=2),
            file_name=f"availability_{person_id}_{y}_{m:02d}.json",
            mime="application/json",
        )


    st.markdown("---")
    st.caption("נבנה ע\"י ד\"ר יאיר חזקיהו שטיינברג")


def parse_worker_blocks(raw_text: str) -> pd.DataFrame:
    """
    Expected block format:
    יאיר
    חסימות- 2,3,4
    חופשים- 3,7

    Blocks can be separated by empty lines or pasted one after another.
    """
    lines = [line.strip() for line in raw_text.splitlines() if line.strip()]
    records = []
    i = 0

    while i < len(lines):
        name = lines[i]
        blocked_line = lines[i + 1] if i + 1 < len(lines) else ""
        vacation_line = lines[i + 2] if i + 2 < len(lines) else ""

        def extract_days(line: str) -> List[int]:
            if "-" in line:
                line = line.split("-", 1)[1]
            parts = re.split(r"[,;\\s]+", line.strip())
            days = []
            for part in parts:
                if not part:
                    continue
                try:
                    value = int(part)
                    if 1 <= value <= 31:
                        days.append(value)
                except ValueError:
                    pass
            return sorted(set(days))

        blocked_days = extract_days(blocked_line)
        vacation_days = extract_days(vacation_line)

        records.append({
            "שם עובד": name,
            "חסימות": ",".join(str(d) for d in blocked_days),
            "חופשים": ",".join(str(d) for d in vacation_days),
            "מספר חסימות": len(blocked_days),
            "מספר חופשים": len(vacation_days),
        })

        i += 3

    return pd.DataFrame(records)


def build_schedule_layout_table(parsed_df: pd.DataFrame, year: int, month: int) -> pd.DataFrame:
    """
    Creates a manual-planning table:
    - Two last day numbers of previous month.
    - All day numbers of current month.
    - Weekday column.
    - Empty manual columns: מחלקה, מיון, שישי בוקר.
    - Employee columns marked by blocked dates.
    """
    previous_year, previous_month = add_months(year, month, -1)
    previous_last_day = calendar.monthrange(previous_year, previous_month)[1]
    current_last_day = calendar.monthrange(year, month)[1]

    row_dates = [
        date(previous_year, previous_month, previous_last_day - 1),
        date(previous_year, previous_month, previous_last_day),
    ] + [date(year, month, day) for day in range(1, current_last_day + 1)]

    employee_names = parsed_df["שם עובד"].tolist() if not parsed_df.empty else []

    blocked_by_employee = {}
    vacation_by_employee = {}
    for _, row in parsed_df.iterrows():
        name = row["שם עובד"]
        blocked_by_employee[name] = {int(x) for x in str(row["חסימות"]).split(",") if str(x).strip().isdigit()}
        vacation_by_employee[name] = {int(x) for x in str(row["חופשים"]).split(",") if str(x).strip().isdigit()}

    rows = []
    for d in row_dates:
        is_current_month = d.year == year and d.month == month
        out = {
            "יום בחודש": d.day,
            "יום בשבוע": hebrew_weekday(d),
            "מחלקה": "",
            "מיון": "",
            "שישי בוקר": "",
        }

        for employee in employee_names:
            if is_current_month and d.day in blocked_by_employee.get(employee, set()):
                out[employee] = "חסום"
            elif is_current_month and d.day in vacation_by_employee.get(employee, set()):
                out[employee] = "חופש"
            else:
                out[employee] = ""

        rows.append(out)

    return pd.DataFrame(rows)


def style_schedule_dataframe(df: pd.DataFrame):
    """
    Streamlit preview styling:
    - Yellow in first two columns for Friday/Saturday.
    - Gray in employee cells for blocked dates.
    """
    def style_cell(value, row_weekday=None, col_name=None):
        if col_name in ["יום בחודש", "יום בשבוע"] and row_weekday in ["שישי", "שבת"]:
            return "background-color: #fff200"
        if value == "חסום":
            return "background-color: #b7b7b7; color: #b7b7b7"
        if value == "חופש":
            return "background-color: #d9eaf7"
        return ""

    styles = pd.DataFrame("", index=df.index, columns=df.columns)
    for idx, row in df.iterrows():
        for col in df.columns:
            styles.loc[idx, col] = style_cell(row[col], row["יום בשבוע"], col)

    return df.style.apply(lambda _: styles, axis=None)


def dataframe_to_xlsx_bytes(parsed_df: pd.DataFrame, schedule_df: pd.DataFrame, year: int, month: int) -> bytes:
    buffer = BytesIO()

    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        schedule_df.to_excel(writer, sheet_name="טבלה כללית", index=False)
        parsed_df.to_excel(writer, sheet_name="סיכום עובדים", index=False)

        wb = writer.book
        ws = wb["טבלה כללית"]
        ws.sheet_view.rightToLeft = True

        # Fills
        header_fill = PatternFill("solid", fgColor="D9EAF7")
        weekend_fill = PatternFill("solid", fgColor="FFF200")
        blocked_fill = PatternFill("solid", fgColor="B7B7B7")
        vacation_fill = PatternFill("solid", fgColor="D9EAF7")
        thin = Side(style="thin", color="D9D9D9")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)

        # Header style
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = border

        # Body style
        max_row = ws.max_row
        max_col = ws.max_column
        for row_idx in range(2, max_row + 1):
            weekday = ws.cell(row=row_idx, column=2).value

            # Yellow only in A-B for Friday/Saturday.
            if weekday in ["שישי", "שבת"]:
                ws.cell(row=row_idx, column=1).fill = weekend_fill
                ws.cell(row=row_idx, column=2).fill = weekend_fill

            for col_idx in range(1, max_col + 1):
                cell = ws.cell(row=row_idx, column=col_idx)
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = border

                if col_idx >= 6 and cell.value == "חסום":
                    cell.fill = blocked_fill
                    cell.font = Font(color="B7B7B7")
                elif col_idx >= 6 and cell.value == "חופש":
                    cell.fill = vacation_fill

        ws.freeze_panes = "F2"

        widths = {
            1: 10,  # יום בחודש
            2: 13,  # יום בשבוע
            3: 14,  # מחלקה
            4: 14,  # מיון
            5: 14,  # שישי בוקר
        }

        for col_idx in range(1, max_col + 1):
            letter = get_column_letter(col_idx)
            ws.column_dimensions[letter].width = widths.get(col_idx, 12)

        # Summary sheet formatting
        summary_ws = wb["סיכום עובדים"]
        summary_ws.sheet_view.rightToLeft = True
        summary_ws.freeze_panes = "A2"

        for cell in summary_ws[1]:
            cell.fill = header_fill
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = border

        for row in summary_ws.iter_rows(min_row=2):
            for cell in row:
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = border

        for col in summary_ws.columns:
            max_len = max(len(str(cell.value or "")) for cell in col)
            summary_ws.column_dimensions[col[0].column_letter].width = min(max(max_len + 2, 10), 28)

    return buffer.getvalue()


def render_general_table_coding():
    st.header("קידוד לטבלה כללית")
    st.caption("הדבק כאן פלטים מכל העובדים. הכלי יוצר טבלת חודש בסגנון Excel ידני: ימים, יום בשבוע, עמודות מילוי ידני, ואז עובדים.")

    st.markdown(
        """
        מבנה צפוי לכל עובד:

        ```text
        יאיר
        חסימות- 2,3,4,5,6,7,12,13,14,15,24
        חופשים- 3,7,15
        ```
        """
    )

    raw_text = st.text_area(
        "הדבק פלטים של עובדים",
        height=260,
        placeholder="יאיר\nחסימות- 2,3,4,5,6,7,12,13,14,15,24\nחופשים- 3,7,15\n\nדנה\nחסימות- 1,8,9\nחופשים- 14",
    )

    default_y, default_m = default_next_month()
    col_year, col_month = st.columns(2)
    with col_year:
        selected_year = st.number_input("שנה", min_value=2024, max_value=2100, value=default_y, step=1)
    with col_month:
        selected_month = st.number_input("חודש", min_value=1, max_value=12, value=default_m, step=1)

    st.info(
        "הטבלה תכלול את שני מספרי הימים האחרונים של החודש הקודם, ואז את כל ימי החודש הנבחר. "
        "עמודות מחלקה, מיון ושישי בוקר נשארות ריקות למילוי ידני."
    )

    save_mode = st.radio(
        "בחר דרך עבודה",
        ["יצירת קובץ Excel לאחר הדבקה", "שמירה בגוגל ספרדשיט משותף - בהמשך"],
        index=0,
    )

    if save_mode == "שמירה בגוגל ספרדשיט משותף - בהמשך":
        st.info("אפשרות זו תדרוש חיבור Google Sheets API והרשאות כתיבה. בשלב הנוכחי מומלץ להשתמש ביצירת Excel.")

    if st.button("קודד לטבלה כללית", type="primary"):
        if not raw_text.strip():
            st.warning("יש להדביק לפחות פלט אחד.")
            return

        parsed_df = parse_worker_blocks(raw_text)
        schedule_df = build_schedule_layout_table(parsed_df, int(selected_year), int(selected_month))

        st.success(f"קודדו {len(parsed_df)} עובדים עבור {month_title(int(selected_year), int(selected_month))}.")

        st.subheader("סיכום עובדים")
        st.dataframe(parsed_df, hide_index=True, use_container_width=True)

        st.subheader("טבלה כללית")
        st.caption("בתצוגה: שישי/שבת צהוב בעמודות הראשונות; חסימות מסומנות באפור בעמודות העובדים.")
        st.dataframe(style_schedule_dataframe(schedule_df), hide_index=True, use_container_width=True)

        xlsx_bytes = dataframe_to_xlsx_bytes(parsed_df, schedule_df, int(selected_year), int(selected_month))
        st.download_button(
            "הורד Excel",
            data=xlsx_bytes,
            file_name=f"availability_general_table_{int(selected_year)}_{int(selected_month):02d}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


def render_calendar_save():
    st.header("שמירה ביומן")
    st.info("כלי זה טרם נבנה. בהמשך ניתן יהיה לשמור תורנויות מאושרות ליומן אישי או מחלקתי.")
    st.markdown(
        """
        כיוון אפשרי לכלי:
        - בחירת סידור תורנויות סופי.
        - בחירת יומן יעד.
        - יצירת אירועי יומן.
        - שמירה לאחר אישור משתמש.
        """
    )


def render_sidebar_navigation() -> str:
    st.sidebar.title("מע׳ לתכנון תורנויות- פנימית ד׳")
    st.sidebar.markdown("### כלי עזר")
    st.sidebar.markdown("---")

    return st.sidebar.radio(
        "בחר כלי",
        [
            "תכנון תורנויות",
            "קידוד לטבלה כללית",
            "שמירה ביומן",
        ],
        index=0,
        label_visibility="collapsed",
    )


def main():
    selected_tool = render_sidebar_navigation()


    if selected_tool == "תכנון תורנויות":
        render_shift_planning()
    elif selected_tool == "קידוד לטבלה כללית":
        render_general_table_coding()
    elif selected_tool == "שמירה ביומן":
        render_calendar_save()


    st.markdown("---")
    st.caption('נבנה ע"י ד"ר יאיר חזקיהו שטיינברג')


if __name__ == "__main__":
    main()
