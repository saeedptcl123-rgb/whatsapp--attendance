import difflib
import io
import json
import re
import time
from datetime import date, datetime

from groq import BadRequestError, Groq, RateLimitError
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell

MODEL_NAME = "openai/gpt-oss-120b"

SYSTEM_PROMPT = """You are an attendance extraction assistant.
You receive messy WhatsApp group chat text containing attendance reports.
Extract every attendance entry and return ONLY a JSON object of this form:
{"records": [{"name": "Person Name", "date": "DD/MM/YYYY", "status": "P"}]}

Rules:
- status must be exactly one of: "P" (present), "A" (absent), "L" (leave).
  Map words such as present, haazir, hazir, here -> "P";
  absent, gair haazir, ghair haazir -> "A";
  leave, chutti, chhutti, off -> "L".
- name: the sender of the attendance message. Use the name exactly as shown
  in the chat (for example "sarfraz" or "Saeed Ahmad"). If the sender is only
  shown as a phone number, return that phone number exactly as written.
  Remove timestamps, brackets and the leading "-" or ":" characters.
- date: the date shown in the message line (for example "01/11/2026" or
  "[08/11/2026, 9:00 PM]"). If the line has no date, use the default date.
  Always output dates in DD/MM/YYYY format.
- Ignore messages that are not about attendance (greetings, jokes, links,
  "Good morning", media messages, and so on).
- If one person appears more than once for the same date, keep only the last status.
- Never invent people or dates. If nothing is found, return {"records": []}.
"""

DATE_FORMATS = [
    "%d/%m/%Y",
    "%d-%m-%Y",
    "%d.%m.%Y",
    "%d/%m/%y",
    "%d-%m-%y",
    "%Y-%m-%d",
    "%d %b %Y",
    "%d %B %Y",
]


def match_key(value):
    """
    Return a simplified key used for matching names and phone numbers.
    It lowercases the text and removes spaces, punctuation and symbols,
    so "+92 321 5856154" and "923215856154" give the same key.
    """
    key = re.sub(r"[^\w]", "", str(value or "").lower())
    # Treat local Pakistani numbers (03xxxxxxxxx) and international ones (923xxxxxxxxx) as the same.
    if key.isdigit() and len(key) == 11 and key.startswith("0"):
        key = "92" + key[1:]
    return key


def normalize_status(raw_status):
    """Convert a status word or letter into 'P', 'A' or 'L'. Returns None if unknown."""
    text = str(raw_status or "").strip().lower()
    if text in {"p", "present", "haazir", "hazir", "here"}:
        return "P"
    if text in {"a", "absent", "gair haazir", "ghair haazir"}:
        return "A"
    if text in {"l", "leave", "chutti", "chhutti", "off"}:
        return "L"
    return None


def parse_date_value(value):
    """Convert a string, datetime or date into a date object. Returns None if not possible."""
    if value is None or value == "":
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    text = str(value).strip()
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    return None


def parse_alias_text(alias_text):
    """
    Parse the optional alias box. Each line should look like:
        +92 321 5856154 = Sarfraz Khan
    Left side: name or phone as it appears in the WhatsApp chat.
    Right side: name exactly as it appears in the Excel template.
    Returns a dict: {match_key(chat name): template name}.
    """
    aliases = {}
    for line in (alias_text or "").splitlines():
        if "=" not in line:
            continue
        left, right = line.split("=", 1)
        left_key = match_key(left)
        right_name = right.strip()
        if left_key and right_name:
            aliases[left_key] = right_name
    return aliases


# ---------------- Local rules (no AI needed) ----------------
# WhatsApp line format: 01/10/2026, 08:40 - Sender: message text
CHAT_LINE_PATTERN = re.compile(r"^(\d{2})/(\d{2})/(\d{4}), \d{1,2}:\d{2}(?::\d{2})? ?(?:AM|PM|am|pm)? - (.+?): (.*)$")
SYSTEM_LINE_PATTERN = re.compile(r"^\d{2}/\d{2}/\d{4}, .*? - ")

# Words that mean the person is on leave (checked first).
LEAVE_WORDS = [
    "leave", "sick", "fever", "can't come", "cannot come", "not come",
    "not join", "death", "died", "uncle", "chutti",
]
# Words that mean the person is present (late counts as present).
PRESENT_WORDS = ["present", "late", "haazir", "hazir"]


def classify_message_locally(text):
    """
    Return "L", "P" or None (None = not clear, send to the AI).
    Leave words are checked first, then present or late words.
    """
    normalized = text.lower().replace("’", "'")
    if any(word in normalized for word in LEAVE_WORDS):
        return "L"
    if any(word in normalized for word in PRESENT_WORDS):
        return "P"
    return None


def split_chat_messages(chat_text):
    """
    Group the chat into messages. Each message is a dict with date, sender,
    text and the full original text (for the AI). System lines (such as
    "X added Y") and deleted messages are dropped.
    """
    messages = []
    for line in chat_text.splitlines():
        match = CHAT_LINE_PATTERN.match(line)
        if match:
            d, mo, y, sender, text = match.groups()
            messages.append({
                "date": date(int(y), int(mo), int(d)),
                "sender": sender.strip(),
                "text": text.strip(),
                "raw": line,
            })
        elif SYSTEM_LINE_PATTERN.match(line):
            # System line such as "Someone added Someone": not a message.
            messages.append(None)
        elif messages and messages[-1] is not None and line.strip():
            # Continuation of a multi-line message.
            messages[-1]["text"] += " " + line.strip()
            messages[-1]["raw"] += "\n" + line
    return [m for m in messages if m is not None and "this message was deleted" not in m["text"].lower()]


def classify_chat_locally(chat_text, default_date=None):
    """
    Split the chat and classify each message with local rules.
    Returns (records, leftover_text):
      - records: classified entries (the last message per person per day wins)
      - leftover_text: messages the rules could not classify, to send to the AI.
    """
    latest = {}
    leftover_lines = []
    for message in split_chat_messages(chat_text):
        status = classify_message_locally(message["text"])
        if status is None:
            leftover_lines.append(message["raw"])
            continue
        key = (match_key(message["sender"]), message["date"])
        latest[key] = {"name": message["sender"], "date": message["date"], "status": status}
    return list(latest.values()), "\n".join(leftover_lines)


# Free-tier Groq limits are small (for example 8000 tokens per minute), so the
# chat is sent in smaller pieces. Each piece stays well under the limit.
CHUNK_MAX_CHARS = 2500       # max characters of chat text per request
MAX_OUTPUT_TOKENS = 3000     # max tokens the model may return per request
JSON_RETRIES = 1             # retry once if the model returns broken JSON
PAUSE_BETWEEN_CHUNKS = 20    # seconds to wait between requests
RATE_LIMIT_RETRIES = 4       # how many times to retry after a rate-limit error
RATE_LIMIT_WAIT = 30         # seconds to wait after a rate-limit error


def split_chat_into_chunks(chat_text, max_chars=CHUNK_MAX_CHARS):
    """
    Split the chat into pieces of at most max_chars, always at line breaks,
    so no WhatsApp message is cut in half.
    """
    chunks = []
    current = []
    current_len = 0
    for line in chat_text.splitlines():
        line_len = len(line) + 1
        if current and current_len + line_len > max_chars:
            chunks.append("\n".join(current))
            current = []
            current_len = 0
        current.append(line)
        current_len += line_len
    if current:
        chunks.append("\n".join(current))
    return chunks


def call_groq_for_chunk(client, chunk_text, default_text):
    """
    Send one piece of chat to Groq and return the parsed JSON dict.
    Returns {"records": []} for a piece that could not be read, and prints a
    warning in the terminal so the problem is visible.
    """
    user_message = (
        f"Default date (use only if the line has no date): {default_text}\n\n"
        f"WhatsApp chat:\n{chunk_text}"
    )

    json_failures = 0
    rate_limit_failures = 0
    while True:
        try:
            response = client.chat.completions.create(
                model=MODEL_NAME,
                messages=[
                    {"role": "system", "content": SYSTEM_PROMPT},
                    {"role": "user", "content": user_message},
                ],
                temperature=0,
                max_tokens=MAX_OUTPUT_TOKENS,
                response_format={"type": "json_object"},
            )
        except RateLimitError:
            rate_limit_failures += 1
            if rate_limit_failures > RATE_LIMIT_RETRIES:
                raise ValueError(
                    "Groq ki rate limit bar bar aa rahi hai. Thodi der baad dobara try karein."
                )
            time.sleep(RATE_LIMIT_WAIT)
            continue
        except BadRequestError as exc:
            # Groq rejects the reply when its JSON is broken (json_validate_failed).
            if "json_validate_failed" not in str(exc):
                raise
            json_failures += 1
            if json_failures > JSON_RETRIES:
                print("WARNING: is hisse ka jawab parha nahi gaya, skip kiya:")
                print(chunk_text[:300])
                return {"records": []}
            continue

        content = response.choices[0].message.content or ""
        try:
            return json.loads(content)
        except json.JSONDecodeError:
            json_failures += 1
            if json_failures > JSON_RETRIES:
                print("WARNING: is hisse ka jawab parha nahi gaya, skip kiya:")
                print(chunk_text[:300])
                return {"records": []}


def extract_attendance_records(chat_text, api_key, default_date=None):
    """
    Send the WhatsApp chat to Groq in small pieces and return a clean list of
    attendance records.

    Each record is a dict: {"name": str, "date": date, "status": "P" | "A" | "L"}.
    Raises ValueError for user-facing problems (empty input, bad API response).
    """
    if not chat_text or not chat_text.strip():
        raise ValueError("WhatsApp chat ka text khali hai.")
    if not api_key:
        raise ValueError("Groq API key zaroori hai.")

    # Step 1: local rules handle every clear message (P, L, late, fever, death, etc.).
    local_records, leftover_text = classify_chat_locally(chat_text, default_date)

    # Step 2: only the unclear messages go to the AI, in small pieces.
    raw_records = []
    if leftover_text.strip():
        default_text = default_date.strftime("%d/%m/%Y") if default_date else "not provided"
        client = Groq(api_key=api_key)
        chunks = split_chat_into_chunks(leftover_text)
        for index, chunk in enumerate(chunks):
            if index > 0:
                time.sleep(PAUSE_BETWEEN_CHUNKS)
            data = call_groq_for_chunk(client, chunk, default_text)
            if isinstance(data, dict):
                raw_records.extend(data.get("records", []) or [])

    default_fallback = default_date if isinstance(default_date, date) else None

    cleaned = list(local_records)
    for item in raw_records:
        if not isinstance(item, dict):
            continue
        name = re.sub(r"\s+", " ", str(item.get("name") or "")).strip()
        status = normalize_status(item.get("status"))
        record_date = parse_date_value(item.get("date")) or default_fallback
        if not name or status is None or record_date is None:
            continue
        cleaned.append({"name": name, "date": record_date, "status": status})

    # If the same person has several entries for the same date, keep only the last one.
    unique = {}
    for record in cleaned:
        unique[(match_key(record["name"]), record["date"])] = record
    return list(unique.values())


def find_template_row(chat_name, name_rows):
    """
    Find the template row for a chat name.
    name_rows maps match_key(template name) -> (row number, template name).
    Matching order: exact key, then one-way containment (for example
    "sarfraz" matches "Sarfraz Khan"), then close fuzzy match.
    Returns (row, template_name) or (None, suggestion).
    """
    key = match_key(chat_name)
    if not key:
        return None, None

    if key in name_rows:
        return name_rows[key]

    contained = [k for k in name_rows if key in k or k in key]
    if len(contained) == 1:
        return name_rows[contained[0]]

    close = difflib.get_close_matches(key, list(name_rows.keys()), n=1, cutoff=0.75)
    if close:
        return None, name_rows[close[0]][1]
    return None, None


LEGEND_LETTERS = {"P", "A", "L", "H"}


def find_free_name_row(worksheet, header_row, name_col):
    """
    Return the first row below the header whose name cell is empty and which
    looks like an employee row (Sr. No. is a number or empty, not a legend letter).
    Returns None if no such row exists.
    """
    for row in range(header_row + 1, worksheet.max_row + 1):
        name_cell = worksheet.cell(row=row, column=name_col)
        sr_cell = worksheet.cell(row=row, column=1)
        # Skip merged cells (for example a header that spans two rows).
        if isinstance(name_cell, MergedCell) or isinstance(sr_cell, MergedCell):
            continue
        if name_cell.value not in (None, ""):
            continue
        sr_value = sr_cell.value
        if isinstance(sr_value, str) and sr_value.strip().upper() in LEGEND_LETTERS:
            continue
        if isinstance(sr_value, str) and sr_value.strip() and not sr_value.strip().isdigit():
            continue
        return row
    return None


def add_name_row(worksheet, header_row, name_col, name, name_rows):
    """
    Write a new name into the first free employee row (or add it at the end
    if no free row exists) and register it in name_rows. Returns the row number.
    """
    row = find_free_name_row(worksheet, header_row, name_col)
    if row is None:
        row = worksheet.max_row + 1
    worksheet.cell(row=row, column=name_col, value=name)
    name_rows[match_key(name)] = (row, name)
    return row


def write_total_formulas(worksheet, header_row, name_col, day_cols):
    """
    Add per-employee totals using Excel formulas, for the columns labelled
    P, H, L and Days placed after the day columns (e.g. AI, AJ, AK, AL).
    Only rows that have a name get formulas.
    Days = number of marked days (any of P, A, L, H).
    """
    if not day_cols:
        return
    first_day_col = min(day_cols.values())
    last_day_col = max(day_cols.values())

    # Find the total columns by their label in the header rows.
    label_cols = {}
    for label_row in (header_row, header_row + 1):
        for col in range(last_day_col + 1, worksheet.max_column + 1):
            cell = worksheet.cell(row=label_row, column=col)
            if isinstance(cell, MergedCell) or cell.value is None:
                continue
            label = str(cell.value).strip().upper()
            if label in {"P", "H", "L", "DAYS"} and label not in label_cols:
                label_cols[label] = col
    if not label_cols:
        return

    day_range = lambda row: (
        f"{worksheet.cell(row=row, column=first_day_col).coordinate}:"
        f"{worksheet.cell(row=row, column=last_day_col).coordinate}"
    )

    for row in range(header_row + 1, worksheet.max_row + 1):
        name_cell = worksheet.cell(row=row, column=name_col)
        if isinstance(name_cell, MergedCell) or name_cell.value in (None, ""):
            continue
        rng = day_range(row)
        formulas = {
            "P": f'=COUNTIF({rng},"P")',
            "H": f'=COUNTIF({rng},"H")',
            "L": f'=COUNTIF({rng},"L")',
            "DAYS": f'=COUNTA({rng})',
        }
        for label, formula in formulas.items():
            if label in label_cols:
                target = worksheet.cell(row=row, column=label_cols[label])
                if not isinstance(target, MergedCell):
                    target.value = formula


def fill_attendance_excel(
    template_bytes,
    records,
    name_col,
    header_row,
    sheet_month,
    sheet_name=None,
    aliases=None,
    auto_add_names=True,
):
    """
    Write attendance records into an Excel template.

    Args:
        template_bytes: the uploaded .xlsx file as bytes.
        records: list of dicts from extract_attendance_records.
        name_col: column number that holds employee names (1-based).
        header_row: row number that holds day numbers or dates (1-based).
        sheet_month: a date inside the month the sheet is for (e.g. 2026-10-01).
                     Records from other months are skipped and reported.
        sheet_name: optional sheet name. If empty, the first sheet is used.
        aliases: optional dict from parse_alias_text.
        auto_add_names: if True, a chat name that is not in the template is
                        written into the next free name row automatically.

    Returns (excel_bytes, summary).
    """
    aliases = aliases or {}
    added_names = []
    workbook = load_workbook(io.BytesIO(template_bytes))

    if sheet_name:
        if sheet_name not in workbook.sheetnames:
            raise ValueError(f"Sheet '{sheet_name}' template mein nahi mili.")
        worksheet = workbook[sheet_name]
    else:
        worksheet = workbook.active

    # Map match_key(name) -> (row, display name). Rows below the header are scanned.
    name_rows = {}
    for row in range(header_row + 1, worksheet.max_row + 1):
        cell_value = worksheet.cell(row=row, column=name_col).value
        if cell_value is None or str(cell_value).strip() == "":
            continue
        key = match_key(cell_value)
        if key and key not in name_rows:
            name_rows[key] = (row, str(cell_value).strip())

    # Map day number -> column. The header can hold day numbers (1-31) or real dates.
    day_cols = {}
    for col in range(1, worksheet.max_column + 1):
        value = worksheet.cell(row=header_row, column=col).value
        day = None
        if isinstance(value, (datetime, date)):
            day = value.day
        elif isinstance(value, (int, float)) and float(value).is_integer() and 1 <= value <= 31:
            day = int(value)
        if day is not None and day not in day_cols:
            day_cols[day] = col

    written = 0
    unmatched = {}
    skipped_dates = []

    for record in records:
        rec_date = record["date"]

        # Skip records that belong to a different month than the sheet.
        if rec_date.year != sheet_month.year or rec_date.month != sheet_month.month:
            skipped_dates.append(record["name"] + " (" + rec_date.strftime("%d/%m/%Y") + ")")
            continue

        col = day_cols.get(rec_date.day)
        if col is None:
            skipped_dates.append(record["name"] + " (" + rec_date.strftime("%d/%m/%Y") + ")")
            continue

        # Use the alias if one exists for this chat name or phone number.
        lookup_name = aliases.get(match_key(record["name"]), record["name"])

        result = find_template_row(lookup_name, name_rows)
        if isinstance(result, tuple) and len(result) == 2 and isinstance(result[0], int):
            row = result[0]
        elif auto_add_names:
            suggestion = result[1] if isinstance(result, tuple) else None
            if suggestion:
                # A very similar name already exists (e.g. a spelling difference),
                # so use that row instead of creating a duplicate.
                row = name_rows[match_key(suggestion)][0]
            else:
                row = add_name_row(worksheet, header_row, name_col, lookup_name, name_rows)
                added_names.append(lookup_name)
        else:
            suggestion = result[1] if isinstance(result, tuple) else None
            unmatched[record["name"]] = suggestion
            continue

        worksheet.cell(row=row, column=col, value=record["status"])
        written += 1

    # Add the P / H / L / Days totals for every employee row.
    write_total_formulas(worksheet, header_row, name_col, day_cols)

    output = io.BytesIO()
    workbook.save(output)

    summary = {
        "written": written,
        "unmatched": unmatched,
        "added_names": added_names,
        "skipped_dates": skipped_dates,
        "total_records": len(records),
    }
    return output.getvalue(), summary

