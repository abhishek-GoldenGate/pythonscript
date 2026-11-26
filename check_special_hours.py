"""
check_special_hours.py

1. Finds the Excel file (.xlsx / .xlsm) sitting next to this script.
2. Scans every sheet for the "Store Status", "Store ID" and "Brand" columns
   and keeps the Store ID + Brand of rows whose Store Status is "Live".
   Then asks for the Excel row number to start from (Enter = row 1).
3. Asks for a curl command (paste it, then press Enter on an empty line),
   then asks the "partnerIds" value ONCE per distinct brand, before any call.
4. For each Store ID, puts it into "sourceIds" and its brand's value into
   "partnerIds" (JSON body or URL query parameter), runs curl, and checks
   the "specialHours" array:
     - it must exist and must not be empty
     - at least one entry must have date == EXPECTED_DATE and
       operatingHours == {startTime: EXPECTED_START_TIME, endTime: EXPECTED_END_TIME}
     - every OTHER entry dated in the same year (2026) must also have exactly
       those startTime/endTime values
     - entries from other years (e.g. 2024, 2025) are ignored
   Spacing/indentation don't matter; values, types and formats must be exact.
5. Prints the Store IDs that fail and saves them to
   mismatched_source_ids_<timestamp>.csv next to this script.

Requires: pip install openpyxl   (and curl on PATH, built into Windows 10+)
Usage:    python check_special_hours.py [optional_path_to_excel]
"""

import csv
import json
import re
from datetime import datetime
import shlex
import shutil
import subprocess
import sys
from pathlib import Path

try:
    from openpyxl import load_workbook
except ImportError:
    sys.exit("openpyxl is not installed. Run:  pip install openpyxl")

STATUS_HEADER = "store status"
ID_HEADER = "store id"
BRAND_HEADER = "brand"
LIVE_VALUE = "live"
# Name of the partner attribute in the curl command (first one found is used)
PARTNER_ATTR_NAMES = ("partnerIds", "parterIds")
HEADER_SEARCH_ROWS = 20      # how many top rows to search for the header row
CURL_TIMEOUT_SECONDS = 60

# ---- Expected specialHours values (edit these when the requirement changes) ----
EXPECTED_DATE = "2026-11-26T00:00:00.001Z"
EXPECTED_START_TIME = "08.00.00"
EXPECTED_END_TIME = "08.01.00"
CHECK_YEAR = EXPECTED_DATE[:4]   # entries of this year must have the expected times


# --------------------------------------------------------------------------- #
# Excel
# --------------------------------------------------------------------------- #
def find_excel_file() -> Path:
    if len(sys.argv) > 1:
        path = Path(sys.argv[1])
        if not path.is_file():
            sys.exit(f"Excel file not found: {path}")
        return path

    script_dir = Path(__file__).resolve().parent
    files = sorted(
        p for p in script_dir.iterdir()
        if p.suffix.lower() in (".xlsx", ".xlsm") and not p.name.startswith("~$")
    )
    if not files:
        sys.exit(f"No .xlsx/.xlsm file found in {script_dir}")
    if len(files) == 1:
        return files[0]

    print("Multiple Excel files found:")
    for i, f in enumerate(files, 1):
        print(f"  {i}. {f.name}")
    while True:
        choice = input("Choose a file number: ").strip()
        if choice.isdigit() and 1 <= int(choice) <= len(files):
            return files[int(choice) - 1]
        print("Invalid choice, try again.")


def normalize(value) -> str:
    return " ".join(str(value).split()).lower() if value is not None else ""


def cell_to_id(value) -> str:
    """Excel often stores numeric IDs as floats (1234.0) -> '1234'."""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def get_live_store_ids(excel_path: Path) -> list[tuple[int, str, str]]:
    """Return (excel_row_number, store_id, brand) for every 'Live' row."""
    wb = load_workbook(excel_path, read_only=True, data_only=True)
    store_ids: list[tuple[int, str, str]] = []
    found_headers = False

    for ws in wb.worksheets:
        rows = ws.iter_rows(values_only=True)
        status_col = id_col = brand_col = None
        rows_seen = 0

        # Scan all columns of the top rows to locate the header row
        for row in rows:
            rows_seen += 1
            headers = [normalize(c) for c in row]
            if all(h in headers for h in (STATUS_HEADER, ID_HEADER, BRAND_HEADER)):
                status_col = headers.index(STATUS_HEADER)
                id_col = headers.index(ID_HEADER)
                brand_col = headers.index(BRAND_HEADER)
                break
            if rows_seen >= HEADER_SEARCH_ROWS:
                break

        if status_col is None:
            print(f"  Sheet '{ws.title}': 'Store Status'/'Store ID'/'Brand' columns "
                  "not found, skipped.")
            continue

        found_headers = True
        sheet_count = 0
        last_col = max(status_col, id_col, brand_col)
        # continues after the header row; row_no is the Excel row number
        for row_no, row in enumerate(rows, start=rows_seen + 1):
            if row is None or len(row) <= last_col:
                continue
            if normalize(row[status_col]) == LIVE_VALUE:
                sid = row[id_col]
                if sid is None or str(sid).strip() == "":
                    continue
                brand = " ".join(str(row[brand_col] or "").split()) or "(no brand)"
                store_ids.append((row_no, cell_to_id(sid), brand))
                sheet_count += 1
        print(f"  Sheet '{ws.title}': {sheet_count} 'Live' store(s) found.")

    wb.close()
    if not found_headers:
        sys.exit("Could not find 'Store Status', 'Store ID' and 'Brand' columns in any sheet.")
    return store_ids


def warn_duplicate_store_ids(store_ids) -> None:
    """Store IDs are expected to be unique; list any that appear more than once."""
    rows_by_id: dict[str, list[int]] = {}
    for row_no, sid, _ in store_ids:
        rows_by_id.setdefault(sid, []).append(row_no)
    dups = {sid: rows for sid, rows in rows_by_id.items() if len(rows) > 1}
    if dups:
        print("  WARNING: these Store IDs appear more than once (each row is still checked):")
        for sid, rows in dups.items():
            print(f"    {sid}: rows {rows}")


def group_brands(store_ids) -> dict[str, str]:
    """
    Map normalized brand -> display name, in first-appearance order.
    'KFC', 'kfc' and ' KFC ' are treated as the same brand.
    """
    brands: dict[str, str] = {}
    for _, _, brand in store_ids:
        brands.setdefault(brand.lower(), brand)
    return brands


# --------------------------------------------------------------------------- #
# User input
# --------------------------------------------------------------------------- #
def read_multiline(prompt: str) -> str:
    print(prompt)
    print("(paste it, then press Enter on an empty line to finish)")
    lines = []
    while True:
        try:
            line = input()
        except EOFError:
            break
        if line.strip() == "" and lines:
            break
        if line.strip() == "":
            continue
        lines.append(line)
    return "\n".join(lines)


def parse_curl(text: str) -> list[str]:
    # Join line continuations: bash "\", cmd "^", PowerShell "`"
    text = re.sub(r"[\\^`][ \t]*\r?\n", " ", text)
    text = text.replace("\r", " ").replace("\n", " ")
    try:
        args = shlex.split(text, posix=True)
    except ValueError as e:
        sys.exit(f"Could not parse curl command: {e}")
    if not args or Path(args[0]).stem.lower() != "curl":
        sys.exit("The command must start with 'curl'.")
    return args


# --------------------------------------------------------------------------- #
# sourceIds substitution
# --------------------------------------------------------------------------- #
JSON_ARRAY_RE = re.compile(r'("sourceIds"\s*:\s*)\[[^\]]*\]')
JSON_STRING_RE = re.compile(r'("sourceIds"\s*:\s*)"[^"]*"')
JSON_NUMBER_RE = re.compile(r'("sourceIds"\s*:\s*)-?\d+(?:\.\d+)?')
QUERY_RE = re.compile(r'([?&]sourceIds=)[^&\s#]*')


def contains_source_ids(args: list[str]) -> bool:
    return any("sourceIds" in a for a in args)


def substitute_source_id(args: list[str], store_id: str) -> list[str]:
    quoted = json.dumps(store_id)
    new_args = []
    for a in args:
        if "sourceIds" in a:
            a = JSON_ARRAY_RE.sub(lambda m: f"{m.group(1)}[{quoted}]", a)
            a = JSON_STRING_RE.sub(lambda m: f"{m.group(1)}{quoted}", a)
            a = JSON_NUMBER_RE.sub(lambda m: f"{m.group(1)}{quoted}", a)
            a = QUERY_RE.sub(lambda m: f"{m.group(1)}{store_id}", a)
        new_args.append(a)
    return new_args


# --------------------------------------------------------------------------- #
# partnerIds substitution (value per brand, entered by the user)
# --------------------------------------------------------------------------- #
def find_partner_attr(args: list[str]) -> str | None:
    for name in PARTNER_ATTR_NAMES:
        if any(name in a for a in args):
            return name
    return None


def partner_uses_quotes(args: list[str], attr: str) -> bool:
    """Keep the curl's original style: ["123"] (strings) vs [123] (numbers)."""
    for a in args:
        m = re.search(rf'"{attr}"\s*:\s*\[?\s*(.)', a)
        if m:
            return m.group(1) == '"'
    return True


def substitute_partner_ids(args: list[str], attr: str, partner_ids: list[str],
                           quoted: bool) -> list[str]:
    items = [json.dumps(p) if quoted else p for p in partner_ids]
    arr_re = re.compile(rf'("{attr}"\s*:\s*)\[[^\]]*\]')
    str_re = re.compile(rf'("{attr}"\s*:\s*)"[^"]*"')
    num_re = re.compile(rf'("{attr}"\s*:\s*)-?\d+(?:\.\d+)?')
    query_re = re.compile(rf'([?&]{attr}=)[^&\s#]*')
    new_args = []
    for a in args:
        if attr in a:
            a = arr_re.sub(lambda m: f"{m.group(1)}[{', '.join(items)}]", a)
            a = str_re.sub(lambda m: f"{m.group(1)}{items[0]}", a)
            a = num_re.sub(lambda m: f"{m.group(1)}{items[0]}", a)
            a = query_re.sub(lambda m: f"{m.group(1)}{','.join(partner_ids)}", a)
        new_args.append(a)
    return new_args


def read_partner_ids(brands: dict[str, str], attr: str) -> dict[str, list[str]]:
    """Ask once per brand. Comma-separated input gives several IDs."""
    print(f"\nEnter the '{attr}' value for each brand "
          "(use commas for more than one, e.g. 123,456):")
    result: dict[str, list[str]] = {}
    for key, name in brands.items():
        while True:
            try:
                text = input(f"  {attr} for brand '{name}': ").strip()
            except EOFError:
                sys.exit("\nNo input; stopping.")
            ids = [p.strip().strip('"\'') for p in text.split(",") if p.strip()]
            if ids:
                result[key] = ids
                break
            print("    Value cannot be empty, please enter it again.")
    return result


# --------------------------------------------------------------------------- #
# Running curl and checking specialHours
# --------------------------------------------------------------------------- #
def run_curl(args: list[str]) -> str:
    exe = shutil.which("curl") or shutil.which("curl.exe")
    if not exe:
        sys.exit("curl was not found on PATH.")
    # -s: silent, -S: still show errors
    cmd = [exe, "-sS"] + args[1:]
    result = subprocess.run(cmd, capture_output=True, timeout=CURL_TIMEOUT_SECONDS)
    if result.returncode != 0:
        raise RuntimeError(result.stderr.decode("utf-8", "replace").strip()
                           or f"curl exited with code {result.returncode}")
    return result.stdout.decode("utf-8", "replace")


def find_special_hours(data, found=None) -> list:
    """Return every value stored under a 'specialHours' key, at any depth."""
    if found is None:
        found = []
    if isinstance(data, dict):
        for k, v in data.items():
            if k == "specialHours":
                found.append(v)
            find_special_hours(v, found)
    elif isinstance(data, list):
        for item in data:
            find_special_hours(item, found)
    return found


class RawNumber:
    """A JSON number kept exactly as written, so 900 != 900.0 != 9e2."""
    __slots__ = ("text",)

    def __init__(self, text: str):
        self.text = text

    def __eq__(self, other):
        return isinstance(other, RawNumber) and self.text == other.text

    def __hash__(self):
        return hash(self.text)


def parse_json_exact(text: str):
    """Parse JSON ignoring whitespace/indentation, but keeping number formats."""
    return json.loads(text, parse_int=RawNumber, parse_float=RawNumber)


def type_name(v) -> str:
    return {RawNumber: "number", str: "string", bool: "boolean", list: "array",
            dict: "object", type(None): "null"}.get(type(v), type(v).__name__)


def find_difference(actual, expected, path="specialHours"):
    """
    Return a description of the first difference, or None if identical.
    Strings (dates/times such as "09:00:00") must match character for
    character, so "09:00" vs "09:00:00" or "09:00:00" vs "09:00:01" fail.
    """
    if type(actual) is not type(expected):
        return (f"{path}: expected {type_name(expected)} {short(expected)}, "
                f"got {type_name(actual)} {short(actual)}")
    if isinstance(expected, dict):
        missing = [k for k in expected if k not in actual]
        extra = [k for k in actual if k not in expected]
        if missing:
            return f"{path}: missing key(s) {missing}"
        if extra:
            return f"{path}: unexpected key(s) {extra}"
        for k in expected:
            diff = find_difference(actual[k], expected[k], f"{path}.{k}")
            if diff:
                return diff
        return None
    if isinstance(expected, list):
        if len(actual) != len(expected):
            return f"{path}: expected {len(expected)} item(s), got {len(actual)}"
        for i, (a, e) in enumerate(zip(actual, expected)):
            diff = find_difference(a, e, f"{path}[{i}]")
            if diff:
                return diff
        return None
    if actual != expected:
        return f"{path}: expected {short(expected)}, got {short(actual)}"
    return None


def check_special_hours(value):
    """Return None if this specialHours value passes, else the reason it fails."""
    if not isinstance(value, list):
        return f"specialHours is not an array: {short(value)}"
    if not value:
        return "specialHours is an empty array []"

    expected_hours = {"startTime": EXPECTED_START_TIME, "endTime": EXPECTED_END_TIME}
    found_expected_date = False

    for i, entry in enumerate(value):
        path = f"specialHours[{i}]"
        if not isinstance(entry, dict):
            return f"{path} is not an object: {short(entry)}"
        date = entry.get("date")
        if not isinstance(date, str) or not date.startswith(CHECK_YEAR):
            continue    # other years (2024, 2025, ...) are ignored

        if "operatingHours" not in entry:
            return f"{path} (date {date}): missing 'operatingHours'"
        diff = find_difference(entry["operatingHours"], expected_hours,
                               f"{path}.operatingHours")
        if diff:
            return f"{diff} (date {date})"
        if date == EXPECTED_DATE:
            found_expected_date = True

    if not found_expected_date:
        return f"no entry with date {EXPECTED_DATE!r} found"
    return None


def to_display(value):
    if isinstance(value, RawNumber):
        return json.loads(value.text)
    if isinstance(value, dict):
        return {k: to_display(v) for k, v in value.items()}
    if isinstance(value, list):
        return [to_display(v) for v in value]
    return value


def short(value, limit=300) -> str:
    s = json.dumps(to_display(value), ensure_ascii=False)
    return s if len(s) <= limit else s[:limit] + "..."


def write_mismatch_csv(mismatched) -> Path:
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    out = Path(__file__).resolve().parent / f"mismatched_source_ids_{stamp}.csv"
    with out.open("w", newline="", encoding="utf-8-sig") as f:   # BOM so Excel opens it cleanly
        writer = csv.writer(f)
        writer.writerow(["excelRow", "sourceId", "brand", "partnerIds", "reason",
                         "actual specialHours"])
        for row in mismatched:
            writer.writerow(row)
    return out


def read_start_row() -> int:
    """Excel row number to start from; just Enter means row 1 (everything)."""
    while True:
        try:
            text = input("Enter Excel row number to start from (press Enter for 1): ").strip()
        except EOFError:
            return 1
        if text == "":
            return 1
        if text.isdigit() and int(text) >= 1:
            return int(text)
        print("Please enter a positive whole number, or just press Enter.")


# --------------------------------------------------------------------------- #
def main():
    excel_path = find_excel_file()
    print(f"\nScanning Excel file: {excel_path.name}")
    store_ids = get_live_store_ids(excel_path)
    if not store_ids:
        sys.exit("No rows with Store Status = 'Live' were found.")
    print(f"Excel scanned successfully. {len(store_ids)} Live Store ID(s) found.")
    warn_duplicate_store_ids(store_ids)
    print()

    start_row = read_start_row()
    store_ids = [s for s in store_ids if s[0] >= start_row]
    if not store_ids:
        sys.exit(f"No 'Live' rows found at or after Excel row {start_row}.")
    brands = group_brands(store_ids)
    print(f"Starting from Excel row {start_row}: {len(store_ids)} Live store(s) "
          f"across {len(brands)} brand(s) to check.\n")

    curl_args = parse_curl(read_multiline("Enter the curl command:"))
    if not contains_source_ids(curl_args):
        sys.exit("The curl command does not contain a 'sourceIds' attribute.")
    partner_attr = find_partner_attr(curl_args)
    if not partner_attr:
        sys.exit(f"The curl command does not contain a "
                 f"{' / '.join(PARTNER_ATTR_NAMES)} attribute.")
    partner_quoted = partner_uses_quotes(curl_args, partner_attr)

    # Ask every brand's partner ID up front, once per brand
    partner_ids = read_partner_ids(brands, partner_attr)

    print(f"\nExpected: specialHours entry with date {EXPECTED_DATE}, "
          f"startTime {EXPECTED_START_TIME}, endTime {EXPECTED_END_TIME}")

    mismatched = []   # (row_no, store_id, brand, partner_ids, reason, actual specialHours)
    print(f"\nChecking {len(store_ids)} store(s)...\n")
    for i, (row_no, sid, brand) in enumerate(store_ids, 1):
        pids = partner_ids[brand.lower()]
        pids_text = ",".join(pids)
        prefix = (f"[{i}/{len(store_ids)}] row {row_no}, brand {brand}, "
                  f"{partner_attr}={pids_text}, sourceIds={sid}:")
        actual = ""
        body = ""
        try:
            args = substitute_partner_ids(curl_args, partner_attr, pids, partner_quoted)
            body = run_curl(substitute_source_id(args, sid))
            data = parse_json_exact(body)
        except subprocess.TimeoutExpired:
            reason = "request timed out"
        except RuntimeError as e:
            reason = f"curl error: {e}"
        except json.JSONDecodeError:
            reason = f"response is not JSON: {body[:200]!r}"
        else:
            values = find_special_hours(data)
            reason = None
            if not values:
                reason = "no 'specialHours' attribute in response"
            for v in values:
                reason = check_special_hours(v)
                if reason:
                    actual = json.dumps(to_display(v), ensure_ascii=False)
                    break

        if reason:
            print(f"{prefix} MISMATCH ({reason})")
            mismatched.append((row_no, sid, brand, pids_text, reason, actual))
        else:
            print(f"{prefix} OK")

    print("\n" + "=" * 60)
    if mismatched:
        print(f"{len(mismatched)} sourceIds DID NOT MATCH the expected specialHours:")
        for row_no, sid, brand, _, reason, _ in mismatched:
            print(f"  row {row_no}: {sid} ({brand})    -> {reason}")
        csv_path = write_mismatch_csv(mismatched)
        print(f"\nMismatched IDs saved to: {csv_path}")
    else:
        print("All sourceIds matched the expected specialHours.")
    print("=" * 60)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nCancelled.")
