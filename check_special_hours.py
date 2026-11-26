"""
check_special_hours.py

1. Finds the Excel file (.xlsx / .xlsm) sitting next to this script.
2. Scans every sheet for the "Store Status" and "Store ID" columns and keeps
   the Store IDs of rows whose Store Status is "Live".
   Then asks for the Excel row number to start from (Enter = row 1).
3. Asks for a curl command (paste it, then press Enter on an empty line).
4. Asks for the expected "specialHours" value (paste it, then empty line).
5. For each Store ID, puts it into the curl command's "sourceIds" (JSON body
   or URL query parameter), runs it, and checks that every "specialHours"
   in the JSON response is exactly equal to the expected value.
   Whitespace/indentation may differ, but structure, keys, value types,
   number formats and date/time strings must be identical (to the second).
6. Prints the Store IDs whose specialHours do not match and saves them to
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
LIVE_VALUE = "live"
HEADER_SEARCH_ROWS = 20      # how many top rows to search for the header row
CURL_TIMEOUT_SECONDS = 60


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


def get_live_store_ids(excel_path: Path) -> list[tuple[int, str]]:
    """Return (excel_row_number, store_id) for every 'Live' row."""
    wb = load_workbook(excel_path, read_only=True, data_only=True)
    store_ids: list[tuple[int, str]] = []
    found_headers = False

    for ws in wb.worksheets:
        rows = ws.iter_rows(values_only=True)
        status_col = id_col = None
        rows_seen = 0

        # Scan all columns of the top rows to locate the header row
        for row in rows:
            rows_seen += 1
            headers = [normalize(c) for c in row]
            if STATUS_HEADER in headers and ID_HEADER in headers:
                status_col = headers.index(STATUS_HEADER)
                id_col = headers.index(ID_HEADER)
                break
            if rows_seen >= HEADER_SEARCH_ROWS:
                break

        if status_col is None:
            print(f"  Sheet '{ws.title}': 'Store Status'/'Store ID' columns not found, skipped.")
            continue

        found_headers = True
        sheet_count = 0
        # continues after the header row; row_no is the Excel row number
        for row_no, row in enumerate(rows, start=rows_seen + 1):
            if row is None or len(row) <= max(status_col, id_col):
                continue
            if normalize(row[status_col]) == LIVE_VALUE:
                sid = row[id_col]
                if sid is None or str(sid).strip() == "":
                    continue
                store_ids.append((row_no, cell_to_id(sid)))
                sheet_count += 1
        print(f"  Sheet '{ws.title}': {sheet_count} 'Live' store(s) found.")

    wb.close()
    if not found_headers:
        sys.exit("Could not find 'Store Status' and 'Store ID' columns in any sheet.")
    return store_ids


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


def read_expected_special_hours():
    """Ask until valid JSON is entered. Accepts an optional "specialHours": prefix."""
    while True:
        text = read_multiline('\nEnter the expected "specialHours" value (JSON):').strip()
        # Accept:  [...]   /   "specialHours": [...]   /   {"specialHours": [...]}
        text = re.sub(r'^"?specialHours"?\s*:\s*', "", text).rstrip(",").strip()
        try:
            value = parse_json_exact(text)
        except json.JSONDecodeError as e:
            print(f"Invalid JSON ({e}). Please enter it again.")
            continue
        if isinstance(value, dict) and list(value) == ["specialHours"]:
            value = value["specialHours"]
        return value


def write_mismatch_csv(mismatched) -> Path:
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    out = Path(__file__).resolve().parent / f"mismatched_source_ids_{stamp}.csv"
    with out.open("w", newline="", encoding="utf-8-sig") as f:   # BOM so Excel opens it cleanly
        writer = csv.writer(f)
        writer.writerow(["excelRow", "sourceId", "reason", "actual specialHours"])
        for row_no, sid, reason, actual in mismatched:
            writer.writerow([row_no, sid, reason, actual])
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
    print(f"Excel scanned successfully. {len(store_ids)} Live Store ID(s) found.\n")

    start_row = read_start_row()
    store_ids = [(r, sid) for r, sid in store_ids if r >= start_row]
    if not store_ids:
        sys.exit(f"No 'Live' rows found at or after Excel row {start_row}.")
    print(f"Starting from Excel row {start_row}: {len(store_ids)} Live store(s) to check.\n")

    curl_args = parse_curl(read_multiline("Enter the curl command:"))
    if not contains_source_ids(curl_args):
        sys.exit("The curl command does not contain a 'sourceIds' attribute.")

    expected = read_expected_special_hours()
    print(f"Expected specialHours: {short(expected, limit=1000)}")

    mismatched = []   # (store_id, reason, actual specialHours)
    print(f"\nChecking {len(store_ids)} store(s)...\n")
    for i, (row_no, sid) in enumerate(store_ids, 1):
        prefix = f"[{i}/{len(store_ids)}] row {row_no}, sourceIds={sid}:"
        actual = ""
        body = ""
        try:
            body = run_curl(substitute_source_id(curl_args, sid))
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
                diff = find_difference(v, expected)
                if diff:
                    reason = diff
                    actual = json.dumps(to_display(v), ensure_ascii=False)
                    break

        if reason:
            print(f"{prefix} MISMATCH ({reason})")
            mismatched.append((row_no, sid, reason, actual))
        else:
            print(f"{prefix} OK")

    print("\n" + "=" * 60)
    if mismatched:
        print(f"{len(mismatched)} sourceIds DID NOT MATCH the expected specialHours:")
        for row_no, sid, reason, _ in mismatched:
            print(f"  row {row_no}: {sid}    -> {reason}")
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
