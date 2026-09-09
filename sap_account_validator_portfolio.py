"""
SAP ACCOUNT MOVEMENT VALIDATOR
=================================

PORTFOLIO / DEMO VERSION -- all company-specific identifiers (paths,
company codes) have been replaced with environment variables or
generic placeholders. This is a sanitized copy of a production
script; no real business data is included.

PROBLEM
-------
A reconciliation team needed to validate, for a batch of account
identifier pairs read from a CSV, whether each account is active in
SAP (FBL5N) and specifically whether it has a movement of a given
type ("initial load"). Checking this manually meant opening FBL5N
and re-entering each identifier one by one -- hundreds of times per
batch.

WHAT THIS SCRIPT DOES
----------------------
For every pair of identifiers in the input CSV, it queries SAP and
derives three flags per identifier:

  1. is_active        -- 1 if the account has any movements at all.
  2. has_target_type   -- 1 if, filtering by the target movement type,
                          at least one row exists (the "initial load").
  3. has_other_only     -- 1 if is_active=1 AND has_target_type=0, i.e.
                          the account has movements, but none of them
                          are the initial load -- a candidate anomaly
                          worth reviewing manually.

The third flag requires no extra SAP query: it's derived purely from
the first two (active but missing the target type implies some other
movement type exists instead).

SAFETY
------
This script is READ-ONLY. It never confirms or saves anything back
to SAP -- it only navigates screens, applies view filters, and reads
what's on screen. There is no write path in this codebase.

CSV LAYOUT
----------
Column positions (0-indexed) are configurable via the COL_* constants
below. The default layout expected is:

  0 ID          | 1 NAME          | 2 Identifier1 | 3 IsActive1 |
  4 HasType1    | 5 HasOtherOnly1 | 6 Identifier2 | 7 IsActive2 |
  8 HasType2    | 9 HasOtherOnly2

REQUIREMENTS
------------
  pip install pywin32 python-dotenv --break-system-packages

Requires an already-open, logged-in SAP GUI session with GUI
Scripting enabled.
"""

import csv
import os
import time

import win32com.client


# =============================================================================
# CONFIGURATION -- populate via environment variables, never hardcode
# real company data here.
# =============================================================================

INPUT_CSV = os.environ.get("INPUT_CSV", r"C:\temp\validator_input.csv")
OUTPUT_CSV = os.environ.get("OUTPUT_CSV", r"C:\temp\validator_output.csv")
COMPANY_CODE = os.environ.get("SAP_COMPANY_CODE", "XXXX")
TARGET_MOVEMENT_TYPE = os.environ.get("TARGET_MOVEMENT_TYPE", "XX")

# 0-indexed column positions in the input CSV -- adjust if your sheet
# layout differs.
COL_IDENT1, COL_ACTIVE1, COL_TYPE1, COL_OTHER1 = 2, 3, 4, 5
COL_IDENT2, COL_ACTIVE2, COL_TYPE2, COL_OTHER2 = 6, 7, 8, 9


def connect_to_sap():
    """Attaches to the first already-open SAP GUI session on this machine."""
    sap_gui_auto = win32com.client.GetObject("SAPGUI")
    application = sap_gui_auto.GetScriptingEngine
    connection = application.Children(0)
    session = connection.Children(0)
    return session


def read_screen_text(session) -> str:
    """
    Recursively walks every text-bearing control under the 'usr' area
    and concatenates their text into a single string.

    This lets the caller search for phrases like "no data found"
    without depending on a fixed screen coordinate or control ID --
    useful for classic SAP list screens where layout can shift
    slightly between environments or SAP versions.
    """
    texts = []

    def walk(element):
        try:
            children = element.Children
        except Exception:
            children = None
        if children is not None and children.Count > 0:
            for i in range(children.Count):
                walk(children.ElementAt(i))
        else:
            try:
                texts.append(element.Text)
            except Exception:
                pass

    try:
        walk(session.findById("wnd[0]/usr"))
    except Exception:
        pass

    return " ".join(t for t in texts if t)


def check_identifier(session, identifier: str) -> dict:
    """
    Queries SAP for a single account identifier and returns its three
    derived flags. Returns all-zero if the identifier is blank or SAP
    reports it as not found.
    """
    result = {"is_active": 0, "has_target_type": 0, "has_other_only": 0}

    if not identifier or not identifier.strip():
        return result

    session.findById("wnd[0]/tbar[0]/okcd").text = "/nFBL5N"
    session.findById("wnd[0]").sendVKey(0)
    time.sleep(1)

    session.findById("wnd[0]/usr/ctxtDD_KUNNR-LOW").text = identifier
    session.findById("wnd[0]/usr/ctxtDD_BUKRS-LOW").text = COMPANY_CODE
    session.findById("wnd[0]").sendVKey(8)
    time.sleep(1)

    try:
        message_type = session.findById("wnd[0]/sbar").MessageType
    except Exception:
        message_type = "E"

    if message_type != "S":
        return result  # not found / inactive -- all flags stay 0

    result["is_active"] = 1

    # Apply the "movement type = TARGET_MOVEMENT_TYPE" view filter
    try:
        session.findById("wnd[0]/tbar[1]/btn[38]").press()
        time.sleep(0.5)
        session.findById(
            "wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW"
        ).text = TARGET_MOVEMENT_TYPE
        session.findById("wnd[1]").sendVKey(0)
        time.sleep(0.5)
    except Exception as e:
        print(f"  (could not apply filter: {e})")
        return result

    screen_text = read_screen_text(session)
    has_target_type = "no data found" not in screen_text.lower() and "no contiene datos" not in screen_text.lower()

    result["has_target_type"] = 1 if has_target_type else 0
    result["has_other_only"] = 0 if has_target_type else 1

    return result


def main():
    session = connect_to_sap()

    with open(INPUT_CSV, newline="", encoding="utf-8-sig") as f:
        rows = list(csv.reader(f))

    header, data_rows = rows[0], rows[1:]
    max_col_needed = max(COL_OTHER1, COL_OTHER2)

    for i, row in enumerate(data_rows):
        while len(row) <= max_col_needed:
            row.append("")

        name = row[1] if len(row) > 1 else ""

        identifier1 = row[COL_IDENT1]
        print(f"[{i + 1}/{len(data_rows)}] {name} | Identifier1: {identifier1}")
        try:
            result1 = check_identifier(session, identifier1)
        except Exception as e:
            print(f"    error checking identifier1: {e}")
            result1 = {"is_active": 0, "has_target_type": 0, "has_other_only": 0}
        row[COL_ACTIVE1], row[COL_TYPE1], row[COL_OTHER1] = (
            result1["is_active"], result1["has_target_type"], result1["has_other_only"]
        )
        print(f"    -> active={result1['is_active']} has_type={result1['has_target_type']} other_only={result1['has_other_only']}")

        identifier2 = row[COL_IDENT2]
        print(f"    Identifier2: {identifier2}")
        try:
            result2 = check_identifier(session, identifier2)
        except Exception as e:
            print(f"    error checking identifier2: {e}")
            result2 = {"is_active": 0, "has_target_type": 0, "has_other_only": 0}
        row[COL_ACTIVE2], row[COL_TYPE2], row[COL_OTHER2] = (
            result2["is_active"], result2["has_target_type"], result2["has_other_only"]
        )
        print(f"    -> active={result2['is_active']} has_type={result2['has_target_type']} other_only={result2['has_other_only']}")

        # Incremental save every 20 rows, so a crash mid-run doesn't
        # lose all progress.
        if (i + 1) % 20 == 0:
            with open(OUTPUT_CSV, "w", newline="", encoding="utf-8-sig") as f:
                writer = csv.writer(f)
                writer.writerow(header)
                writer.writerows(data_rows)
            print(f"    (progress saved: {i + 1}/{len(data_rows)} rows)")

    with open(OUTPUT_CSV, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow(header)
        writer.writerows(data_rows)

    print(f"\nDone. Output saved to {OUTPUT_CSV}")


if __name__ == "__main__":
    main()
