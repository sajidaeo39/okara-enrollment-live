# OKARA SIS - LATEST STUDENT ENROLLMENT / GENDER
# Simple Windows version
# Tehsil IDs: Depalpur=23, Okara=89, Renala Khurd=102

import sys, subprocess, importlib, re, time, shutil
from pathlib import Path
from datetime import datetime

# Install required packages automatically
for module, package in {
    "requests": "requests",
    "pandas": "pandas",
    "openpyxl": "openpyxl",
    "bs4": "beautifulsoup4",
}.items():
    try:
        importlib.import_module(module)
    except ImportError:
        print("Installing", package, "...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", package])

import requests
import pandas as pd
from bs4 import BeautifulSoup
from openpyxl import load_workbook

BASE = "https://sis.pesrp.edu.pk"
DISTRICT = "26"
TEHSILS = {
    "23": "Depalpur",
    "89": "Okara",
    "102": "Renala Khurd",
}
TIMEOUT = 45

UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/151.0.0.0 Safari/537.36"
)

HEADERS = {
    "Accept": "application/json, text/javascript, */*; q=0.01",
    "Accept-Language": "en-US,en;q=0.9",
    "X-Requested-With": "XMLHttpRequest",
    "User-Agent": UA,
    "Referer": BASE + "/",
}


def new_session():
    s = requests.Session()
    s.headers.update({"User-Agent": UA})
    try:
        r = s.get(BASE + "/", headers=HEADERS, timeout=TIMEOUT)
        print("SIS connection:", r.status_code)
    except Exception as e:
        print("Initial connection warning:", e)
    return s


def csrf(s):
    return s.cookies.get("csrf_cookie_name", "")


def request(s, endpoint, params):
    last = ""

    for attempt in range(1, 6):
        try:
            # Always refresh the CSRF value from the current session.
            params = dict(params)
            params["csrf_test_name"] = csrf(s)

            r = s.get(
                BASE + endpoint,
                params=params,
                headers=HEADERS,
                timeout=TIMEOUT,
            )

            if r.status_code == 403:
                last = "403 Forbidden"
                s = new_session()
                time.sleep(attempt)
                continue

            r.raise_for_status()

            text = r.text.strip()

            if not text.startswith("{"):
                last = "SIS returned HTML instead of JSON"
                time.sleep(attempt)
                continue

            return r.json(), s

        except Exception as e:
            last = str(e)
            print("   retry", attempt, "-", last[:100])
            time.sleep(min(attempt * 2, 10))

    raise RuntimeError(last)


def parse_options(data):
    soup = BeautifulSoup(data.get("html", ""), "html.parser")
    result = []
    seen = set()

    for option in soup.find_all("option"):
        value = str(option.get("value", "")).strip()
        text = " ".join(option.get_text(" ", strip=True).split())

        if value and value not in seen:
            seen.add(value)
            result.append((value, text))

    return result


def get_markazes(s, tehsil):
    data, s = request(
        s,
        "/user/get_markazes",
        {
            "tehsil": tehsil,
            "selectedMarkaz": "false",
            "all": "All",
        },
    )
    return parse_options(data), s


def get_schools(s, markaz):
    data, s = request(
        s,
        "/user/get_schools",
        {
            "markaz": markaz,
            "selectedSchool": "false",
            "all": "All",
        },
    )

    result = []

    for sid, text in parse_options(data):
        m = re.match(r"(\d+)\s*-\s*(.*)", text)

        if m:
            emis = m.group(1)
            name = m.group(2).strip()
        else:
            emis = ""
            name = text

        result.append((sid, emis, name))

    return result, s


def get_enrollment(s, tehsil, markaz, school):
    # THIS IS THE ENROLLMENT/GENDER API.
    # Attendance API is NOT used.
    data, s = request(
        s,
        "/dashboard_revamp/get_gender_summary_pie",
        {
            "district": DISTRICT,
            "tehsil": tehsil,
            "markaz": markaz,
            "school": school,
            "s_id_emis_code": "",
        },
    )

    keys = ("male_count", "female_count", "other_count", "total")

    if not any(k in data for k in keys):
        raise RuntimeError("Gender enrollment fields not returned by SIS")

    def n(k):
        value = data.get(k, 0)
        digits = re.sub(r"[^\d-]", "", str(value))
        return int(digits or 0)

    return {
        "Male": n("male_count"),
        "Female": n("female_count"),
        "Other": n("other_count"),
        "Total": n("total"),
    }, s


# Confirmed classification supplied for this report.
SECONDARY_SCHOOL_IDS = {'53040', '53615', '53041', '53644', '53587', '53614', '52233', '52245', '53595', '53554', '53586', '53072', '52229', '53045', '53608', '52235', '52242'}
EXCLUDED_SCHOOL_IDS = {"55048"}
EXCLUDED_EMIS = {"39399999"}


def wing(markaz, school, school_id=""):
    x = (markaz + " " + school).upper()

    # Confirmed male markaz.
    if "JOYIA" in markaz.upper():
        return "Male Wing"

    # All High / Higher Secondary schools are Secondary Wing.
    if (
        str(school_id).strip() in SECONDARY_SCHOOL_IDS
        or "HIGHER SECONDARY" in x
        or "HIGH SCHOOL" in x
        or re.search(r"\bHSS\b", x)
        or re.search(r"\bGHSS\b", x)
        or re.search(r"\bGHS\b", x)
        or re.search(r"\bGGHS\b", x)
    ):
        return "Secondary Wing"

    if "FEMALE" in x or "GIRLS" in x or re.search(r"\bGG[A-Z]", school.upper()):
        return "Female Wing"

    if "MALE" in x or "BOYS" in x or re.search(r"\bGM[A-Z]|\bGB[A-Z]", school.upper()):
        return "Male Wing"

    # Keep normal primary/elementary schools classified by their markaz/school
    # rather than putting them in Needs Review where possible.
    if re.search(r"\b(GGPS|GGES|GMPS|GMES)\b", school.upper()):
        if re.search(r"\bGG", school.upper()):
            return "Female Wing"
        return "Male Wing"

    # Final fallback: classify by markaz name when school name is ambiguous.
    if "FEMALE" in markaz.upper() or "GIRLS" in markaz.upper():
        return "Female Wing"
    if "MALE" in markaz.upper() or "BOYS" in markaz.upper():
        return "Male Wing"
    return "Male Wing"


def safe_int_series(df, col):
    return pd.to_numeric(df[col], errors="coerce").fillna(0)


def summary(df, groups):
    if df.empty:
        return pd.DataFrame()

    d = df.copy()

    for c in ["Male Enrollment", "Female Enrollment", "Other Enrollment", "Total Enrollment"]:
        d[c] = safe_int_series(d, c)

    return (
        d.groupby(groups, dropna=False)
        .agg(
            Schools=("School ID", "nunique"),
            Male_Enrollment=("Male Enrollment", "sum"),
            Female_Enrollment=("Female Enrollment", "sum"),
            Other_Enrollment=("Other Enrollment", "sum"),
            Total_Enrollment=("Total Enrollment", "sum"),
        )
        .reset_index()
    )


def write_excel(path, school_df, school_list_df, errors_df):
    # Write a genuine XLSX.
    with pd.ExcelWriter(path, engine="openpyxl", mode="w") as writer:
        school_df.to_excel(writer, sheet_name="School Wise", index=False)

        summary(
            school_df,
            ["Tehsil", "Markaz", "Markaz ID", "Wing"],
        ).to_excel(writer, sheet_name="Markaz Wise", index=False)

        summary(
            school_df,
            ["Tehsil", "Wing"],
        ).to_excel(writer, sheet_name="Tehsil Wise", index=False)

        summary(
            school_df,
            ["Wing"],
        ).to_excel(writer, sheet_name="Wing Summary", index=False)

        if not school_df.empty:
            total = pd.DataFrame([{
                "Male Enrollment": safe_int_series(
                    school_df, "Male Enrollment"
                ).sum(),
                "Female Enrollment": safe_int_series(
                    school_df, "Female Enrollment"
                ).sum(),
                "Other Enrollment": safe_int_series(
                    school_df, "Other Enrollment"
                ).sum(),
                "Total Enrollment": safe_int_series(
                    school_df, "Total Enrollment"
                ).sum(),
            }])
        else:
            total = pd.DataFrame([{
                "Male Enrollment": 0,
                "Female Enrollment": 0,
                "Other Enrollment": 0,
                "Total Enrollment": 0,
            }])

        total.to_excel(writer, sheet_name="Gender Summary", index=False)
        school_list_df.to_excel(writer, sheet_name="ALL School List", index=False)
        errors_df.to_excel(writer, sheet_name="Errors", index=False)

    # Validate XLSX before calling it final.
    wb = load_workbook(path, read_only=True)
    sheets = wb.sheetnames
    wb.close()

    if not sheets:
        raise RuntimeError("Excel workbook has no sheets")

    # Formatting pass.
    wb = load_workbook(path)

    for ws in wb.worksheets:
        ws.freeze_panes = "A2"

        if ws.max_row and ws.max_column:
            ws.auto_filter.ref = ws.dimensions

        for column in ws.columns:
            letter = column[0].column_letter
            longest = max(
                len(str(cell.value or ""))
                for cell in column
            )
            ws.column_dimensions[letter].width = min(
                max(10, longest + 2),
                45,
            )

    wb.save(path)
    wb.close()

    # Open again to ensure the saved file is readable.
    test = load_workbook(path, read_only=True)
    test.close()


def main():
    print()
    print("=" * 70)
    print(" OKARA - LATEST STUDENT ENROLLMENT / GENDER REPORT")
    print("=" * 70)
    print()
    print("1  Depalpur")
    print("2  Okara")
    print("3  Renala Khurd")
    print("4  ALL OKARA")
    print()

    # GitHub Actions runs the complete Okara district automatically.
    choice = "4"

    selected = list(TEHSILS.items())

    s = new_session()

    rows = []
    schools_all = []
    errors = []

    started = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    for tid, tname in selected:
        print()
        print("TEHSIL:", tname, "ID:", tid)

        try:
            marks, s = get_markazes(s, tid)
        except Exception as e:
            errors.append(["Tehsil", tname, "", "", str(e)])
            print("  Could not load Markaz:", e)
            continue

        print("  Markaz found:", len(marks))

        for mi, (mid, mname) in enumerate(marks, 1):
            print(f"  Markaz {mi}/{len(marks)}: {mname}")

            try:
                schools, s = get_schools(s, mid)
            except Exception as e:
                errors.append(["Markaz", tname, mname, "", str(e)])
                print("    School list failed:", e)
                continue

            print("    Schools:", len(schools))

            for sid, emis, sname in schools:
                # Completely exclude Cadet College Okara.
                if str(sid).strip() in EXCLUDED_SCHOOL_IDS or str(emis).strip() in EXCLUDED_EMIS:
                    print("    Skipping excluded school:", sname)
                    continue

                w = wing(mname, sname, sid)

                # One logical Secondary Wing per tehsil in the report.
                report_markaz = "SECONDARY-WING" if w == "Secondary Wing" else mname
                report_markaz_id = "SECONDARY-WING" if w == "Secondary Wing" else mid

                schools_all.append([
                    tname, tid, report_markaz, report_markaz_id,
                    sname, sid, emis, w
                ])

                try:
                    en, s = get_enrollment(
                        s, tid, mid, sid
                    )

                    rows.append([
                        started,
                        tname,
                        tid,
                        report_markaz,
                        report_markaz_id,
                        sname,
                        sid,
                        emis,
                        w,
                        en["Male"],
                        en["Female"],
                        en["Other"],
                        en["Total"],
                        "OK",
                    ])

                except Exception as e:
                    errors.append([
                        "School",
                        tname,
                        mname,
                        sname,
                        str(e),
                    ])

                    rows.append([
                        started,
                        tname,
                        tid,
                        mname,
                        mid,
                        sname,
                        sid,
                        emis,
                        w,
                        None,
                        None,
                        None,
                        None,
                        "FAILED",
                    ])

                    print("      FAILED:", sname[:55])

    school_columns = [
        "Run Date",
        "Tehsil",
        "Tehsil ID",
        "Markaz",
        "Markaz ID",
        "School",
        "School ID",
        "EMIS",
        "Wing",
        "Male Enrollment",
        "Female Enrollment",
        "Other Enrollment",
        "Total Enrollment",
        "Status",
    ]

    school_df = pd.DataFrame(rows, columns=school_columns)

    school_list_df = pd.DataFrame(
        schools_all,
        columns=[
            "Tehsil",
            "Tehsil ID",
            "Markaz",
            "Markaz ID",
            "School",
            "School ID",
            "EMIS",
            "Wing",
        ],
    )

    errors_df = pd.DataFrame(
        errors,
        columns=[
            "Type",
            "Tehsil",
            "Markaz",
            "School",
            "Error",
        ],
    )

    # GitHub-friendly output: one stable "latest" workbook plus a dated archive.
    # The dashboard reads the stable file; the dated file provides a daily history.
    stamp = datetime.now().strftime("%Y-%m-%d_%H%M%S")
    out_dir = Path(__file__).resolve().parent / "data"
    out_dir.mkdir(parents=True, exist_ok=True)
    final = out_dir / "Okara_Latest_Enrollment_Gender_Latest.xlsx"
    archive = out_dir / f"Okara_Latest_Enrollment_Gender_{stamp}.xlsx"
    temp = out_dir / f"creating_{stamp}.xlsx"

    try:
        write_excel(temp, school_df, school_list_df, errors_df)

        # Completeness gate: do not publish a partial workbook.
        discovered = len(school_list_df)
        rows_count = len(school_df)
        failed = int((school_df["Status"].astype(str).str.upper() != "OK").sum()) if not school_df.empty else 0
        if discovered == 0:
            raise RuntimeError("No schools were discovered from SIS.")
        if rows_count != discovered:
            raise RuntimeError(f"School row mismatch: discovered {discovered}, enrollment rows {rows_count}.")
        if failed:
            raise RuntimeError(f"{failed} school enrollment request(s) failed. Latest workbook was not published.")

        shutil.copy2(temp, archive)
        shutil.copy2(temp, final)
        temp.unlink(missing_ok=True)
    except Exception:
        if temp.exists():
            try:
                temp.unlink()
            except Exception:
                pass
        raise

    print()
    print("=" * 70)
    print("DONE")
    print("=" * 70)
    print("Excel file:")
    print(final)
    print()
    print("Schools found:", len(school_list_df))
    print("School rows:", len(school_df))
    print("Errors:", len(errors_df))
    print()
    print("Latest dashboard workbook:")
    print(final)
    print()
    # Do not wait for keyboard input in CI/GitHub Actions.
    # Local interactive runs can still pause before closing.
    if sys.stdin.isatty():
        input("Press Enter to close...")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print()
        print("ERROR:")
        print(e)
        raise
