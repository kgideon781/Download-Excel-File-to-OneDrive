"""Build stats_latest.json for the CARTA Our Impact dashboard.

Reads the daily-refreshed `_latest.xlsx` files in ./data/ and emits
`data/stats_latest.json` containing:

- `rows`: cleaned row-level data for fellows, postdocs, grants, and trainings
  (the page re-aggregates client-side for Power-BI-style cross-filtering).
- `measures_static`: numbers that are not derivable from `rows` (currently empty —
  every published measure is derivable).
- `meta`: source filenames, generated_at, and counts for quick sanity checks.

Run locally:  python build_stats.py
In CI:        invoked from .github/workflows/download-excel.yml after the
              download step succeeds.
"""
import datetime as dt
import json
import re
import sys
import warnings
from pathlib import Path

from collections import Counter

import openpyxl

warnings.filterwarnings("ignore", category=UserWarning)

DATA_DIR = Path(__file__).resolve().parent / "data"
OUT_FILE = DATA_DIR / "stats_latest.json"

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def read_sheet(xlsx_path: Path, sheet_name: str) -> list[dict]:
    """Return a list of dicts from the named sheet, keyed by trimmed header."""
    wb = openpyxl.load_workbook(xlsx_path, read_only=True, data_only=True)
    if sheet_name not in wb.sheetnames:
        raise KeyError(f"{xlsx_path.name}: sheet '{sheet_name}' missing. Have: {wb.sheetnames}")
    ws = wb[sheet_name]
    rows = ws.iter_rows(values_only=True)
    # Blank header cells get a positional name so their data survives: the
    # "Local and Institutional ToTs" sheet holds the training topic in column 9
    # with an empty header cell, and dropping it loses the whole topic breakdown.
    header = []
    for idx, h in enumerate(next(rows), start=1):
        name = str(h).strip() if h is not None else ""
        header.append(name or f"__col{idx}")
    out = []
    for row in rows:
        if all(v is None for v in row):
            continue
        out.append({h: v for h, v in zip(header, row) if h})
    wb.close()
    return out


def to_int(v, default=None):
    if v is None or v == "":
        return default
    if isinstance(v, (int, float)):
        return int(v)
    try:
        return int(str(v).strip())
    except (ValueError, TypeError):
        return default


def to_float(v, default=None):
    if v is None or v == "":
        return default
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).replace(",", "").strip())
    except (ValueError, TypeError):
        return default


def to_year(v):
    """Coerce a value to a 4-digit year int, or None."""
    if v is None or v == "":
        return None
    if isinstance(v, dt.datetime):
        return v.year
    if isinstance(v, dt.date):
        return v.year
    if isinstance(v, (int, float)):
        n = int(v)
        return n if 1990 <= n <= 2100 else None
    s = str(v).strip()
    m = re.search(r"(19|20)\d{2}", s)
    return int(m.group(0)) if m else None


def normalize_status(v: str | None) -> str | None:
    """Map various status strings to a canonical set: Completed / In progress / Terminated."""
    if not v:
        return None
    s = str(v).strip().lower()
    if "complete" in s or "defend" in s or "graduat" in s:
        return "Completed"
    if "progress" in s:
        return "In progress"
    if "terminat" in s:
        return "Terminated"
    UNMAPPED["status"][s] += 1
    return str(v).strip()


# Values that were not recognised during this build, reported at the end so a
# vocabulary change in a source workbook cannot pass unnoticed. The demographics
# file switched from Female/Male to Woman/Man in 2026 and the old rule, which
# only looked for a leading "f", silently mapped every woman to None: the
# published page showed 0 female fellows while still looking healthy.
UNMAPPED: dict[str, "Counter[str]"] = {
    "gender": Counter(),
    "status": Counter(),
}

_FEMALE = {"f", "female", "woman", "women", "girl"}
_MALE = {"m", "male", "man", "men", "boy"}


def normalize_gender(v: str | None) -> str | None:
    if not v:
        return None
    s = str(v).strip().lower()
    if s in _FEMALE:
        return "Female"
    if s in _MALE:
        return "Male"
    # Prefix fallback for decorated values such as "Female (F)". Test the female
    # spellings first, because "woman" ends in "man".
    if s.startswith(("female", "woman", "women")):
        return "Female"
    if s.startswith(("male", "man", "men")):
        return "Male"
    UNMAPPED["gender"][s] += 1
    return None


def normalize_intervention(v: str | None) -> str | None:
    if not v:
        return None
    s = str(v).strip().upper()
    for tag in ("JAS", "SW", "APAS", "GGWW"):
        if s == tag or s.startswith(tag + " ") or s.startswith(tag + "-"):
            return tag
    if "PHD" in s:
        return "PhD training"
    return str(v).strip()


# Intervention tag -> the label the published dashboard shows.
TOPIC_LABELS = {
    "APAS": "Institutional support",
    "SW": "Supervisory skills",
    "GGWW": "Grant Writing",
    "JAS": "PhD training",
    "PhD training": "PhD training",
}


def topic_label(intervention: str | None) -> str | None:
    if not intervention:
        return None
    return TOPIC_LABELS.get(intervention, intervention)


def normalize_training_type(v: str | None) -> str | None:
    """Map type-of-training labels to canonical values."""
    if not v:
        return None
    s = str(v).strip().lower()
    if "joint" in s:
        return "Joint training"
    if "common tot" in s or "common to t" in s:
        return "Common ToT"
    if "local tot" in s or "local to t" in s:
        return "Local ToT"
    if "local training" in s:
        return "Local training"
    if "general workshop" in s:
        return "General workshop"
    return str(v).strip()


def normalize_duty(v: str | None) -> str:
    """ToT facilitator vs participant. Reduce composite labels to primary role."""
    if not v:
        return "Participant"
    s = str(v).strip().lower()
    if "facilit" in s or "organizer" in s or "management" in s:
        return "Facilitator"
    return "Participant"


# ---------------------------------------------------------------------------
# Per-source extractors
# ---------------------------------------------------------------------------


# Demographics workbooks in preference order. Only the first is refreshed by
# download_delegated.py; "Cohort_1_11_..." is an orphan whose newest snapshot is
# 30 Apr 2026, and reading it silently froze the graduate count at 195 while the
# maintained file had already moved to 197. Both carry the same population
# (287 fellows, cohorts 1-12), so the names are historical labels, not scope.
DEMOGRAPHICS_CANDIDATES = (
    "Cohort_1_10_Demographics_latest.xlsx",
    "Cohort_1_11_Demographics_latest.xlsx",
)


def demographics_file() -> Path:
    """First demographics workbook that exists, by preference."""
    for name in DEMOGRAPHICS_CANDIDATES:
        candidate = DATA_DIR / name
        if candidate.exists():
            return candidate
    raise FileNotFoundError(
        "no demographics workbook found, looked for: "
        + ", ".join(DEMOGRAPHICS_CANDIDATES)
    )


def load_fellows() -> list[dict]:
    """Fellow-level demographics, from whichever workbook the pipeline maintains."""
    src = demographics_file()
    print(f"  Fellows source: {src.name}")
    rows = read_sheet(src, "Fellows")
    out = []
    for r in rows:
        unique_id = r.get("Unique ID")
        if not unique_id:
            continue
        # Skip the appended summary/lookup table at the bottom of the Fellows sheet:
        # those rows have a Unique ID (e.g. "Cohort", 1, 2, ..., "Terminated", "Deceased")
        # but no actual fellow name. A real fellow always has a first name or surname.
        if not (r.get("First Name") or r.get("Surname")):
            continue
        out.append({
            "id": str(unique_id),
            "gender": normalize_gender(r.get("Gender")),
            "cohort": to_int(r.get("Cohort")),
            "nationality": (r.get("Nationality") or "").strip() or None,
            "institution_employment": (r.get("Institution of employment at registration") or "").strip() or None,
            "institution_registration": (r.get("Institution of registration") or "").strip() or None,
            "year_admission": to_year(r.get("Year of admission into CARTA")) or to_year(r.get("Date of PhD registration")),
            "year_completion": to_year(r.get("Date of completion(Defended/Graduated)")),
            "status": normalize_status(r.get("Current PhD Status ( Completed/Defended/In Progress)")),
            "ttc_months": to_float(r.get("Time to completion since PhD registration (Months)")),
            # The published report's "Average of time to completion" card measures
            # from enrolment into CARTA, not from PhD registration.
            "ttc_months_carta": to_float(r.get("Time to completion since enrolling CARTA (Months)")),
            "promotion": (r.get("Promotion event") or "").strip() or None,
            "responsibilities": (str(r.get("Other responsibilities") or "").strip() or None),
            "pubs_at_enrollment": to_int(r.get("No of Publications at Enrollment")),
            "pubs_during_phd": to_int(r.get("No of Publications During PhD")),
            "pubs_after_phd": to_int(r.get("No of Publications after PhD")),
            "first_author_phd_pubs": to_int(r.get("1st Author PhD  Publications")),
            "last_author_phd_pubs": to_int(r.get("Last Author PhD Publications")),
            "first_author_grad_pubs": to_int(r.get("1st Author Graduate  Publications")),
            "last_author_grad_pubs": to_int(r.get("Last Author  Graduate Publications")),
            "pubs_for_graduation": to_int(r.get("No of Publications for Graduation")),
            "terminated_date": to_year(r.get("Terminated")),
            "funder": (str(r.get("Fellow Funder") or "").strip() or None),
            "jas_attended": sum(
                1 for k in ("Month/ Year JAS1", "Month/ Year JAS2", "Month/ Year JAS3", "Month/ Year JAS4")
                if r.get(k) not in (None, "")
            ),
        })
    return out


def load_postdocs() -> list[dict]:
    rows = read_sheet(DATA_DIR / "Postdocs_latest.xlsx", "Post Doc")
    out = []
    for r in rows:
        if not r.get("Unique ID") and not r.get("Name of Awardee"):
            continue
        out.append({
            "id": str(r.get("Unique ID") or ""),
            "sex": normalize_gender(r.get("Sex")),
            "institution_employment": (r.get("Institution of employment at the time of award") or "").strip() or None,
            "host_country": (r.get("Host Country") or "").strip() or None,
            "award_type": (r.get("Award Type") or "").strip() or None,
            "year_award": to_year(r.get("Year of Award")),
            "year_completion": to_year(r.get("Year of Completion")),
            "status": normalize_status(r.get("Status (Active/Completed")),
            "funder": (r.get("Funder") or "").strip() or None,
        })
    return out


def load_grants() -> list[dict]:
    rows = read_sheet(DATA_DIR / "Extra Grants_latest.xlsx", "Extra Grants")
    out = []
    for r in rows:
        amount = to_float(r.get("Total amount in $"))
        if amount is None:
            # Power BI sums by amount; drop rows we can't sum. Counts already preserved upstream.
            continue
        out.append({
            "sex": normalize_gender(r.get("Sex")),
            "cohort": to_int(r.get("Cohort Number")),
            "institution": (r.get("Institution of employment at registration") or "").strip() or None,
            "type": (r.get("Type of Grant") or "").strip() or None,
            "year": to_year(r.get("Year")) or to_year(r.get("Award date")),
            "amount_usd": amount,
            "funder": (r.get("Name of Funder") or "").strip() or None,
            "duration_months": to_int(r.get("Duration of Funding (in months)")),
        })
    return out


# Distinct people who appear in the training sheets, counted during the build.
# Names are available here but are deliberately not published in the JSON, so
# only the count travels. Populated by load_trainings().
TRAINED_INDIVIDUALS: set[str] = set()


def load_trainings():
    src = DATA_DIR / "Institutionalization_latest.xlsx"

    def _load(sheet_name, source_tag):
        rows = read_sheet(src, sheet_name)
        out = []
        for r in rows:
            if not r.get("Full Name") and not r.get("Intervention") and not r.get("__col9"):
                continue
            name = " ".join(str(r.get("Full Name") or "").split()).strip().lower()
            if name:
                TRAINED_INDIVIDUALS.add(name)
            raw_topic = r.get("Intervention")
            if raw_topic in (None, ""):
                raw_topic = r.get("__col9")   # blank-header topic column
            intervention = normalize_intervention(raw_topic)
            out.append({
                "source": source_tag,  # "carta" | "institutional"
                "intervention": intervention,
                "topic": topic_label(intervention),
                "type": normalize_training_type(r.get("Type of training")),
                "duty": normalize_duty(r.get("Principal duty at event (Participant/Facilitator)")),
                "sex": normalize_gender(r.get("Sex")),
                "year": to_year(r.get("Year")) or to_year(r.get("Event Start Date")),
                "institution": (r.get("Associated institution") or "").strip() or None,
                "is_carta_fellow": (str(r.get("CARTA Fellow (Yes/No)") or "").strip().lower().startswith("y")),
            })
        return out

    return _load("CARTA Organized", "carta") + _load("Local and Institutional ToTs", "institutional")


# Values that mean "nothing recorded" in the recognitions workbook.
_BLANK_FLAGS = {"", "none", "no", "n/a", "na", "0", "-", "nil"}


def _recorded(v) -> bool:
    return str(v or "").strip().lower() not in _BLANK_FLAGS


def load_recognitions() -> list[dict]:
    """Promotions, internal/external appointments and awards, one row per event.

    Source: "Fellows promotions, Recognitions and awards.xlsx", sheet
    "Recognitions". Rows without a Unique ID are CARTA directors rather than
    fellows and are dropped.

    The published cards count *distinct fellows*, not events ("Fellows have
    taken up ..."), so the page de-duplicates on `id`. Emitting rows rather
    than totals keeps the cards filterable by cohort, institution and gender.
    """
    src = DATA_DIR / "Recognitions_latest.xlsx"
    if not src.exists():
        print(f"  Recognitions: {src.name} not present, skipping")
        return []
    rows = read_sheet(src, "Recognitions")
    out = []
    for r in rows:
        uid = str(r.get("Unique ID") or "").strip()
        if not uid:
            continue
        out.append({
            "id": uid,
            "gender": normalize_gender(r.get("Gender")),
            "cohort": to_int(r.get("Cohort")),
            "institution_employment": (r.get("Home institution at registration") or "").strip() or None,
            "year": to_year(r.get("Year of Award")),
            "promotion": _recorded(r.get("Promotion")),
            "internal_appointment": _recorded(r.get("Extra responsibilities (Internal appointments)")),
            # NB: the source header really is "Recognition (Award" with an
            # unbalanced bracket. Match it exactly.
            "award": _recorded(r.get("Recognition (Award")),
            "external_appointment": _recorded(r.get("Impact on field (External appointment)")),
        })
    return out


def load_curricula_institutions(trainings: list[dict]) -> list[str]:
    """Derive list of institutions that have adopted CARTA curricula from institutional trainings."""
    from collections import Counter
    counts = Counter()
    for t in trainings:
        if t["source"] == "institutional" and t["institution"]:
            counts[t["institution"]] += 1
    # Threshold: at least 5 training-events at that institution to count as "adopted"
    return sorted(name for name, n in counts.items() if n >= 5)


# ---------------------------------------------------------------------------
# Aggregate sanity-check measures (for meta block only — page recomputes everything)
# ---------------------------------------------------------------------------


def quick_measures(fellows, postdocs, grants, trainings) -> dict:
    total_fellows = len(fellows)
    by_status = {}
    for f in fellows:
        by_status[f["status"] or "Unknown"] = by_status.get(f["status"] or "Unknown", 0) + 1
    by_gender = {}
    for f in fellows:
        by_gender[f["gender"] or "Unknown"] = by_gender.get(f["gender"] or "Unknown", 0) + 1
    ttcs = [f["ttc_months"] for f in fellows if f["ttc_months"]]
    avg_ttc = round(sum(ttcs) / len(ttcs), 1) if ttcs else None
    median_ttc = None
    if ttcs:
        s = sorted(ttcs)
        median_ttc = round(s[len(s) // 2] if len(s) % 2 else (s[len(s)//2 - 1] + s[len(s)//2]) / 2, 1)
    total_grants = round(sum(g["amount_usd"] for g in grants))
    pubs = sum((f["pubs_during_phd"] or 0) + (f["pubs_after_phd"] or 0) for f in fellows)
    trained_participations = sum(1 for t in trainings if t["duty"] == "Participant")
    completed_fellows = by_status.get("Completed", 0)
    retention_rate = round(((total_fellows - by_status.get("Terminated", 0)) / total_fellows) * 100, 1) if total_fellows else None

    # Training counts (by intervention × type × source)
    from collections import defaultdict
    training_counts = defaultdict(int)
    for t in trainings:
        if not t["intervention"] or not t["type"]:
            continue
        key = f"{t['source']}::{t['intervention']}::{t['type']}"
        training_counts[key] += 1
    jas_person_events = sum(f.get("jas_attended", 0) or 0 for f in fellows)

    return {
        "total_fellows": total_fellows,
        "completed": completed_fellows,
        "in_progress": by_status.get("In progress", 0),
        "terminated": by_status.get("Terminated", 0),
        "fellows_by_gender": by_gender,
        "avg_ttc_months": avg_ttc,
        "median_ttc_months": median_ttc,
        "retention_rate_pct": retention_rate,
        "total_postdocs": len(postdocs),
        "postdocs_completed": sum(1 for p in postdocs if p["status"] == "Completed"),
        "extra_grants_usd": total_grants,
        "extra_grants_count": len(grants),
        "peer_reviewed_articles": pubs,
        # Distinct people trained, versus the number of training attendances.
        "trained_individuals": len(TRAINED_INDIVIDUALS),
        "trained_participations": trained_participations,
        "jas_person_events": jas_person_events,
        "training_counts": dict(training_counts),
    }


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------


def main():
    print("Building stats_latest.json...")
    fellows = load_fellows()
    print(f"  Fellows: {len(fellows)}")
    postdocs = load_postdocs()
    print(f"  Postdocs: {len(postdocs)}")
    grants = load_grants()
    print(f"  Grants (with amounts): {len(grants)}")
    trainings = load_trainings()
    print(f"  Training records: {len(trainings)}")
    institutions = load_curricula_institutions(trainings)
    recognitions = load_recognitions()
    print(f"  Recognition records: {len(recognitions)}")
    print(f"  Curricula institutions: {len(institutions)}")

    measures = quick_measures(fellows, postdocs, grants, trainings)
    if recognitions:
        def _fellows_with(flag):
            return len({r["id"] for r in recognitions if r[flag]})
        measures["recognition_counts"] = {
            "responsibilities_in_institution": _fellows_with("internal_appointment"),
            "promoted_since_joining": _fellows_with("promotion"),
            "responsibilities_outside": _fellows_with("external_appointment"),
            "awards_and_recognition": _fellows_with("award"),
        }
    print(f"  Quick measures: total_fellows={measures['total_fellows']}, "
          f"completed={measures['completed']}, grants=${measures['extra_grants_usd']:,}")

    payload = {
        "generated_at": dt.datetime.now(dt.UTC).isoformat(timespec="seconds"),
        "source_files": {
            "fellows": demographics_file().name,
            "postdocs": "Postdocs_latest.xlsx",
            "grants": "Extra Grants_latest.xlsx",
            "institutionalization": "Institutionalization_latest.xlsx",
            "recognitions": "Recognitions_latest.xlsx",
        },
        "rows": {
            "fellows": fellows,
            "postdocs": postdocs,
            "grants": grants,
            "trainings": trainings,
            "curricula_institutions": [{"name": n} for n in institutions],
            "recognitions": recognitions,
        },
        "measures_static": {},
        "meta": {"summary": measures},
    }

    for field, counts in UNMAPPED.items():
        if counts:
            print(f"  WARNING: unrecognised {field} values (treated as unknown):")
            for value, n in counts.most_common():
                print(f"    {n:>6}x {value!r}")

    OUT_FILE.write_text(json.dumps(payload, ensure_ascii=False, default=str), encoding="utf-8")
    size_kb = OUT_FILE.stat().st_size / 1024
    print(f"Wrote {OUT_FILE} ({size_kb:.1f} KB)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
