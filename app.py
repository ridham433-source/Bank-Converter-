"""
BankConverter — upload a bank statement PDF, get a clean Excel/CSV file.
Single-file Streamlit app. No coding needed to run it — see README.md
for how to put this online for free.
"""
import difflib
import io
import os
import re
import tempfile
import uuid
import zipfile
from pathlib import Path

import pandas as pd
import pdfplumber
import pikepdf
import streamlit as st

# ---------------------------------------------------------------------------
# BETA ARCHIVING (currently DISABLED)
# Saves a copy of every uploaded statement -- success or failure -- so
# formats that don't convert yet can be used later to add support for more
# banks. Runs before any OCR/parsing/password handling, so a copy survives
# even if everything downstream fails or raises an exception.
# NOTE: on Streamlit Community Cloud the filesystem is ephemeral -- this
# folder is wiped whenever the app restarts/redeploys, so it's only a
# short-term beta capture, not permanent storage.
# Turned off now that extraction works reliably across banks -- flip
# ARCHIVE_ENABLED back to True (and restore the beta notices in the UI
# section below) if it's ever needed again for a new unsupported format.
# ---------------------------------------------------------------------------
ARCHIVE_ENABLED = False
ARCHIVE_DIR = Path(__file__).parent / "archived_statements"


def archive_uploaded_file(uploaded_file):
    """Save an exact copy of an uploaded file under a unique name, without
    ever overwriting an existing archived file. Never raises -- on failure
    it shows a soft warning and lets the app continue normally."""
    try:
        ARCHIVE_DIR.mkdir(parents=True, exist_ok=True)
        original_name = os.path.basename(uploaded_file.name)
        stem, ext = os.path.splitext(original_name)
        dest = ARCHIVE_DIR / f"{stem}__{uuid.uuid4().hex[:12]}{ext}"
        uploaded_file.seek(0)
        with open(dest, "wb") as f:
            f.write(uploaded_file.read())
        uploaded_file.seek(0)  # reset pointer so normal processing can still read it
        return True
    except Exception as e:
        st.warning(f"Couldn't archive a debug copy of '{uploaded_file.name}' ({e}). Conversion will continue normally.")
        return False


def render_admin_panel():
    """Beta-only: lets the app owner retrieve archived statements as a ZIP
    download, since Streamlit Community Cloud has no file browser. Hidden
    behind a URL flag (?admin=true) plus a passcode, so ordinary users
    visiting the app never see or reach this."""
    st.markdown("---")
    st.subheader("🔐 Admin: Archived Statements")

    try:
        admin_secret = st.secrets.get("ADMIN_PASSCODE", None)
    except Exception:
        admin_secret = None
    if not admin_secret:
        st.info(
            "No admin passcode is set up yet. Go to your app's settings on "
            "Streamlit Community Cloud → Secrets, and add:\n\n"
            "`ADMIN_PASSCODE = \"choose-your-own-secret-word\"`"
        )
        return

    entered = st.text_input("Enter admin passcode", type="password", key="admin_passcode_input")
    if not entered:
        st.stop()
    if entered != admin_secret:
        st.error("Incorrect passcode.")
        st.stop()

    ARCHIVE_DIR.mkdir(parents=True, exist_ok=True)
    files = sorted(ARCHIVE_DIR.glob("*"))
    if not files:
        st.info("No archived statements yet.")
        return

    st.write(f"**{len(files)} archived file(s):**")
    for f in files:
        size_kb = f.stat().st_size / 1024
        st.write(f"- `{f.name}` ({size_kb:.1f} KB)")

    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for f in files:
            zf.write(f, arcname=f.name)
    zip_buf.seek(0)

    st.download_button(
        "⬇️ Download all as ZIP",
        data=zip_buf,
        file_name="archived_statements.zip",
        mime="application/zip",
    )

# ---------------------------------------------------------------------------
# EXTRACTION ENGINE
# (validated against real HDFC Bank and Saraswat Bank statements — every
#  extracted row's running balance was checked against the statement's own
#  stated balance with zero mismatches)
# ---------------------------------------------------------------------------

DATE_RE = re.compile(r'^\d{1,2}[/-]\d{1,2}[/-]\d{2,4}$')
MONTH_NAME_DATE_RE = re.compile(
    r'^\d{1,2}\s+[A-Za-z]{3,9}\.?\s+\d{2,4}$'  # e.g. "01 Jun 2026", "1 January 26"
)


def looks_like_date(text):
    """True if text (possibly multiple words joined) reads as a date,
    covering both numeric ("01/06/2026") and month-name ("01 Jun 2026")
    formats used by different banks. Tolerant of trailing punctuation
    (a common OCR artifact on scanned/photographed statements)."""
    t = text.strip().rstrip('.,;:')
    return bool(DATE_RE.match(t) or MONTH_NAME_DATE_RE.match(t))

CANONICAL_COLUMNS = ["date", "narration", "ref_no", "value_dt", "withdrawal", "deposit", "balance"]

DISPLAY_NAMES = {
    "date": "Date",
    "narration": "Narration",
    "ref_no": "Chq./Ref.No.",
    "value_dt": "Value Dt.",
    "withdrawal": "Withdrawal Amt.",
    "deposit": "Deposit Amt.",
    "balance": "Closing Balance",
}

HEADER_KEYWORDS = {
    "value_dt":  ["value date", "value"],
    "date":      ["date"],
    "narration": ["narration", "particular", "particulars", "description",
                  "transaction details", "remarks", "transaction remarks"],
    "ref_no":    ["chq", "cheque", "ref", "instrument", "instruments", "reference"],
    "withdrawal":["withdrawal", "debit", "paid out", "dr amount", " dr "],
    "deposit":   ["deposit", "credit", "paid in", "cr amount", " cr "],
    "balance":   ["balance", "total", "running balance", "closing balance"],
}

# Only the longer, more distinctive keyword per column -- used for fuzzy
# (edit-distance) matching as a fallback when OCR has corrupted a header
# label past what exact/substring matching can recognize (e.g. photographed
# statements). Short/generic words like "cr", "dr", "ref", "total" are
# deliberately excluded here: they're too easy to accidentally fuzzy-match
# against unrelated garbled text (a real failure mode found during testing --
# a scrambled fragment of "Closing" matched "credit" at a plausible-looking
# ratio). Keeping this list to distinctive multi-letter/multi-word keywords
# keeps false positives rare while still tolerating real OCR noise.
FUZZY_HEADER_KEYWORDS = {
    "value_dt":  ["value date"],
    "narration": ["narration", "particulars", "description", "transaction details", "remarks"],
    "ref_no":    ["cheque", "reference", "instrument"],
    "withdrawal":["withdrawal", "paid out"],
    "deposit":   ["deposit", "paid in"],
    "balance":   ["closing balance", "running balance"],
}
FUZZY_MATCH_THRESHOLD = 0.65


def classify_header_word(text):
    """Classify a (possibly multi-word) header label into a canonical column."""
    # keep letters/spaces only so things like "Debit (Rs.)" match cleanly
    t = re.sub(r'[^a-z ]', ' ', text.lower())
    t = re.sub(r'\s+', ' ', t).strip()
    padded = f" {t} "
    for key in ["value_dt", "date", "narration", "ref_no", "withdrawal", "deposit", "balance"]:
        for kw in HEADER_KEYWORDS[key]:
            kw_clean = kw.strip()
            if kw_clean == t or (len(kw_clean) > 2 and kw_clean in t) or f" {kw_clean} " in padded:
                return key
    return None


def classify_header_labels_text_only(labels):
    """Tiers 1+2 only (exact match, then fuzzy match) -- no positional
    guessing. Used to DETECT whether a candidate line actually is a header
    row, based purely on real text evidence. (Positional fallback is
    deliberately excluded here: if it were used for detection too, almost
    any line could superficially "look like" a header, since a leftmost/
    rightmost label would auto-qualify as date/balance regardless of its
    real content.)
    """
    assigned = [classify_header_word(lab['text']) for lab in labels]

    for i, lab in enumerate(labels):
        if assigned[i] is not None:
            continue
        t = re.sub(r'[^a-z ]', ' ', lab['text'].lower())
        t = re.sub(r'\s+', ' ', t).strip()
        if len(t) < 4:
            continue
        best_key, best_ratio = None, 0.0
        for key, kws in FUZZY_HEADER_KEYWORDS.items():
            for kw in kws:
                r = difflib.SequenceMatcher(None, kw, t).ratio()
                if r > best_ratio:
                    best_ratio, best_key = r, key
        if best_ratio >= FUZZY_MATCH_THRESHOLD:
            assigned[i] = best_key

    return assigned


def classify_header_labels(labels):
    """Classify every header label on a row ALREADY CONFIRMED to be the
    header (via classify_header_labels_text_only + find_header_row), in
    three tiers:

    1) exact/substring keyword match (classify_header_word) -- tried first for
       every label, so this never changes behaviour for clean digital PDFs.
    2) fuzzy text match -- tolerates OCR noise (e.g. "Narratiea" -> narration),
       only used for labels tier 1 couldn't classify at all.
    3) positional fallback -- on every bank statement layout seen so far the
       leftmost header column is the date and the rightmost is the balance,
       regardless of OCR quality. Used only as a last resort, only for
       'date'/'balance', and only if that column wasn't found anywhere in
       this row by tiers 1-2. Safe here specifically because this row has
       already been confirmed a real header by classify_header_labels_text_only.

    Returns a list of canonical column names (or None), same length/order as
    `labels`.
    """
    assigned = classify_header_labels_text_only(labels)

    if labels:
        order = sorted(range(len(labels)), key=lambda i: labels[i]['x0'])
        if 'date' not in assigned:
            left_i = order[0]
            if assigned[left_i] is None:
                assigned[left_i] = 'date'
        if 'balance' not in assigned:
            right_i = order[-1]
            if assigned[right_i] is None:
                assigned[right_i] = 'balance'

    return assigned


def merge_header_labels(line_words, gap=7, max_label_width=160):
    """
    Merge adjacent header words into full labels before classifying, so
    multi-word headers like "Transaction Details" or "Debit (Rs.)" are
    recognized as one label instead of failing word-by-word.

    max_label_width caps how wide a single merged label can grow: real
    bank-statement header labels (even multi-word ones like "Closing
    Balance" or "Chq./Ref.No.") are always well under this. The cap exists
    to stop a chain of noisy/duplicate words (e.g. from OCR on a
    photographed statement, or from merging two overlapping OCR passes)
    from silently fusing into one giant blob that would swallow the space
    of multiple real columns -- a much worse failure than simply not
    recognizing a label, since it can misclassify a whole neighboring
    column's data (e.g. deposits read as withdrawals) without any error.
    """
    if not line_words:
        return []
    ordered = sorted(line_words, key=lambda w: w['x0'])
    labels = []
    cur = {"text": ordered[0]['text'], "x0": ordered[0]['x0'], "x1": ordered[0]['x1']}
    for w in ordered[1:]:
        if w['x0'] - cur['x1'] <= gap and (w['x1'] - cur['x0']) <= max_label_width:
            cur['text'] += ' ' + w['text']
            cur['x1'] = max(cur['x1'], w['x1'])
        else:
            labels.append(cur)
            cur = {"text": w['text'], "x0": w['x0'], "x1": w['x1']}
    labels.append(cur)
    return labels


def cluster_words_by_line(words, tolerance=3):
    """Group words into visual lines, tolerant of small baseline jitter
    between words that are meant to be on the same header row. Compares
    each word to the PREVIOUS word's top (rolling), not a fixed anchor,
    so gradual drift across a row (common in scanned/OCR'd documents)
    doesn't fracture one row into several."""
    ordered = sorted(words, key=lambda w: w['top'])
    lines = []
    cur = []
    prev_top = None
    for w in ordered:
        if prev_top is None or abs(w['top'] - prev_top) <= tolerance:
            cur.append(w)
        else:
            lines.append(cur)
            cur = [w]
        prev_top = w['top']
    if cur:
        lines.append(cur)
    return [(sum(w['top'] for w in ln) / len(ln), ln) for ln in lines]


def find_header_row(words):
    lines = cluster_words_by_line(words, tolerance=3)

    def is_valid_header(line_words):
        labels = merge_header_labels(line_words)
        classified = classify_header_labels_text_only(labels)
        has_date = 'date' in classified
        has_narr = 'narration' in classified
        has_amt = any(c in ('withdrawal', 'deposit', 'balance') for c in classified)
        # Require date + narration + an amount-type column, all via real text
        # evidence. This is deliberately strict: bank statements often have
        # an unrelated "Account Summary" block before the real transaction
        # table, and that block can coincidentally contain words that match
        # narration/amount keywords (e.g. "Nomination" fuzzy-matching
        # "narration", or a "Fixed Deposits" column literally containing
        # "deposit") -- requiring a real date match alongside them is what
        # actually distinguishes the transaction header from that noise,
        # since a summary block never has a date column.
        return has_date and has_narr and has_amt

    # A header row's words can be scattered across several clustered "lines"
    # on a photographed/skewed statement (baseline jitter exceeds the tight
    # tolerance used for body-row splitting). Try growing windows of
    # consecutive lines -- starting from a single line and extending -- so
    # header fragments spread across a wider vertical band still get
    # combined, without loosening the tolerance used elsewhere for accuracy.
    # Prefer the smallest possible header window, searched across the WHOLE
    # document at each size before growing wider. This matters: trying
    # "start here and grow forward" per starting line (the old approach)
    # let an early messy multi-line combination -- e.g. an unrelated
    # Account Summary block that happens to sit within reach of the real
    # header a few lines later -- win before the real header was ever
    # tried alone. Checking every single line first, then every 2-line
    # combination, etc., means a clean single-line header always wins.
    MAX_HEADER_SPAN = 70  # points; header fragments rarely spread wider than this
    MAX_WINDOW_LINES = 5
    for window_size in range(1, MAX_WINDOW_LINES + 1):
        for i in range(len(lines)):
            j = i + window_size - 1
            if j >= len(lines):
                break
            if lines[j][0] - lines[i][0] > MAX_HEADER_SPAN:
                continue
            combined = []
            for k in range(i, j + 1):
                combined += lines[k][1]
            if is_valid_header(combined):
                return lines[j][0], combined

    return None, None


def is_combined_amount_label(text):
    """Detect a single merged Withdrawal/Deposit column, e.g. Kotak's
    'Withdrawal(Dr)/Deposit(Cr)' -- one physical column holding both types
    of amount, distinguished only by a (D)/(C) suffix on each value."""
    t = re.sub(r'[^a-z ]', ' ', text.lower())
    has_debit_word = any(kw in t for kw in ("withdrawal", "debit"))
    has_credit_word = any(kw in t for kw in ("deposit", "credit"))
    return has_debit_word and has_credit_word


def build_columns(header_words, page_width):
    labels = merge_header_labels(header_words)
    # This function only ever receives the words of a row find_header_row has
    # already confirmed to be the real header (via text-only classification),
    # so it's safe to use the full classifier here, including positional
    # fallback for date/balance.
    classified = classify_header_labels(labels)
    groups = {}
    for lab, col in zip(labels, classified):
        if col is None:
            continue
        g = groups.setdefault(col, {"x0": lab['x0'], "x1": lab['x1'], "text": lab['text']})
        g["x0"] = min(g["x0"], lab['x0'])
        g["x1"] = max(g["x1"], lab['x1'])
        g["text"] += ' ' + lab['text']

    if not groups:
        return None, None

    combined_amount_col = None
    for name, g in groups.items():
        if name in ('withdrawal', 'deposit') and is_combined_amount_label(g['text']):
            combined_amount_col = name

    ordered = sorted(groups.items(), key=lambda kv: kv[1]["x0"])
    boundaries = []
    for i, (name, g) in enumerate(ordered):
        left = 0 if i == 0 else (ordered[i-1][1]["x1"] + g["x0"]) / 2
        right = page_width if i == len(ordered) - 1 else (g["x1"] + ordered[i+1][1]["x0"]) / 2
        boundaries.append((name, left, right))
    return boundaries, combined_amount_col


def calibrate_date_narration_boundary(boundaries, body_words):
    """Header-label position is unreliable for the Date/Narration split
    specifically (the 'Narration' label sits far right of where narration
    text actually starts) -- recalibrate using real body text.

    Uses the actual gap between date-column text and the next word on the
    same line (rather than a fixed buffer), since that gap can be very
    small on tightly-spaced scanned/photographed statements."""
    date_tokens = [w for w in body_words if DATE_RE.match(w['text'])]
    if not date_tokens:
        return boundaries
    date_tokens.sort(key=lambda w: w['x0'])
    leftmost_x0 = date_tokens[0]['x0']
    date_col_tokens = [w for w in date_tokens if w['x0'] < leftmost_x0 + 15]
    if len(date_col_tokens) < 3:
        return boundaries
    date_max_x1 = max(w['x1'] for w in date_col_tokens)

    # for each date token, find the nearest word to its right on the same
    # line -- that word's x0 is the real narration start for that row
    next_word_x0s = []
    for dt in date_col_tokens:
        same_line = [w for w in body_words if abs(w['top'] - dt['top']) < 3 and w['x0'] > dt['x1']]
        if same_line:
            nearest = min(same_line, key=lambda w: w['x0'])
            next_word_x0s.append(nearest['x0'])

    if next_word_x0s:
        narration_min_x0 = min(next_word_x0s)
        gap = narration_min_x0 - date_max_x1
        # sit close to the date text's own right edge rather than the
        # midpoint -- a midpoint can still clip a narration token that
        # starts unusually close on one particular row (e.g. scanned docs)
        calibrated_right = date_max_x1 + min(3, gap * 0.3)
    else:
        calibrated_right = date_max_x1 + 3

    new_boundaries = []
    for (name, left, right) in boundaries:
        if name == 'date':
            new_boundaries.append((name, left, calibrated_right))
        elif name == 'narration':
            new_boundaries.append((name, calibrated_right, right))
        else:
            new_boundaries.append((name, left, right))
    return new_boundaries


def looks_like_amount(text):
    """True if a token looks like a currency value (allowing the common
    OCR artifacts we normalize elsewhere: hyphen-as-decimal, (D)/(C)
    suffixes). Used to catch narration fragments that drift into an
    amount column due to per-page boundary drift on scanned documents."""
    t = text.strip()
    return bool(re.match(r'^-?[\d,]+([.\-]\d{1,2})?\s*\(?[A-Za-z]{0,2}\)?$', t))


def looks_like_amount(text):
    """True if a token looks like a currency value (allowing the common
    OCR artifacts we normalize elsewhere: hyphen-as-decimal, (D)/(C)
    suffixes). Used to catch narration fragments that drift into an
    amount column due to per-page boundary drift on scanned documents."""
    t = text.strip()
    return bool(re.match(r'^-?[\d,]+([.\-]\d{1,2})?\s*\(?[A-Za-z]{0,2}\)?$', t))


def assign_token(row, col, text):
    """Put a token in its classified column -- unless that column is an
    amount field and the token doesn't actually look like an amount, in
    which case it's almost certainly a narration fragment that drifted
    in from per-page boundary noise, so keep it with the narration."""
    if col in ('withdrawal', 'deposit') and not looks_like_amount(text):
        row['narration'].append(text)
    else:
        row[col].append(text)


def col_for_x(boundaries, x0):
    for name, left, right in boundaries:
        if left <= x0 < right:
            return name
    return None


def find_title_anchor(words):
    """Fallback anchor for continuation pages that repeat a statement
    title (e.g. 'Statement of account') but not the column header row."""
    for w in words:
        if w['top'] < 260 and w['text'].lower().strip('.:') in ('account', 'accounts', 'transactions'):
            same_line = [ww for ww in words if abs(ww['top'] - w['top']) < 2]
            joined = " ".join(ww['text'].lower() for ww in sorted(same_line, key=lambda ww: ww['x0']))
            if 'statement' in joined:
                return w['bottom']
    return None


def find_footer_top(words, page_height, header_top=None):
    # These are specific enough bank-statement footer phrases that
    # searching the whole page (not just the bottom half) is safe -- a
    # page that's mostly a closing summary can start well above the
    # midpoint. The one thing that must be excluded is the header row
    # itself: a "Closing Balance" COLUMN TITLE matches the same markers
    # we're looking for in the footer, so skip anything near header_top.
    markers = ["hdfcbanklimited", "statementsummary", "endofstatement",
               "totals", "generatedon", "closingbalance"]
    page_of_re = re.compile(r'page\s*\d+\s*of\s*\d+')
    candidates = []
    by_line = {}
    for w in words:
        if header_top is not None and (w['top'] <= header_top or abs(w['top'] - header_top) < 20):
            continue
        key = round(w['top'])
        by_line.setdefault(key, []).append(w)
    for top, line_words in by_line.items():
        joined = "".join(w['text'].lower() for w in sorted(line_words, key=lambda w: w['x0']))
        if page_of_re.search(joined):
            candidates.append(top)
            continue
        for m in markers:
            if m in joined:
                candidates.append(top)
    return min(candidates) if candidates else None


def get_grid_row_bands(page, table_top, table_bottom):
    lines = [l for l in page.lines if table_top - 2 <= l['top'] <= table_bottom + 2]
    tops = sorted(set(round(l['top'], 1) for l in lines))
    if len(tops) < 3:
        return None
    # Some layouts don't have a distinct ruled line right between the header
    # banner and the first data row (only the row-1/row-2 separator exists
    # within our range) -- without this, the entire first transaction row
    # silently has no band at all and gets dropped. If there's a gap after
    # table_top wide enough to plausibly hold a full row, anchor the first
    # band there instead of at the first internal line.
    if tops[0] - table_top > 10:
        tops = [table_top] + tops
    return [(tops[i], tops[i+1]) for i in range(len(tops)-1)]


def process_pages(pages):
    """
    Core row/column reconstruction, shared by both the PDF path (real
    pdfplumber pages) and the image/OCR path (synthetic pages built from
    Tesseract output). A "page" only needs: .width, .height, .lines (can be
    empty), and .extract_words(x_tolerance=...).
    """
    rows = []
    boundaries = None
    combined_amount_col = None
    prev_table_top = None
    header_page_size = None
    found_any_header = False
    current = None  # in-progress transaction; persists ACROSS pages so a
                     # narration that wraps across a physical page break
                     # (common in scanned multi-page passbooks) isn't lost

    for page in pages:
        words = page.extract_words(x_tolerance=1.5)
        header_top, header_words = find_header_row(words)

        if header_words:
            found_any_header = True
            new_boundaries, new_combined = build_columns(header_words, page.width)
            if new_boundaries:
                boundaries = new_boundaries
                combined_amount_col = new_combined
            table_top = header_top + 8
            header_page_size = (page.width, page.height)
        elif boundaries is not None:
            title_bottom = find_title_anchor(words)
            same_size_as_header_page = (
                header_page_size is not None
                and abs(page.width - header_page_size[0]) < header_page_size[0] * 0.05
                and abs(page.height - header_page_size[1]) < header_page_size[1] * 0.05
            )
            if title_bottom is not None:
                table_top = title_bottom + 4
            elif same_size_as_header_page and prev_table_top is not None:
                # true continuation page of a consistent multi-page PDF --
                # safe to reuse the offset from where the header was found
                table_top = prev_table_top
            else:
                # different page size (e.g. a separately cropped/scanned
                # photo) -- there's no reliable anchor, so don't guess a
                # foreign y-offset. Start near the very top instead.
                table_top = 2
        else:
            continue

        prev_table_top = table_top
        footer_top = find_footer_top(words, page.height, header_top=header_top)
        table_bottom = footer_top - 2 if footer_top else page.height - 20

        table_words = [w for w in words if table_top <= w['top'] <= table_bottom]
        if not table_words or boundaries is None:
            continue

        # recalibrate on EVERY page, not just pages with their own header --
        # each page may be an independently cropped photo whose columns
        # drift slightly from the page the boundaries were first built on.
        # Use a page-local copy so a noisy page can't corrupt the shared
        # boundaries carried forward to later pages.
        page_boundaries = calibrate_date_narration_boundary(boundaries, table_words)
        if header_words:
            boundaries = page_boundaries

        bands = get_grid_row_bands(page, table_top, table_bottom)

        if bands:
            for (top, bottom) in bands:
                band_words = [w for w in table_words if top <= w['top'] < bottom]
                if not band_words:
                    continue
                band_words.sort(key=lambda w: (w['top'], w['x0']))
                row = {c: [] for c in CANONICAL_COLUMNS}
                for w in band_words:
                    col = col_for_x(page_boundaries, w['x0'])
                    if col:
                        assign_token(row, col, w['text'])
                rows.append(row)
        else:
            for _avg_top, line_words in cluster_words_by_line(table_words, tolerance=4):
                line_words = sorted(line_words, key=lambda w: w['x0'])
                date_col_tokens = [w['text'] for w in line_words if col_for_x(page_boundaries, w['x0']) == 'date']
                starts_new = looks_like_date(' '.join(date_col_tokens))
                if starts_new or current is None:
                    if current is not None:
                        rows.append(current)
                    current = {c: [] for c in CANONICAL_COLUMNS}
                for w in line_words:
                    col = col_for_x(page_boundaries, w['x0'])
                    if col:
                        assign_token(current, col, w['text'])

    if current is not None:
        rows.append(current)

    if not found_any_header:
        return None  # could not identify a transaction table at all

    result = []
    for r in rows:
        if not looks_like_date(' '.join(r["date"])):
            continue
        flat = {c: " ".join(r[c]) for c in CANONICAL_COLUMNS}
        flat['date'] = flat['date'].strip().rstrip('.,;:')

        # OCR/scan artifact: a decimal point sometimes gets misread as a
        # hyphen (e.g. "1000-00" instead of "1000.00"). Only safe to fix
        # in amount fields, which should contain nothing but currency.
        for amt_col in ('withdrawal', 'deposit', 'balance'):
            flat[amt_col] = re.sub(r'(?<=\d)-(?=\d{2}\b)', '.', flat[amt_col])

        # A stray short numeral from the end of a wrapped narration line
        # (e.g. an address fragment like "...DO 3 MUMBAI...") can
        # occasionally land just past the narration/amount column boundary,
        # producing a malformed cell like "3,25,000.00 3". A bare token like
        # that can't be told apart from a genuine tiny amount by its text
        # alone (looks_like_amount would accept "3" either way) -- but every
        # real amount in these statements is properly formatted (comma
        # grouping and/or a decimal point), so when a cell has multiple
        # tokens, keep only the properly-formatted one and send any bare
        # stray token back to the narration instead of corrupting the amount.
        def _is_well_formed_amount(tok):
            return bool(re.match(r'^-?[\d,]*\d\.\d{1,2}$', tok))

        for amt_col in ('withdrawal', 'deposit', 'balance'):
            tokens = flat[amt_col].split()
            if len(tokens) > 1:
                well_formed = [t for t in tokens if _is_well_formed_amount(t)]
                stray = [t for t in tokens if not _is_well_formed_amount(t)]
                if well_formed:
                    flat[amt_col] = well_formed[0]
                    if stray:
                        flat['narration'] = (flat['narration'] + ' ' + ' '.join(stray)).strip()

        # Same bleed pattern, different shape: a transaction that's really
        # only a withdrawal (or only a deposit) sometimes ends up with a
        # lone bare stray token in the OTHER amount column instead of it
        # being empty -- e.g. withdrawal="2,500.00" alongside a phantom
        # deposit="3". If exactly one side is a real well-formed amount and
        # the other is a non-empty token that isn't, the second one is
        # narration bleed too.
        wd, dep = flat['withdrawal'], flat['deposit']
        if wd and dep:
            wd_ok, dep_ok = _is_well_formed_amount(wd), _is_well_formed_amount(dep)
            if wd_ok and not dep_ok:
                flat['narration'] = (flat['narration'] + ' ' + dep).strip()
                flat['deposit'] = ''
            elif dep_ok and not wd_ok:
                flat['narration'] = (flat['narration'] + ' ' + wd).strip()
                flat['withdrawal'] = ''

        if combined_amount_col:
            # single Withdrawal/Deposit column (e.g. "980.50(D)") -- route
            # into the correct canonical field using its (D)/(C) suffix
            raw = flat[combined_amount_col]
            other_col = 'deposit' if combined_amount_col == 'withdrawal' else 'withdrawal'
            if re.search(r'\(c\)|cr\b', raw, re.IGNORECASE):
                flat[other_col] = raw
                flat[combined_amount_col] = ''
            # else: leave it in combined_amount_col (already the debit-side field)

        result.append(flat)
    return result


def extract_transactions(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        return process_pages(list(pdf.pages))


# ---------------------------------------------------------------------------
# IMAGE / PHOTO SUPPORT (OCR)
# For photos of printed statements or passbooks. Accuracy depends heavily
# on photo quality -- a flat, well-lit, cropped scan will read far better
# than a tilted phone photo with background clutter.
# ---------------------------------------------------------------------------

class OCRPage:
    """Mimics just enough of a pdfplumber Page for process_pages() to work
    on OCR output: width, height, lines (always empty -- OCR gives us no
    vector graphics), and extract_words()."""
    def __init__(self, words, width, height):
        self._words = words
        self.width = width
        self.height = height
        self.lines = []

    def extract_words(self, x_tolerance=1.5):
        return self._words


def preprocess_image_for_ocr(pil_image):
    import cv2
    import numpy as np
    img = np.array(pil_image.convert("RGB"))[:, :, ::-1]  # RGB -> BGR
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

    # upscale small/low-res photos -- OCR does much better on larger text
    h, w = gray.shape
    if max(h, w) < 2200:
        scale = 2200 / max(h, w)
        gray = cv2.resize(gray, None, fx=scale, fy=scale, interpolation=cv2.INTER_CUBIC)

    _, otsu = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    return otsu


def ocr_image_to_words(pil_image):
    import pytesseract
    processed = preprocess_image_for_ocr(pil_image)
    ph, pw = processed.shape[:2]
    orig_w, orig_h = pil_image.size
    scale_x = orig_w / pw
    scale_y = orig_h / ph

    data = pytesseract.image_to_data(processed, config='--psm 6', output_type=pytesseract.Output.DICT)
    words = []
    n = len(data['text'])
    for i in range(n):
        text = data['text'][i].strip()
        if not text:
            continue
        try:
            conf = float(data['conf'][i])
        except (ValueError, TypeError):
            conf = -1
        if conf < 25:  # discard very low-confidence noise
            continue
        left = data['left'][i] * scale_x
        top = data['top'][i] * scale_y
        width = data['width'][i] * scale_x
        height = data['height'][i] * scale_y
        words.append({
            "text": text,
            "x0": left,
            "x1": left + width,
            "top": top,
            "bottom": top + height,
        })
    return words, orig_w, orig_h


def extract_transactions_from_images(pil_images):
    pages = []
    for img in pil_images:
        words, w, h = ocr_image_to_words(img)
        pages.append(OCRPage(words, w, h))
    return process_pages(pages)


# ---------------------------------------------------------------------------
# STREAMLIT UI
# ---------------------------------------------------------------------------

st.set_page_config(page_title="Bank2Excel", page_icon="🧾", layout="centered")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

:root {
    --ink: #1F2430;
    --ink-soft: #6B7280;
    --ink-faint: #9CA3AF;
    --bg: #F5F7FB;
    --card: #FFFFFF;
    --border: #E5E7EB;
    --border-soft: #EDEFF3;
    --accent: #4F63D2;
    --accent-soft: #EEF1FD;
    --gradient: linear-gradient(90deg, #3E6FE0 0%, #17B4A6 100%);
    --good: #17A567;
    --good-soft: #EAFBF3;
    --bad: #E0562F;
    --bad-soft: #FDEEE9;
    --purple: #7C5CE0;
    --purple-soft: #F1EEFC;
    --shadow-sm: 0 1px 2px rgba(20, 24, 41, 0.04);
    --shadow-md: 0 4px 14px rgba(31, 36, 48, 0.07);
    --shadow-btn: 0 2px 8px rgba(79, 99, 210, 0.28);
    --space-1: 0.5rem;
    --space-2: 1rem;
    --space-3: 1.5rem;
    --space-4: 2.25rem;
    color-scheme: light !important;
}

html { color-scheme: light !important; }
.stApp { background-color: var(--bg) !important; }
[data-testid="stHeader"] { background-color: transparent !important; }
html, body { font-family: 'Inter', sans-serif; }
[data-testid="stMarkdownContainer"], [data-testid="stMarkdownContainer"] p,
[data-testid="stMarkdownContainer"] h1, [data-testid="stMarkdownContainer"] h2, [data-testid="stMarkdownContainer"] h3,
[data-testid="stMarkdownContainer"] li {
    font-family: 'Inter', sans-serif !important;
    color: var(--ink) !important;
}

/* ---- page width: wider, still centered, never edge-to-edge ---- */
.block-container {
    padding-top: 1.6rem; padding-bottom: 2rem;
    max-width: 1200px !important;
}
@media (max-width: 1300px) { .block-container { max-width: 92vw !important; } }

/* Any wrapper we use purely as a CSS hook (via st.container(key=...)) must
   never itself show a visible box -- if it ever ends up with nothing but
   whitespace in it, it should be visually inert, not an empty card. */
[class*="st-key-"] { background: transparent; border: none; box-shadow: none; }

/* Streamlit's native file-uploader instruction text ("Drag and drop
   file here", size/type limits) isn't covered by our custom markdown,
   so it needs its own explicit color rule. */
[data-testid="stFileUploaderDropzoneInstructions"] span,
[data-testid="stFileUploaderDropzoneInstructions"] small,
[data-testid="stFileUploaderDropzoneInstructions"] div {
    color: var(--ink-soft) !important;
}

/* ---- navbar ---- */
.bc-navbar {
    display: flex; align-items: center; justify-content: space-between;
    padding: 0.5rem 0 1.4rem 0; margin-bottom: var(--space-3); border-bottom: 1px solid var(--border-soft);
}
.bc-navbar .brand { display: flex; align-items: center; gap: 0.55rem; font-size: 1.3rem; font-weight: 800; color: var(--ink); }
.bc-navbar .brand .icon { font-size: 1.35rem; line-height: 1; }
.bc-navbar .right { display: flex; align-items: center; gap: 0.75rem; }
.bc-secure {
    font-size: 0.82rem; color: var(--ink-soft); font-weight: 500;
    display: flex; align-items: center; gap: 0.35rem; padding: 0.4rem 0.2rem;
}
.bc-howitworks {
    font-size: 0.82rem; color: var(--ink-soft); font-weight: 500; border: 1px solid var(--border);
    border-radius: 999px; padding: 0.42rem 0.95rem; background: var(--card);
    transition: background 0.15s ease, border-color 0.15s ease, color 0.15s ease;
}
.bc-howitworks:hover { background: #FAFBFD; border-color: #D6DAE3; color: var(--ink); }

/* ---- hero ---- */
.bc-pill {
    display: inline-block; background: var(--accent-soft); color: var(--accent);
    font-size: 0.78rem; font-weight: 600; letter-spacing: 0.02em;
    padding: 0.32rem 0.95rem; border-radius: 999px; margin-bottom: 1.5rem;
}
.bc-hero { text-align: center; padding: 1rem 0 0.8rem 0; }
.bc-hero h1 {
    font-size: 2.5rem; font-weight: 800; margin: 0; line-height: 1.28;
    letter-spacing: -0.02em; color: var(--ink) !important;
}
.bc-hero h1.grad {
    background: var(--gradient); -webkit-background-clip: text; background-clip: text;
    -webkit-text-fill-color: transparent; margin-bottom: 1rem;
}
.bc-hero p { color: var(--ink-soft) !important; font-size: 1.05rem; margin-top: 0.5rem; }

/* ---- white cards (each is a real st.container(key=...) wrapper, so
   styling always nests correctly -- no risk of an empty/disconnected box) ---- */
.st-key-col_selector_card, .st-key-preview_card, .st-key-download_card {
    background: var(--card); border: 1px solid var(--border); border-radius: 16px;
    box-shadow: var(--shadow-sm);
    padding: 1.75rem 1.85rem 1.4rem 1.85rem; margin: var(--space-3) 0;
}
.st-key-col_selector_card h3, .st-key-preview_card h3, .st-key-download_card h3 {
    font-size: 1.08rem; font-weight: 700; color: var(--ink); margin: 0 0 1.1rem 0;
}

/* ---- upload card ---- */
.st-key-upload_card {
    background: var(--card); border: 1px solid var(--border); border-radius: 18px;
    box-shadow: var(--shadow-sm); padding: 2rem 2rem 1.6rem 2rem; margin: var(--space-3) 0 var(--space-2) 0;
}
.bc-upload-icon {
    width: 60px; height: 60px; border-radius: 50%; background: var(--accent-soft);
    display: flex; align-items: center; justify-content: center; margin: 0.2rem auto 1.1rem auto;
    font-size: 1.6rem;
}
.bc-upload-text { text-align: center; margin-bottom: 1.1rem; }
.bc-upload-text .main { font-size: 1.05rem; font-weight: 600; color: var(--ink); }
.bc-upload-text .sub { font-size: 0.85rem; color: var(--ink-soft); margin-top: 0.2rem; }
.bc-upload-text .sub a, .bc-upload-text .sub span.link { color: var(--accent); font-weight: 600; }
.bc-upload-hint { text-align: center; font-size: 0.78rem; color: var(--ink-faint); margin-top: 0.9rem; }

[data-testid="stFileUploaderDropzone"] {
    background-color: #FCFCFE !important;
    border: 1.5px dashed #D5DBEC !important;
    border-radius: 14px !important;
    min-height: 130px !important;
    padding: 1.3rem !important;
    flex-direction: column !important;
    align-items: center !important;
    justify-content: center !important;
    gap: 0.5rem !important;
    transition: border-color 0.15s ease, background-color 0.15s ease;
}
[data-testid="stFileUploaderDropzone"]:hover {
    border-color: var(--accent) !important;
    background-color: var(--accent-soft) !important;
}
/* Native "Browse files" button: give it visible light styling instead of
   Streamlit's default (which was rendering dark-on-dark in our theme). */
[data-testid="stFileUploaderDropzone"] button {
    background-color: #FFFFFF !important;
    color: var(--accent) !important;
    border: 1.5px solid #C7CEE8 !important;
    border-radius: 8px !important;
    transition: background-color 0.15s ease;
}
[data-testid="stFileUploaderDropzone"] button:hover {
    background-color: var(--accent-soft) !important;
}
[data-testid="stFileUploaderDropzone"] button svg {
    fill: var(--accent) !important;
}
/* Uploaded-file row (name, size, remove button) sits inside the same
   dropzone once a file is added — keep its icons/remove button visible
   and just fix text contrast, never hide/remove elements here. */
[data-testid="stFileUploaderDropzone"] small,
[data-testid="stFileUploaderDropzone"] span {
    color: var(--ink-soft) !important;
}

/* ---- format chips ---- */
.bc-chiprow-label { font-size: 0.85rem; font-weight: 600; color: var(--ink); margin-bottom: 0.7rem; }
.bc-chips { display: flex; gap: 0.6rem; flex-wrap: wrap; }
.bc-chip {
    display: flex; align-items: center; justify-content: center; gap: 0.4rem; border: 1.5px solid var(--border);
    border-radius: 10px; padding: 0.55rem 1rem; font-size: 0.85rem; font-weight: 500; color: var(--ink-soft);
    background: var(--card); min-width: 84px; min-height: 40px;
    transition: border-color 0.15s ease, background-color 0.15s ease, transform 0.1s ease;
}
.bc-chip:hover { border-color: #C7CEE8; transform: translateY(-1px); }
.bc-chip.selected { border-color: var(--accent); background: var(--accent-soft); color: var(--accent); font-weight: 700; }

.bc-privacy {
    text-align: center; color: var(--ink-soft); font-size: 0.8rem;
    margin: 0.6rem 0 var(--space-2) 0; display: flex; align-items: center; justify-content: center; gap: 0.3rem;
}

/* ---- stat cards: responsive CSS grid ---- */
.bc-stats {
    display: grid; grid-template-columns: repeat(4, 1fr); gap: 1rem; margin: var(--space-3) 0;
}
@media (max-width: 900px) { .bc-stats { grid-template-columns: repeat(2, 1fr); } }
@media (max-width: 520px) { .bc-stats { grid-template-columns: 1fr; } }
.bc-stat {
    background: var(--card); border: 1px solid var(--border); box-shadow: var(--shadow-sm);
    border-radius: 14px; padding: 1.2rem; display: flex; align-items: center; gap: 0.9rem;
    min-height: 84px;
}
.bc-stat .ic {
    width: 42px; height: 42px; border-radius: 50%; flex-shrink: 0;
    display: flex; align-items: center; justify-content: center; font-size: 1.15rem;
}
.bc-stat .ic.blue { background: var(--accent-soft); }
.bc-stat .ic.good { background: var(--good-soft); }
.bc-stat .ic.bad { background: var(--bad-soft); }
.bc-stat .txt { min-width: 0; }
.bc-stat .txt .label { font-size: 0.78rem; color: var(--ink-soft); margin-bottom: 0.15rem; }
.bc-stat .txt .value { font-size: 1.3rem; font-weight: 800; color: var(--ink); line-height: 1.2; }
.bc-stat .txt .value.good { color: var(--good); }
.bc-stat .txt .value.bad { color: var(--bad); }

/* ---- buttons: targeted via st.container(key=...) wrappers, which
   produce a REAL parent div (class st-key-<name>) around the widget --
   unlike markdown-opened divs, this reliably nests the button inside,
   so styling always applies. ---- */
.stButton>button, .stDownloadButton>button {
    border-radius: 10px !important; font-weight: 600 !important; padding: 0.6rem 1.1rem !important;
    transition: background-color 0.15s ease, border-color 0.15s ease, box-shadow 0.15s ease, transform 0.08s ease;
}

.st-key-convert_btn_wrap { display: flex; justify-content: center; margin: 1.5rem 0; }
.st-key-convert_btn_wrap .stButton>button {
    background-color: var(--accent) !important; color: #FFFFFF !important; border: none !important;
    min-width: 220px; height: 48px; font-size: 0.98rem !important;
    box-shadow: var(--shadow-btn) !important;
}
.st-key-convert_btn_wrap .stButton>button:hover {
    background-color: #3C4EB8 !important; box-shadow: 0 6px 16px rgba(79, 99, 210, 0.34) !important;
    transform: translateY(-1px);
}

.st-key-select_all_wrap .stButton>button, .st-key-clear_all_wrap .stButton>button {
    background: transparent !important; color: var(--accent) !important;
    border: none !important; box-shadow: none !important;
    font-weight: 600 !important; font-size: 0.85rem !important; padding: 0.2rem 0.4rem !important;
}
.st-key-select_all_wrap .stButton>button:hover, .st-key-clear_all_wrap .stButton>button:hover {
    text-decoration: underline; background: transparent !important;
}
.st-key-toggle_rows_wrap .stButton>button {
    background: var(--card) !important; color: var(--ink) !important; border: 1px solid var(--border) !important;
    box-shadow: none !important;
}
.st-key-toggle_rows_wrap .stButton>button:hover { background: #FAFBFD !important; border-color: #D6DAE3 !important; }

[data-testid="stCheckbox"] label p { font-size: 0.9rem !important; color: var(--ink) !important; }

/* ---- preview footer row ---- */
.note { font-size: 0.85rem; color: var(--ink-soft); }

/* ---- download buttons ---- */
.st-key-dl_excel_wrap .stDownloadButton>button {
    background: #FFFFFF !important; color: var(--good) !important;
    border: 1.5px solid var(--good) !important; box-shadow: none !important;
    height: 46px; width: 100%; font-size: 0.92rem !important;
}
.st-key-dl_excel_wrap .stDownloadButton>button:hover { background: var(--good-soft) !important; }
.st-key-dl_csv_wrap .stDownloadButton>button {
    background: #FFFFFF !important; color: var(--purple) !important;
    border: 1.5px solid var(--purple) !important; box-shadow: none !important;
    height: 46px; width: 100%; font-size: 0.92rem !important;
}
.st-key-dl_csv_wrap .stDownloadButton>button:hover { background: var(--purple-soft) !important; }

[data-testid="stDataFrame"] {
    border: 1px solid var(--border); border-radius: 12px; overflow: hidden; box-shadow: var(--shadow-sm);
}

.bc-footer {
    color: var(--ink-faint); font-size: 0.82rem; text-align: center; margin-top: var(--space-4);
    padding-top: 1.2rem; border-top: 1px solid var(--border-soft); opacity: 0.9;
}
</style>
""", unsafe_allow_html=True)

# ---- navbar ----
st.markdown("""
<div class="bc-navbar">
  <div class="brand"><span class="icon">📊</span>Bank2Excel</div>
  <div class="right">
    <div class="bc-secure">🛡️ 100% Secure</div>
    <div class="bc-howitworks">❓ How it works</div>
  </div>
</div>
""", unsafe_allow_html=True)

# ---- hero ----
st.markdown("""
<div class="bc-hero">
  <span class="bc-pill">Fast &bull; Secure &bull; Accurate</span>
  <h1>Convert Bank Statements</h1>
  <h1 class="grad">into Excel in Seconds</h1>
  <p>Clean, reliable, and private. Your data is never stored.</p>
</div>
""", unsafe_allow_html=True)

# ---- upload card ----
with st.container(key="upload_card"):
    st.markdown("""
    <div class="bc-upload-icon">⬆️</div>
    <div class="bc-upload-text">
      <div class="main">Drag &amp; drop your file here, or use Browse below</div>
    </div>
    """, unsafe_allow_html=True)

    # Note: the native uploader below already shows a working "Browse files"
    # button, a remove ("x") button per uploaded file, and the size/type
    # limit text — we style those natively instead of hiding/faking them,
    # so nothing here duplicates or breaks that built-in functionality.
    uploaded_files = st.file_uploader(
        "Upload PDF or Photos",
        type=["pdf", "png", "jpg", "jpeg"],
        accept_multiple_files=True,
        label_visibility="collapsed",
    )

    chip_col1, chip_col2 = st.columns(2)
    with chip_col1:
        st.markdown("""
        <div class="bc-chiprow-label">Input Format</div>
        <div class="bc-chips">
          <div class="bc-chip selected">📄 PDF</div>
          <div class="bc-chip">🖼️ Image</div>
          <div class="bc-chip">🖼️ JPG</div>
          <div class="bc-chip">🖼️ PNG</div>
        </div>
        """, unsafe_allow_html=True)
    with chip_col2:
        st.markdown("""
        <div class="bc-chiprow-label">Output Format</div>
        <div class="bc-chips">
          <div class="bc-chip selected">📗 Excel</div>
          <div class="bc-chip">📄 CSV</div>
        </div>
        """, unsafe_allow_html=True)


st.markdown('<div class="bc-privacy">🔒 We do not store your files. Your privacy is our priority.</div>', unsafe_allow_html=True)

if uploaded_files:
    if ARCHIVE_ENABLED:
        if "archived_file_ids" not in st.session_state:
            st.session_state["archived_file_ids"] = set()
        for _f in uploaded_files:
            _fid = getattr(_f, "file_id", None) or (_f.name, _f.size)
            if _fid not in st.session_state["archived_file_ids"]:
                if archive_uploaded_file(_f):
                    st.session_state["archived_file_ids"].add(_fid)

    pdfs = [f for f in uploaded_files if f.type == "application/pdf" or f.name.lower().endswith(".pdf")]
    images = [f for f in uploaded_files if f not in pdfs]

    if pdfs and images:
        st.error("Please upload either one PDF, or one or more photos — not both at once.")
        st.stop()
    if len(pdfs) > 1:
        st.error("Please upload one PDF at a time.")
        st.stop()

    file_signature = tuple((f.name, f.size) for f in uploaded_files)
    if st.session_state.get("last_file_signature") != file_signature:
        st.session_state["converted"] = False
        st.session_state["last_file_signature"] = file_signature

    st.markdown('<div style="height: 0.25rem;"></div>', unsafe_allow_html=True)
    with st.container(key="convert_btn_wrap"):
        convert_clicked = st.button("✅  Convert", type="primary")

    if "converted" not in st.session_state:
        st.session_state["converted"] = False
    if convert_clicked:
        st.session_state["converted"] = True

    if not st.session_state["converted"]:
        st.info("Uploaded — click **Convert** above to process this statement.")
        st.stop()

    rows = None

    if pdfs:
        uploaded = pdfs[0]
        pdf_bytes = uploaded.read()

        # --- handle password-protected PDFs ---
        is_encrypted = False
        try:
            with pikepdf.open(io.BytesIO(pdf_bytes)):
                pass
        except pikepdf.PasswordError:
            is_encrypted = True
        except Exception as e:
            st.error(f"Could not open this PDF: {e}")
            st.stop()

        working_bytes = pdf_bytes
        if is_encrypted:
            password = st.text_input("This PDF is password-protected. Enter the password:", type="password")
            if not password:
                st.info("Enter the password above to continue.")
                st.stop()
            try:
                with pikepdf.open(io.BytesIO(pdf_bytes), password=password) as pdf:
                    buf = io.BytesIO()
                    pdf.save(buf)
                    working_bytes = buf.getvalue()
            except pikepdf.PasswordError:
                st.error("That password didn't work. Please try again.")
                st.stop()

        with st.spinner("Reading your statement... this can take up to a minute for longer, multi-page statements"):
            with tempfile.NamedTemporaryFile(suffix=".pdf") as tmp:
                tmp.write(working_bytes)
                tmp.flush()
                rows = extract_transactions(tmp.name)

    else:
        st.info(
            "📷 Photo extraction is in beta — accuracy depends a lot on photo quality. "
            "For best results: use your phone's document **scan** mode (not a regular photo), "
            "make sure the page is flat, well-lit, and fills the frame."
        )
        from PIL import Image as PILImage
        pil_images = [PILImage.open(f) for f in images]
        with st.spinner("Reading your photo(s)... this can take a moment"):
            rows = extract_transactions_from_images(pil_images)

    if rows is None or len(rows) == 0:
        st.error(
            "Couldn't identify a transaction table. "
            "This can happen with some statement layouts (or unclear photos) we haven't seen yet — "
            "if you can share the file, it can be added to future support."
        )
        st.stop()

    df = pd.DataFrame(rows)
    df = df.rename(columns=DISPLAY_NAMES)
    df = df[[DISPLAY_NAMES[c] for c in CANONICAL_COLUMNS]]

    if images:
        st.warning("This was read from a photo — please double-check the numbers before relying on this file.")

    # --- stats ---
    def _num(s):
        s = str(s).replace(',', '').replace('CR', '').replace('DR', '').replace('(C)', '').replace('(D)', '').strip()
        try:
            return float(s)
        except ValueError:
            return 0.0

    total_withdrawal = df[DISPLAY_NAMES["withdrawal"]].apply(_num).sum()
    total_deposit = df[DISPLAY_NAMES["deposit"]].apply(_num).sum()
    net = total_deposit - total_withdrawal
    net_class = "good" if net >= 0 else "bad"

    st.markdown(f"""
    <div class="bc-stats">
      <div class="bc-stat">
        <div class="ic blue">📋</div>
        <div class="txt"><div class="value">{len(df)}</div><div class="label">Transactions</div></div>
      </div>
      <div class="bc-stat">
        <div class="ic bad">₹</div>
        <div class="txt"><div class="value">₹{total_withdrawal:,.2f}</div><div class="label">Total Withdrawn</div></div>
      </div>
      <div class="bc-stat">
        <div class="ic blue">⬇️</div>
        <div class="txt"><div class="value">₹{total_deposit:,.2f}</div><div class="label">Total Deposited</div></div>
      </div>
      <div class="bc-stat">
        <div class="ic good">📈</div>
        <div class="txt"><div class="value {net_class}">₹{net:,.2f}</div><div class="label">Net Change</div></div>
      </div>
    </div>
    """, unsafe_allow_html=True)

    # --- column selection ---
    with st.container(key="col_selector_card"):
        all_cols = list(df.columns)

        def _select_all():
            for c in all_cols:
                st.session_state[f"chk_{c}"] = True

        def _clear_all():
            for c in all_cols:
                st.session_state[f"chk_{c}"] = False

        header_col, link1_col, link2_col = st.columns([6, 1, 1])
        with header_col:
            st.markdown('<h3>Choose columns to include</h3>', unsafe_allow_html=True)
        with link1_col:
            with st.container(key="select_all_wrap"):
                st.button("Select All", on_click=_select_all, key="select_all_btn")
        with link2_col:
            with st.container(key="clear_all_wrap"):
                st.button("Clear All", on_click=_clear_all, key="clear_all_btn")

        selected = []
        checkbox_cols = st.columns(4)
        for i, col_name in enumerate(all_cols):
            with checkbox_cols[i % 4]:
                key = f"chk_{col_name}"
                if key not in st.session_state:
                    st.session_state[key] = True  # default: everything selected
                checked = st.checkbox(col_name, key=key)
                if checked:
                    selected.append(col_name)

    if not selected:
        st.warning("Select at least one column to see a preview and download.")
        st.stop()

    preview_df = df[selected]

    # --- preview ---
    with st.container(key="preview_card"):
        st.markdown('<h3>Preview</h3>', unsafe_allow_html=True)

        if "show_all_rows" not in st.session_state:
            st.session_state["show_all_rows"] = False

        rows_to_show = preview_df if st.session_state["show_all_rows"] else preview_df.head(5)
        st.dataframe(rows_to_show, width="stretch", hide_index=True)

        st.markdown('<div style="height: 0.4rem;"></div>', unsafe_allow_html=True)
        pf_col1, pf_col2 = st.columns([2, 1])
        with pf_col1:
            shown_count = len(preview_df) if st.session_state["show_all_rows"] else min(5, len(preview_df))
            st.markdown(f'<div class="note" style="padding-top: 0.5rem;">Showing {"all" if st.session_state["show_all_rows"] else "preview of"} {shown_count} rows</div>', unsafe_allow_html=True)
        with pf_col2:
            with st.container(key="toggle_rows_wrap"):
                toggle_label = "👁 Hide Extra Rows" if st.session_state["show_all_rows"] else "👁 View All Rows"
                if st.button(toggle_label, key="toggle_rows_btn", width="stretch"):
                    st.session_state["show_all_rows"] = not st.session_state["show_all_rows"]
                    st.rerun()

    # --- downloads ---
    with st.container(key="download_card"):
        st.markdown('<h3>Download your file</h3>', unsafe_allow_html=True)

        excel_buf = io.BytesIO()
        with pd.ExcelWriter(excel_buf, engine="openpyxl") as writer:
            preview_df.to_excel(writer, index=False, sheet_name="Statement")
        excel_buf.seek(0)

        csv_bytes = preview_df.to_csv(index=False).encode("utf-8")

        dl1, dl2 = st.columns(2)
        with dl1:
            with st.container(key="dl_excel_wrap"):
                st.download_button(
                    "📗  Download Excel",
                    data=excel_buf,
                    file_name="bank_statement.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    width="stretch",
                )
        with dl2:
            with st.container(key="dl_csv_wrap"):
                st.download_button(
                    "📄  Download CSV",
                    data=csv_bytes,
                    file_name="bank_statement.csv",
                    mime="text/csv",
                    width="stretch",
                )

st.markdown('<div class="bc-footer">🛡️ Bank2Excel — your files are processed for this session only.</div>', unsafe_allow_html=True)

if st.query_params.get("admin") == "true":
    render_admin_panel()
