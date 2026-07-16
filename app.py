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
# BETA ARCHIVING (temporary, for beta testing only)
# Saves a copy of every uploaded statement -- success or failure -- so
# formats that don't convert yet can be used later to add support for more
# banks. Runs before any OCR/parsing/password handling, so a copy survives
# even if everything downstream fails or raises an exception.
# NOTE: on Streamlit Community Cloud the filesystem is ephemeral -- this
# folder is wiped whenever the app restarts/redeploys, so it's only a
# short-term beta capture, not permanent storage.
# ---------------------------------------------------------------------------
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
                  "details", "transaction details", "remarks", "transaction remarks"],
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
FUZZY_MATCH_THRESHOLD = 0.55


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
        has_narr = 'narration' in classified
        has_amt = any(c in ('withdrawal', 'deposit', 'balance') for c in classified)
        # Require narration + an amount-type column via real text evidence --
        # specific enough that a random body-text line won't coincidentally
        # match both. 'date' is deliberately not required here: OCR on
        # photographed statements sometimes drops the short word "Date"
        # entirely, and positional fallback (leftmost = date) safely covers
        # that once a row has already qualified as a real header this way.
        return has_narr and has_amt

    # A header row's words can be scattered across several clustered "lines"
    # on a photographed/skewed statement (baseline jitter exceeds the tight
    # tolerance used for body-row splitting). Try growing windows of
    # consecutive lines -- starting from a single line and extending -- so
    # header fragments spread across a wider vertical band still get
    # combined, without loosening the tolerance used elsewhere for accuracy.
    MAX_HEADER_SPAN = 70  # points; header fragments rarely spread wider than this
    for i, (top, line_words) in enumerate(lines):
        combined = list(line_words)
        last_top = top
        if is_valid_header(combined):
            return last_top, combined
        j = i + 1
        while j < len(lines) and lines[j][0] - top <= MAX_HEADER_SPAN:
            combined = combined + lines[j][1]
            last_top = lines[j][0]
            if is_valid_header(combined):
                return last_top, combined
            j += 1

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
        if header_top is not None and abs(w['top'] - header_top) < 20:
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
    --accent: #4F63D2;
    --accent-soft: #EEF1FD;
    --gradient: linear-gradient(90deg, #3E6FE0 0%, #17B4A6 100%);
    --good: #17A567;
    --good-soft: #EAFBF3;
    --bad: #E0562F;
    --bad-soft: #FDEEE9;
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

.block-container { padding-top: 1.2rem; max-width: 760px; }

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
    padding-bottom: 1.2rem; margin-bottom: 1.6rem; border-bottom: 1px solid var(--border);
}
.bc-navbar .brand { display: flex; align-items: center; gap: 0.5rem; font-size: 1.25rem; font-weight: 800; color: var(--ink); }
.bc-navbar .brand .icon { font-size: 1.3rem; }
.bc-navbar .right { display: flex; align-items: center; gap: 0.9rem; }
.bc-secure { font-size: 0.82rem; color: var(--ink-soft); font-weight: 500; display: flex; align-items: center; gap: 0.3rem; }
.bc-howitworks {
    font-size: 0.82rem; color: var(--ink); font-weight: 500; border: 1px solid var(--border);
    border-radius: 999px; padding: 0.32rem 0.85rem;
}

/* ---- hero ---- */
.bc-pill {
    display: inline-block; background: var(--accent-soft); color: var(--accent);
    font-size: 0.78rem; font-weight: 600; letter-spacing: 0.02em;
    padding: 0.3rem 0.9rem; border-radius: 999px; margin-bottom: 1.1rem;
}
.bc-hero { text-align: center; padding: 0.4rem 0 0.6rem 0; }
.bc-hero h1 {
    font-size: 2.35rem; font-weight: 800; margin: 0; line-height: 1.2;
    letter-spacing: -0.02em; color: var(--ink) !important;
}
.bc-hero h1.grad {
    background: var(--gradient); -webkit-background-clip: text; background-clip: text;
    -webkit-text-fill-color: transparent; margin-bottom: 0.6rem;
}
.bc-hero p { color: var(--ink-soft) !important; font-size: 1.02rem; margin-top: 0.3rem; }

/* ---- white cards ---- */
.bc-card-wrap {
    background: var(--card); border: 1px solid var(--border); border-radius: 16px;
    padding: 1.6rem 1.6rem 1.2rem 1.6rem; margin: 1.4rem 0;
}
.bc-card-wrap h3 { font-size: 1.05rem; font-weight: 700; color: var(--ink); margin: 0 0 1rem 0; }

/* ---- upload dropzone visuals ---- */
.bc-upload-icon {
    width: 56px; height: 56px; border-radius: 50%; background: var(--accent-soft);
    display: flex; align-items: center; justify-content: center; margin: 0.2rem auto 0.9rem auto;
    font-size: 1.5rem;
}
.bc-upload-text { text-align: center; margin-bottom: 0.6rem; }
.bc-upload-text .main { font-size: 1rem; font-weight: 600; color: var(--ink); }
.bc-upload-text .sub { font-size: 0.85rem; color: var(--ink-soft); margin-top: 0.15rem; }
.bc-upload-text .sub a, .bc-upload-text .sub span.link { color: var(--accent); font-weight: 600; }
.bc-upload-hint { text-align: center; font-size: 0.78rem; color: var(--ink-faint); margin-top: 0.6rem; }

[data-testid="stFileUploaderDropzone"] {
    background-color: #FCFCFE !important;
    border: 1.5px dashed #C7CEE8 !important;
    border-radius: 12px !important;
    min-height: 150px !important;
    padding: 1.2rem !important;
    flex-direction: column !important;
    align-items: center !important;
    justify-content: center !important;
    gap: 0.4rem !important;
}
/* Native "Browse files" button: give it visible light styling instead of
   Streamlit's default (which was rendering dark-on-dark in our theme). */
[data-testid="stFileUploaderDropzone"] button {
    background-color: #FFFFFF !important;
    color: var(--accent) !important;
    border: 1.5px solid #C7CEE8 !important;
    border-radius: 8px !important;
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
.bc-chiprow-label { font-size: 0.85rem; font-weight: 600; color: var(--ink); margin-bottom: 0.6rem; }
.bc-chips { display: flex; gap: 0.55rem; flex-wrap: wrap; }
.bc-chip {
    display: flex; align-items: center; gap: 0.4rem; border: 1.5px solid var(--border);
    border-radius: 10px; padding: 0.45rem 0.9rem; font-size: 0.85rem; font-weight: 500; color: var(--ink-soft);
    background: var(--card);
}
.bc-chip.selected { border-color: var(--accent); background: var(--accent-soft); color: var(--accent); font-weight: 700; }

.bc-privacy { text-align: center; color: var(--ink-soft); font-size: 0.85rem; margin: 1.1rem 0 0.4rem 0; }

/* ---- stat cards ---- */
.bc-stats { display: flex; gap: 0.9rem; flex-wrap: wrap; margin: 1.4rem 0; }
.bc-stat {
    flex: 1 1 150px; background: var(--card); border: 1px solid var(--border);
    border-radius: 14px; padding: 1.1rem 1.1rem; display: flex; align-items: center; gap: 0.8rem;
}
.bc-stat .ic {
    width: 40px; height: 40px; border-radius: 50%; flex-shrink: 0;
    display: flex; align-items: center; justify-content: center; font-size: 1.1rem;
}
.bc-stat .ic.blue { background: var(--accent-soft); }
.bc-stat .ic.good { background: var(--good-soft); }
.bc-stat .ic.bad { background: var(--bad-soft); }
.bc-stat .txt .label { font-size: 0.78rem; color: var(--ink-soft); margin-bottom: 0.1rem; }
.bc-stat .txt .value { font-size: 1.15rem; font-weight: 700; color: var(--ink); line-height: 1.2; }
.bc-stat .txt .value.good { color: var(--good); }
.bc-stat .txt .value.bad { color: var(--bad); }

/* ---- select all / clear all as text links ---- */
.bc-linkbtns .stButton>button {
    background: transparent !important; color: var(--accent) !important; border: none !important;
    font-weight: 600 !important; font-size: 0.85rem !important; padding: 0 !important; box-shadow: none !important;
}
.bc-linkbtns .stButton>button:hover { text-decoration: underline; background: transparent !important; }

[data-testid="stCheckbox"] label p { font-size: 0.9rem !important; color: var(--ink) !important; }

/* ---- preview footer row ---- */
.bc-preview-foot { display: flex; align-items: center; justify-content: space-between; margin-top: 0.7rem; }
.bc-preview-foot .note { font-size: 0.85rem; color: var(--ink-soft); }

/* ---- generic buttons ---- */
.stButton>button, .stDownloadButton>button {
    border-radius: 8px !important; font-weight: 600 !important; padding: 0.5rem 1rem !important;
}
.bc-primary-btn .stButton>button {
    background-color: var(--accent) !important; color: #FFFFFF !important; border: none !important;
}
.bc-primary-btn .stButton>button:hover { background-color: #3C4EB8 !important; }
.bc-ghost-btn .stButton>button {
    background: var(--card) !important; color: var(--ink) !important; border: 1px solid var(--border) !important;
}

/* ---- download cards ---- */
.bc-dl-excel .stDownloadButton>button {
    background: var(--good-soft) !important; color: var(--good) !important;
    border: 1.5px solid #BEEBD4 !important; width: 100%;
}
.bc-dl-excel .stDownloadButton>button:hover { background: #DEF7EA !important; }
.bc-dl-csv .stDownloadButton>button {
    background: var(--accent-soft) !important; color: var(--accent) !important;
    border: 1.5px solid #D7DEFA !important; width: 100%;
}
.bc-dl-csv .stDownloadButton>button:hover { background: #E3E8FB !important; }

[data-testid="stDataFrame"] { border: 1px solid var(--border); border-radius: 10px; overflow: hidden; }

.bc-footer {
    color: var(--ink-soft); font-size: 0.85rem; text-align: center; margin-top: 2.4rem;
    padding-top: 1.1rem; border-top: 1px solid var(--border);
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
  <p>Clean and reliable. Beta note: uploaded files may be temporarily saved for debugging.</p>
</div>
""", unsafe_allow_html=True)

# ---- upload card ----
st.markdown('<div class="bc-card-wrap">', unsafe_allow_html=True)
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
st.markdown('</div>', unsafe_allow_html=True)  # end bc-card-wrap

st.markdown('<div class="bc-privacy">⚠️ Beta notice: uploaded statements (including account details) may be saved on the server to help fix unsupported bank formats. Please avoid uploading a statement you\'re not comfortable sharing during this beta.</div>', unsafe_allow_html=True)

if uploaded_files:
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

    st.markdown('<div class="bc-primary-btn" style="text-align:center; margin: 1rem 0;">', unsafe_allow_html=True)
    convert_clicked = st.button("✅  Convert", use_container_width=False, type="primary")
    st.markdown('</div>', unsafe_allow_html=True)

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

        with st.spinner("Reading your statement..."):
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
    st.markdown('<div class="bc-card-wrap">', unsafe_allow_html=True)
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
        st.markdown('<div class="bc-linkbtns">', unsafe_allow_html=True)
        st.button("Select All", on_click=_select_all, key="select_all_btn")
        st.markdown('</div>', unsafe_allow_html=True)
    with link2_col:
        st.markdown('<div class="bc-linkbtns">', unsafe_allow_html=True)
        st.button("Clear All", on_click=_clear_all, key="clear_all_btn")
        st.markdown('</div>', unsafe_allow_html=True)

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
    st.markdown('</div>', unsafe_allow_html=True)  # end bc-card-wrap

    if not selected:
        st.warning("Select at least one column to see a preview and download.")
        st.stop()

    preview_df = df[selected]

    # --- preview ---
    st.markdown('<div class="bc-card-wrap">', unsafe_allow_html=True)
    st.markdown('<h3>Preview</h3>', unsafe_allow_html=True)

    if "show_all_rows" not in st.session_state:
        st.session_state["show_all_rows"] = False

    rows_to_show = preview_df if st.session_state["show_all_rows"] else preview_df.head(5)
    st.dataframe(rows_to_show, use_container_width=True, hide_index=True)

    st.markdown('<div class="bc-preview-foot">', unsafe_allow_html=True)
    pf_col1, pf_col2 = st.columns([2, 1])
    with pf_col1:
        shown_count = len(preview_df) if st.session_state["show_all_rows"] else min(5, len(preview_df))
        st.markdown(f'<div class="note">Showing {"all" if st.session_state["show_all_rows"] else "preview of"} {shown_count} rows</div>', unsafe_allow_html=True)
    with pf_col2:
        st.markdown('<div class="bc-ghost-btn">', unsafe_allow_html=True)
        toggle_label = "👁 Hide Extra Rows" if st.session_state["show_all_rows"] else "👁 View All Rows"
        if st.button(toggle_label, key="toggle_rows_btn", use_container_width=True):
            st.session_state["show_all_rows"] = not st.session_state["show_all_rows"]
            st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)  # end bc-preview-foot
    st.markdown('</div>', unsafe_allow_html=True)  # end bc-card-wrap

    # --- downloads ---
    st.markdown('<div class="bc-card-wrap">', unsafe_allow_html=True)
    st.markdown('<h3>Download your file</h3>', unsafe_allow_html=True)

    excel_buf = io.BytesIO()
    with pd.ExcelWriter(excel_buf, engine="openpyxl") as writer:
        preview_df.to_excel(writer, index=False, sheet_name="Statement")
    excel_buf.seek(0)

    csv_bytes = preview_df.to_csv(index=False).encode("utf-8")

    dl1, dl2 = st.columns(2)
    with dl1:
        st.markdown('<div class="bc-dl-excel">', unsafe_allow_html=True)
        st.download_button(
            "📗  Download Excel",
            data=excel_buf,
            file_name="bank_statement.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
        st.markdown('</div>', unsafe_allow_html=True)
    with dl2:
        st.markdown('<div class="bc-dl-csv">', unsafe_allow_html=True)
        st.download_button(
            "📄  Download CSV",
            data=csv_bytes,
            file_name="bank_statement.csv",
            mime="text/csv",
            use_container_width=True,
        )
        st.markdown('</div>', unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)  # end bc-card-wrap

st.markdown('<div class="bc-footer">🛡️ Bank2Excel — beta version: uploaded files may be retained temporarily to help improve support for more banks.</div>', unsafe_allow_html=True)

if st.query_params.get("admin") == "true":
    render_admin_panel()
