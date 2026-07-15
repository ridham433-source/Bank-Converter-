"""
BankConverter — upload a bank statement PDF, get a clean Excel/CSV file.
Single-file Streamlit app. No coding needed to run it — see README.md
for how to put this online for free.
"""
import io
import re
import tempfile

import pandas as pd
import pdfplumber
import pikepdf
import streamlit as st

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


def merge_header_labels(line_words, gap=7):
    """
    Merge adjacent header words into full labels before classifying, so
    multi-word headers like "Transaction Details" or "Debit (Rs.)" are
    recognized as one label instead of failing word-by-word.
    """
    if not line_words:
        return []
    ordered = sorted(line_words, key=lambda w: w['x0'])
    labels = []
    cur = {"text": ordered[0]['text'], "x0": ordered[0]['x0'], "x1": ordered[0]['x1']}
    for w in ordered[1:]:
        if w['x0'] - cur['x1'] <= gap:
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
        classified = [classify_header_word(lab['text']) for lab in labels]
        has_date = 'date' in classified
        has_narr = 'narration' in classified
        has_amt = any(c in ('withdrawal', 'deposit', 'balance') for c in classified)
        return has_date and has_narr and has_amt

    for i, (top, line_words) in enumerate(lines):
        if is_valid_header(line_words):
            return top, line_words

        # some headers wrap a label across two stacked lines (e.g.
        # "Transaction" / "Date") -- try combining with the next line too
        if i + 1 < len(lines):
            next_top, next_words = lines[i + 1]
            if next_top - top < 15:
                combined = line_words + next_words
                if is_valid_header(combined):
                    return top, combined

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
    groups = {}
    for lab in labels:
        col = classify_header_word(lab['text'])
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
    --bg: #FAFBFC;
    --card: #FFFFFF;
    --border: #E5E7EB;
    --accent: #4F63D2;
    --accent-soft: #EEF1FD;
    --good: #1F9D55;
    --bad: #D0402B;
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

/* Streamlit's native file-uploader instruction text ("Drag and drop
   file here", size/type limits) isn't covered by our custom markdown,
   so it needs its own explicit color rule. */
[data-testid="stFileUploaderDropzoneInstructions"] span,
[data-testid="stFileUploaderDropzoneInstructions"] small,
[data-testid="stFileUploaderDropzoneInstructions"] div {
    color: var(--ink-soft) !important;
}

.bc-hero { text-align: center; padding: 2.2rem 0 0.6rem 0; }
.bc-hero h1 { font-size: 2.3rem; font-weight: 800; margin-bottom: 0.5rem; letter-spacing: -0.02em; color: var(--ink) !important; }
.bc-hero p { color: var(--ink-soft) !important; font-size: 1.05rem; margin-top: 0; }

.bc-badgerow { text-align: center; margin: 0.35rem 0; }
.bc-badgerow .tag {
    display: inline-block; font-size: 0.72rem; color: var(--ink-soft);
    font-weight: 600; letter-spacing: 0.06em; margin-right: 0.5rem;
}
.bc-badge {
    display: inline-block; background: #F1F3F6; color: var(--ink);
    border: 1px solid var(--border); border-radius: 999px;
    padding: 0.2rem 0.75rem; font-size: 0.82rem; font-weight: 500;
    margin: 0.15rem 0.2rem;
}
.bc-badge.out { background: var(--accent-soft); color: var(--accent); border-color: #D7DEFA; }

.bc-divider { height: 1px; background: var(--border); margin: 1.6rem 0; }

.bc-privacy { text-align: center; color: var(--ink-soft); font-size: 0.85rem; margin-top: 0.4rem; }

[data-testid="stFileUploaderDropzone"] {
    background-color: var(--card) !important;
    border: 1.5px solid var(--border) !important;
    border-radius: 14px !important;
}

.bc-cards { display: flex; gap: 0.8rem; flex-wrap: wrap; margin: 1rem 0 1.6rem 0; }
.bc-card {
    flex: 1 1 150px; background: var(--card); border: 1px solid var(--border);
    border-radius: 10px; padding: 0.9rem 1.1rem;
}
.bc-card .label { font-size: 0.76rem; color: var(--ink-soft); text-transform: uppercase; letter-spacing: 0.04em; margin-bottom: 0.3rem; }
.bc-card .value { font-size: 1.3rem; font-weight: 700; color: var(--ink); }
.bc-card.good .value { color: var(--good); }
.bc-card.bad .value { color: var(--bad); }

.stButton>button, .stDownloadButton>button {
    background-color: var(--accent) !important; color: #FFFFFF !important;
    border-radius: 8px !important; border: none !important;
    font-weight: 600 !important; padding: 0.5rem 1rem !important;
}
.stButton>button:hover, .stDownloadButton>button:hover { background-color: #3C4EB8 !important; }

[data-testid="stDataFrame"] { border: 1px solid var(--border); border-radius: 10px; overflow: hidden; }

.bc-brand {
    display: flex; align-items: center; justify-content: center; gap: 0.5rem;
    font-size: 1.7rem; font-weight: 800; color: var(--accent) !important;
    letter-spacing: -0.01em; margin-bottom: 1.6rem;
}
.bc-brand .icon { font-size: 1.5rem; }

.bc-footer { color: var(--ink-soft); font-size: 0.85rem; text-align: center; margin-top: 3rem; padding-top: 1rem; border-top: 1px solid var(--border); }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="bc-hero">
  <div class="bc-brand"><span class="icon">📊</span>Bank2Excel</div>
  <h1>Convert Statements with Precision</h1>
  <p>Clean, reliable extraction for all your financial documents.</p>
</div>
""", unsafe_allow_html=True)

uploaded_files = st.file_uploader(
    "Upload PDF or Photos",
    type=["pdf", "png", "jpg", "jpeg"],
    accept_multiple_files=True,
    label_visibility="collapsed",
)

st.markdown("""
<div class="bc-badgerow"><span class="tag">INPUT</span>
  <span class="bc-badge">PDF</span><span class="bc-badge">Photos</span>
  <span class="bc-badge">JPG</span><span class="bc-badge">PNG</span>
</div>
<div class="bc-badgerow"><span class="tag">OUTPUT</span>
  <span class="bc-badge out">Excel</span><span class="bc-badge out">CSV</span>
</div>
<div class="bc-divider"></div>
<div class="bc-privacy">🔒 We do not store your data. Your privacy is our priority.</div>
""", unsafe_allow_html=True)

if uploaded_files:
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

    st.markdown('<div style="text-align:center; margin: 1rem 0;">', unsafe_allow_html=True)
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

    st.markdown(f"""
    <div class="bc-cards">
      <div class="bc-card">
        <div class="label">Transactions</div>
        <div class="value">{len(df)}</div>
      </div>
      <div class="bc-card bad">
        <div class="label">Total Withdrawn</div>
        <div class="value">₹{total_withdrawal:,.2f}</div>
      </div>
      <div class="bc-card good">
        <div class="label">Total Deposited</div>
        <div class="value">₹{total_deposit:,.2f}</div>
      </div>
      <div class="bc-card {'good' if net >= 0 else 'bad'}">
        <div class="label">Net Change</div>
        <div class="value">₹{net:,.2f}</div>
      </div>
    </div>
    """, unsafe_allow_html=True)

    # --- column selection ---
    st.subheader("Choose columns to include")
    all_cols = list(df.columns)

    def _select_all():
        for c in all_cols:
            st.session_state[f"chk_{c}"] = True

    def _clear_all():
        for c in all_cols:
            st.session_state[f"chk_{c}"] = False

    col1, col2 = st.columns(2)
    with col1:
        st.button("Select all", on_click=_select_all, use_container_width=True)
    with col2:
        st.button("Clear all", on_click=_clear_all, use_container_width=True)

    selected = []
    checkbox_cols = st.columns(3)
    for i, col_name in enumerate(all_cols):
        with checkbox_cols[i % 3]:
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
    st.markdown('<div class="bc-divider"></div>', unsafe_allow_html=True)
    st.subheader("Preview")
    st.dataframe(preview_df.head(20), use_container_width=True, hide_index=True)

    # --- downloads ---
    st.subheader("Download")

    excel_buf = io.BytesIO()
    with pd.ExcelWriter(excel_buf, engine="openpyxl") as writer:
        preview_df.to_excel(writer, index=False, sheet_name="Statement")
    excel_buf.seek(0)

    csv_bytes = preview_df.to_csv(index=False).encode("utf-8")

    dl1, dl2 = st.columns(2)
    with dl1:
        st.download_button(
            "⬇️  Download Excel",
            data=excel_buf,
            file_name="bank_statement.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    with dl2:
        st.download_button(
            "⬇️  Download CSV",
            data=csv_bytes,
            file_name="bank_statement.csv",
            mime="text/csv",
            use_container_width=True,
        )

st.markdown('<div class="bc-footer">Bank2Excel — your files are processed for this session only.</div>', unsafe_allow_html=True)
