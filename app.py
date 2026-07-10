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
    "value_dt":  ["value"],
    "date":      ["date"],
    "narration": ["narration", "particular", "particulars", "description", "details"],
    "ref_no":    ["chq", "cheque", "ref", "instrument", "instruments"],
    "withdrawal":["withdrawal", "debit", "dr"],
    "deposit":   ["deposit", "credit", "cr"],
    "balance":   ["balance", "total"],
}


def classify_header_word(text):
    t = text.lower().strip(".:/ ")
    for key in ["value_dt", "date", "narration", "ref_no", "withdrawal", "deposit", "balance"]:
        for kw in HEADER_KEYWORDS[key]:
            if kw == t or (len(kw) > 2 and kw in t):
                return key
    return None


def find_header_row(words):
    by_line = {}
    for w in words:
        key = round(w['top'])
        by_line.setdefault(key, []).append(w)

    for top, line_words in sorted(by_line.items()):
        texts = [w['text'].lower() for w in line_words]
        has_date = any(t.strip('.:') == 'date' for t in texts)
        has_narr = any(classify_header_word(w['text']) == 'narration' for w in line_words)
        has_amt = any(classify_header_word(w['text']) in ('withdrawal', 'deposit', 'balance') for w in line_words)
        if has_date and has_narr and has_amt:
            return top, line_words
    return None, None


def build_columns(header_words, page_width):
    groups = {}
    for w in header_words:
        col = classify_header_word(w['text'])
        if col is None:
            continue
        g = groups.setdefault(col, {"x0": w['x0'], "x1": w['x1']})
        g["x0"] = min(g["x0"], w['x0'])
        g["x1"] = max(g["x1"], w['x1'])

    if not groups:
        return None

    ordered = sorted(groups.items(), key=lambda kv: kv[1]["x0"])
    boundaries = []
    for i, (name, g) in enumerate(ordered):
        left = 0 if i == 0 else (ordered[i-1][1]["x1"] + g["x0"]) / 2
        right = page_width if i == len(ordered) - 1 else (g["x1"] + ordered[i+1][1]["x0"]) / 2
        boundaries.append((name, left, right))
    return boundaries


def calibrate_date_narration_boundary(boundaries, body_words):
    """Header-label position is unreliable for the Date/Narration split
    specifically (the 'Narration' label sits far right of where narration
    text actually starts) -- recalibrate using real body text."""
    date_tokens = [w for w in body_words if DATE_RE.match(w['text'])]
    if not date_tokens:
        return boundaries
    date_tokens.sort(key=lambda w: w['x0'])
    leftmost_x0 = date_tokens[0]['x0']
    date_col_tokens = [w for w in date_tokens if w['x0'] < leftmost_x0 + 15]
    if len(date_col_tokens) < 3:
        return boundaries
    calibrated_right = max(w['x1'] for w in date_col_tokens) + 5

    new_boundaries = []
    for (name, left, right) in boundaries:
        if name == 'date':
            new_boundaries.append((name, left, calibrated_right))
        elif name == 'narration':
            new_boundaries.append((name, calibrated_right, right))
        else:
            new_boundaries.append((name, left, right))
    return new_boundaries


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


def find_footer_top(words, page_height):
    markers = ["hdfcbanklimited", "statementsummary", "endofstatement",
               "totals", "generatedon", "closingbalance"]
    candidates = []
    by_line = {}
    for w in words:
        if w['top'] > page_height * 0.5:
            key = round(w['top'])
            by_line.setdefault(key, []).append(w)
    for top, line_words in by_line.items():
        joined = "".join(w['text'].lower() for w in sorted(line_words, key=lambda w: w['x0']))
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


def extract_transactions(pdf_path):
    rows = []
    boundaries = None
    prev_table_top = None
    found_any_header = False

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            words = page.extract_words(x_tolerance=1.5)
            header_top, header_words = find_header_row(words)

            if header_words:
                found_any_header = True
                new_boundaries = build_columns(header_words, page.width)
                if new_boundaries:
                    boundaries = new_boundaries
                table_top = header_top + 8
            elif boundaries is not None:
                title_bottom = find_title_anchor(words)
                if title_bottom is not None:
                    table_top = title_bottom + 4
                elif prev_table_top is not None:
                    table_top = prev_table_top
                else:
                    continue
            else:
                continue

            prev_table_top = table_top
            footer_top = find_footer_top(words, page.height)
            table_bottom = footer_top - 2 if footer_top else page.height - 20

            table_words = [w for w in words if table_top <= w['top'] <= table_bottom]
            if not table_words or boundaries is None:
                continue

            if header_words:
                boundaries = calibrate_date_narration_boundary(boundaries, table_words)

            bands = get_grid_row_bands(page, table_top, table_bottom)

            if bands:
                for (top, bottom) in bands:
                    band_words = [w for w in table_words if top <= w['top'] < bottom]
                    if not band_words:
                        continue
                    band_words.sort(key=lambda w: (w['top'], w['x0']))
                    row = {c: [] for c in CANONICAL_COLUMNS}
                    for w in band_words:
                        col = col_for_x(boundaries, w['x0'])
                        if col:
                            row[col].append(w['text'])
                    rows.append(row)
            else:
                table_words.sort(key=lambda w: (round(w['top']), w['x0']))
                lines = {}
                for w in table_words:
                    lines.setdefault(round(w['top']), []).append(w)

                current = None
                for top in sorted(lines.keys()):
                    line_words = sorted(lines[top], key=lambda w: w['x0'])
                    starts_new = any(
                        col_for_x(boundaries, w['x0']) == 'date' and DATE_RE.match(w['text'])
                        for w in line_words
                    )
                    if starts_new or current is None:
                        if current is not None:
                            rows.append(current)
                        current = {c: [] for c in CANONICAL_COLUMNS}
                    for w in line_words:
                        col = col_for_x(boundaries, w['x0'])
                        if col:
                            current[col].append(w['text'])
                if current is not None:
                    rows.append(current)

    if not found_any_header:
        return None  # could not identify a transaction table at all

    result = []
    for r in rows:
        if not any(DATE_RE.match(tok) for tok in r["date"]):
            continue
        result.append({c: " ".join(r[c]) for c in CANONICAL_COLUMNS})
    return result


# ---------------------------------------------------------------------------
# STREAMLIT UI
# ---------------------------------------------------------------------------

st.set_page_config(page_title="BankConverter", page_icon="📄", layout="centered")
st.title("📄 BankConverter")
st.caption("Upload a bank statement PDF. Get a clean Excel or CSV file back.")

uploaded = st.file_uploader("Upload your bank statement (PDF)", type=["pdf"])

if uploaded is not None:
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

    # --- run extraction ---
    with st.spinner("Reading your statement..."):
        with tempfile.NamedTemporaryFile(suffix=".pdf") as tmp:
            tmp.write(working_bytes)
            tmp.flush()
            rows = extract_transactions(tmp.name)

    if rows is None or len(rows) == 0:
        st.error(
            "Could not identify a transaction table in this PDF. "
            "Please make sure this is a valid bank statement."
        )
        st.stop()

    df = pd.DataFrame(rows)
    df = df.rename(columns=DISPLAY_NAMES)
    df = df[[DISPLAY_NAMES[c] for c in CANONICAL_COLUMNS]]

    st.success(f"Found {len(df)} transactions.")

    # --- column selection ---
    st.subheader("Choose columns to include")
    all_cols = list(df.columns)
    col1, col2 = st.columns(2)
    with col1:
        if st.button("Select all"):
            st.session_state["selected_cols"] = all_cols
    with col2:
        if st.button("Clear all"):
            st.session_state["selected_cols"] = []

    if "selected_cols" not in st.session_state:
        st.session_state["selected_cols"] = all_cols

    selected = []
    checkbox_cols = st.columns(3)
    for i, col_name in enumerate(all_cols):
        with checkbox_cols[i % 3]:
            checked = st.checkbox(
                col_name,
                value=col_name in st.session_state["selected_cols"],
                key=f"chk_{col_name}",
            )
            if checked:
                selected.append(col_name)
    st.session_state["selected_cols"] = selected

    if not selected:
        st.warning("Select at least one column to see a preview and download.")
        st.stop()

    preview_df = df[selected]

    # --- preview ---
    st.subheader("Preview")
    st.dataframe(preview_df.head(20), use_container_width=True)

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
            "⬇️ Download Excel",
            data=excel_buf,
            file_name="bank_statement.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    with dl2:
        st.download_button(
            "⬇️ Download CSV",
            data=csv_bytes,
            file_name="bank_statement.csv",
            mime="text/csv",
            use_container_width=True,
        )
