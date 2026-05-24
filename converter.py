import pandas as pd
from io import BytesIO
import re

# ---------------- HELPERS ----------------
def format_size_group(from_size, to_size):
    try:
        return f"{float(from_size):.2f} - {float(to_size):.2f}"
    except:
        return f"{from_size} - {to_size}"

def round2(v):
    try:
        return round(float(v), 2)
    except:
        return v

def safe(df, idx, col, default=None):
    if col and col in df.columns:
        v = df.iloc[idx][col]
        return v if pd.notna(v) else default
    return default

# ---------------- MASTER ----------------
def load_master(file_bytes):
    df = pd.read_excel(BytesIO(file_bytes))
    df.columns = [str(c).strip() for c in df.columns]

    df['Shape'] = df['Shape'].astype(str).str.upper().str.strip()
    df['Color'] = df['Color'].astype(str).str.upper().str.strip()
    df['Clarity'] = df['Clarity'].astype(str).str.upper().str.strip()

    df['From Size'] = pd.to_numeric(df['From Size'], errors='coerce')
    df['To Size'] = pd.to_numeric(df['To Size'], errors='coerce')

    df['Grid'] = pd.to_numeric(df.get('Grid', 0), errors='coerce').fillna(0).astype(int)
    df['Available'] = pd.to_numeric(df.get('Available', 0), errors='coerce').fillna(0).astype(int)

    return df

def master_lookup(master_df, shape, cts, color, clarity):
    try:
        cts = float(cts)
    except:
        return None

    mask = (
        (master_df['Shape'] == shape) &
        (master_df['From Size'] <= cts) &
        (master_df['To Size'] >= cts) &
        (master_df['Color'] == color) &
        (master_df['Clarity'] == clarity)
    )

    rows = master_df[mask]
    return rows.iloc[0] if len(rows) > 0 else None

# ---------------- PARTY ----------------
def detect_header_row(file_bytes):
    df_raw = pd.read_excel(BytesIO(file_bytes), header=None, nrows=6)
    for i in range(6):
        row = df_raw.iloc[i].astype(str).str.lower().tolist()
        s = ' '.join(row)
        if any(x in s for x in ['stock', 'packet', 'stone', 'status']):
            return i
    return 0

def load_party(file_bytes):
    h = detect_header_row(file_bytes)
    df = pd.read_excel(BytesIO(file_bytes), header=h)
    df.columns = [str(c).strip() for c in df.columns]
    return df

def find_col(df, candidates):
    cols_lower = {c.lower(): c for c in df.columns}
    for cand in candidates:
        if cand.lower() in cols_lower:
            return cols_lower[cand.lower()]
    for cand in candidates:
        for cl, c in cols_lower.items():
            if cand.lower() in cl:
                return c
    return None

def map_columns(df):
    cols = list(df.columns)

    has_vdb = any('vdb' in c.lower() for c in cols)
    has_status = any('status' in c.lower() for c in cols)

    if has_vdb and has_status:
        def gc(i): return cols[i] if i < len(cols) else None
        return {
            'STONE_ID': gc(1),
            'SHAPE': gc(2),
            'CTS': gc(3),
            'SIZE_GROUP': gc(4),
            'COLOR': gc(5),
            'CLARITY': gc(6),
            'PER_CARAT': gc(7),
            'VDB_IND': gc(9),
            'VDB_USA': gc(11),
            'WITH_TARIFF': gc(13),
            'IS_QC': True
        }
    else:
        return {
            'STONE_ID': find_col(df, ['Stock', 'Packet', 'Stone']),
            'SHAPE': find_col(df, ['Shape']),
            'CTS': find_col(df, ['Weight', 'CTS']),
            'COLOR': find_col(df, ['Color']),
            'CLARITY': find_col(df, ['Clarity']),
            'PER_CARAT': find_col(df, ['VDB', 'Price', 'Carat']),
            'VDB_IND': find_col(df, ['VDB']),
            'VDB_USA': find_col(df, ['USA']),
            'WITH_TARIFF': find_col(df, ['Tariff']),
            'IS_QC': False
        }

# ---------------- CONVERT ----------------
def convert(master_df, party_df, col_map, user_inputs):
    rows = []

    for i in range(len(party_df)):
        stone = safe(party_df, i, col_map['STONE_ID'], '')
        shape = str(safe(party_df, i, col_map['SHAPE'], '')).upper()
        cts = safe(party_df, i, col_map['CTS'], 0)
        color = str(safe(party_df, i, col_map['COLOR'], '')).upper()
        clarity = str(safe(party_df, i, col_map['CLARITY'], '')).upper()
        per_carat = safe(party_df, i, col_map['PER_CARAT'], 0)

        try:
            cts = float(cts)
        except:
            cts = 0

        m = master_lookup(master_df, shape, cts, color, clarity)

        if m is not None:
            size_group = format_size_group(m['From Size'], m['To Size'])
            grid = m['Grid']
        else:
            size_group = ''
            grid = 0

        total = round2(float(per_carat) * cts)

        rows.append({
            "STONE ID": stone,
            "SHAPE": shape,
            "CTS": cts,
            "COLOR": color,
            "CLARITY": clarity,
            "SIZE GROUP": size_group,
            "GRID": grid,
            "PER CARAT": per_carat,
            "TOTAL": total
        })

    return pd.DataFrame(rows)

# ---------------- EXCEL ----------------
def to_excel(df):
    buf = BytesIO()
    df.to_excel(buf, index=False)
    buf.seek(0)
    return buf.getvalue()