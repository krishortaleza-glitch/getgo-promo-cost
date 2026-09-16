from pathlib import Path
import io
import pandas as pd
import streamlit as st

st.set_page_config(page_title='GetGo – Promo Cost', page_icon='📦', layout='wide')
BASE = Path(__file__).resolve().parent
ALIASES_FILE = BASE / 'VendorAliases.xlsx'


def read_file(upload):
    if upload.name.lower().endswith('.csv'):
        return pd.read_csv(upload, dtype='string', keep_default_na=False)
    return pd.read_excel(upload, dtype='string', keep_default_na=False)


def clean(s):
    return s.astype('string').fillna('').str.strip()


def col(df, letter):
    n = 0
    for ch in letter.upper():
        n = n * 26 + ord(ch) - 64
    return clean(df.iloc[:, n - 1])


def key(*parts):
    out = parts[0]
    for part in parts[1:]:
        out = out + part
    return clean(out)


def process(cost, raw, aliases):
    # Vendor Aliases: C + ' ' + D + A
    aliases = aliases.copy()
    aliases['_alias_key'] = key(col(aliases, 'C'), pd.Series(' ', index=aliases.index, dtype='string'), col(aliases, 'D'), col(aliases, 'A'))

    # Raw key 1: D + ' ' + F + ' ' + E + A
    raw = raw.copy()
    raw['_raw_key1'] = key(col(raw, 'D'), pd.Series(' ', index=raw.index, dtype='string'), col(raw, 'F'), pd.Series(' ', index=raw.index, dtype='string'), col(raw, 'E'), col(raw, 'A'))

    # Link raw key 1 to aliases key. Raw vendor zone + vendor group is expected
    # to correspond to the aliases Vendor Zone value.
    alias_map = aliases.drop_duplicates('_alias_key', keep='first').set_index('_alias_key')
    # Map using positional Excel columns, not literal column names like 'E'/'F'.
    alias_values = pd.Series(col(aliases, 'E').to_numpy(dtype=object), index=aliases['_alias_key'])
    alias_cost_zones = pd.Series(col(aliases, 'F').to_numpy(dtype=object), index=aliases['_alias_key'])
    alias_values = alias_values[~alias_values.index.duplicated(keep='first')]
    alias_cost_zones = alias_cost_zones[~alias_cost_zones.index.duplicated(keep='first')]
    raw['_alias'] = raw['_raw_key1'].map(alias_values).fillna('').astype('string')
    raw['_cost_zone'] = raw['_raw_key1'].map(alias_cost_zones).fillna('').astype('string')

    # Raw key 2: alias + raw Column O + cost zone
    raw['_raw_key2'] = key(raw['_alias'], col(raw, 'O'), raw['_cost_zone'])

    # Cost key: B + G + L
    cost = cost.copy()
    cost['_cost_key'] = key(col(cost, 'B'), col(cost, 'G'), col(cost, 'L'))

    # Use a separately named source series to avoid dtype/index assignment issues.
    raw_values = pd.Series(col(raw, 'N').to_numpy(dtype=object), index=raw['_raw_key2'])
    raw_values = raw_values[~raw_values.index.duplicated(keep='first')]
    matched = cost['_cost_key'].map(raw_values)

    destination = 14  # Excel Column O, zero-based
    result = cost.copy()
    result.iloc[:, destination] = pd.Series(matched.fillna('').astype(object).to_numpy(), index=result.index, dtype=object)

    # Remove Excel Column T from the output file, if it exists.
    # Column T is the 20th column, so its zero-based index is 19.
    if result.shape[1] > 19:
        result = result.drop(columns=result.columns[19])

    stats = {
        'cost_rows': len(cost),
        'raw_rows': len(raw),
        'alias_rows': len(aliases),
        'matched': int(matched.notna().sum()),
        'unmatched': int(matched.isna().sum()),
        'alias_unmatched': int((raw['_alias'] == '').sum()),
    }
    diagnostics = raw[['_raw_key1', '_alias', '_cost_zone', '_raw_key2']].copy()
    diagnostics['raw_column_N_source'] = col(raw, 'N').to_numpy(dtype=object)
    diagnostics['raw_column_O_source'] = col(raw, 'O').to_numpy(dtype=object)
    return result, diagnostics, stats


def excel_bytes(df):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='GetGo Promo Cost')
    return buf.getvalue()

st.title('GetGo – Promo Cost')
st.caption('Populate Cost File Column O from Raw Vendor Store Cost Column N (endDate).')

if not ALIASES_FILE.exists():
    st.error(f'Missing static file: {ALIASES_FILE.name}')
else:
    st.success(f'Static aliases file found: {ALIASES_FILE.name}')

raw_upload = st.file_uploader('Raw Vendor Store Cost', type=['csv', 'xlsx', 'xls'])
cost_upload = st.file_uploader('Cost File', type=['csv', 'xlsx', 'xls'])

if st.button('Process GetGo Promo Cost', type='primary', disabled=not (raw_upload and cost_upload and ALIASES_FILE.exists())):
    try:
        raw_df = read_file(raw_upload)
        cost_df = read_file(cost_upload)
        aliases_df = pd.read_excel(ALIASES_FILE, dtype='string', keep_default_na=False)
        result, diagnostics, stats = process(cost_df, raw_df, aliases_df)
        st.session_state.result = result
        st.session_state.diagnostics = diagnostics
        st.session_state.stats = stats
        st.success('Processing completed.')
    except Exception as exc:
        st.exception(exc)

if 'stats' in st.session_state:
    s = st.session_state.stats
    c1, c2, c3, c4 = st.columns(4)
    c1.metric('Cost rows', s['cost_rows'])
    c2.metric('Matched', s['matched'])
    c3.metric('Unmatched', s['unmatched'])
    c4.metric('Alias unmatched', s['alias_unmatched'])
    st.subheader('Lookup diagnostics')
    st.dataframe(st.session_state.diagnostics, use_container_width=True)
    st.subheader('Output preview')
    st.dataframe(st.session_state.result.head(100), use_container_width=True)
    st.download_button('Download GetGo Promo Cost', excel_bytes(st.session_state.result), 'GetGo_Promo_Cost_Output.xlsx', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
