from pathlib import Path
import io
import pandas as pd
import streamlit as st

st.set_page_config(page_title='GetGo - Promo Cost', page_icon='📦', layout='wide')
BASE_DIR = Path(__file__).resolve().parent
ALIASES_CANDIDATES = [BASE_DIR / 'VendorAliases.xlsx', BASE_DIR / 'vendor_aliases.xlsx', BASE_DIR / 'VendorAliases.csv']


def read_file(uploaded):
    if uploaded.name.lower().endswith('.csv'):
        return pd.read_csv(uploaded, dtype=str, keep_default_na=False).fillna('')
    return pd.read_excel(uploaded, dtype=str, keep_default_na=False).fillna('')


def read_alias_file():
    path = next((p for p in ALIASES_CANDIDATES if p.exists()), None)
    if path is None:
        raise FileNotFoundError('Place VendorAliases.xlsx, vendor_aliases.xlsx, or VendorAliases.csv in the repository root.')
    if path.suffix.lower() == '.csv':
        df = pd.read_csv(path, dtype=str, keep_default_na=False)
    else:
        df = pd.read_excel(path, dtype=str, keep_default_na=False)
    return df.fillna(''), path.name


def norm(value):
    """Normalize either a pandas Series or a single text value."""
    if isinstance(value, pd.Series):
        return (
            value.fillna('')
            .astype(str)
            .str.strip()
            .str.replace(r'\s+', ' ', regex=True)
            .str.upper()
        )

    if value is None:
        return ''

    return ' '.join(str(value).strip().upper().split())


def col(df, letter):
    n = 0
    for ch in letter.upper():
        n = n * 26 + ord(ch) - 64
    idx = n - 1
    if idx >= len(df.columns):
        raise ValueError(f'Column {letter} is not available. File has {len(df.columns)} columns.')
    return df.iloc[:, idx].fillna('').astype(str)


def add_key(*parts):
    """Concatenate aligned Series values while preserving literal separators."""
    series_parts = [part for part in parts if isinstance(part, pd.Series)]

    if not series_parts:
        return ''.join(str(part) if part is not None else '' for part in parts)

    result = pd.Series('', index=series_parts[0].index, dtype='object')

    for part in parts:
        if isinstance(part, pd.Series):
            result = result + norm(part).reindex(result.index, fill_value='')
        else:
            # Preserve separators such as " "; do not strip them.
            result = result + (str(part) if part is not None else '')

    return result


def process(cost, raw, aliases):
    # Vendor Aliases: lookupkey = C + ' ' + D + A
    aliases = aliases.copy()
    aliases['_alias_lookupkey'] = add_key(col(aliases, 'C'), ' ', col(aliases, 'D'), col(aliases, 'A'))
    aliases['_alias_value'] = col(aliases, 'E').str.strip()
    aliases['_cost_zone_value'] = col(aliases, 'F').str.strip()

    # Raw Vendor Store Cost: X = D + ' ' + F + ' ' + C; raw key = X + A
    raw = raw.copy()
    raw['_raw_vendor_lookupkey'] = add_key(col(raw, 'D'), ' ', col(raw, 'F'), ' ', col(raw, 'C'), col(raw, 'A'))

    # Separate-field link: vendor name, vendor zone, and store are matched independently.
    # This is equivalent to linking the alias lookup key to the raw vendor lookup key,
    # without concatenating the two keys into one lookup field.
    alias_index = aliases.drop_duplicates('_alias_lookupkey', keep='first').set_index('_alias_lookupkey')
    raw['_alias_lookupkey'] = add_key(col(raw, 'D'), ' ', col(raw, 'F'), col(raw, 'A'))
    raw['_alias'] = raw['_alias_lookupkey'].map(alias_index['_alias_value']).fillna('')
    raw['_cost_zone'] = raw['_alias_lookupkey'].map(alias_index['_cost_zone_value']).fillna('')

    # raw.vendor.lookupkey2 = alias + O + cost zone
    raw['_raw_vendor_lookupkey2'] = add_key(raw['_alias'], col(raw, 'O'), raw['_cost_zone'])

    # GetGo Promo Cost lookupkey = B + G + L
    cost = cost.copy()
    cost['_getgo_lookupkey'] = add_key(col(cost, 'B'), col(cost, 'G'), col(cost, 'L'))

    # Final lookup: return Raw Column N (endDate)
    raw_unique = raw.drop_duplicates('_raw_vendor_lookupkey2', keep='first').copy()
    raw_unique['_raw_end_date'] = col(raw, 'N').reindex(raw_unique.index).fillna('')
    raw_values = raw_unique.set_index('_raw_vendor_lookupkey2')['_raw_end_date']
    matched = cost['_getgo_lookupkey'].map(raw_values)

    destination_idx = 14  # Excel Column O
    if destination_idx >= len(cost.columns):
        raise ValueError('Cost File does not contain Excel Column O.')
    destination_name = cost.columns[destination_idx]
    original = cost.iloc[:, destination_idx].fillna('').astype(str)

    # Pandas 3 / newer Streamlit environments may read CSV string columns
    # using an Arrow-backed string dtype. Assigning a NumPy array directly
    # into that column can raise: "Invalid value '['']' for dtype 'str'".
    # Build a normal object-dtype Series and assign it by column name.
    replacement = matched.astype(object).where(
        matched.notna(),
        original.astype(object),
    )
    cost[destination_name] = pd.Series(
        replacement.to_numpy(dtype=object),
        index=cost.index,
        dtype=object,
    )

    diagnostics = {
        'cost_rows': len(cost),
        'raw_rows': len(raw),
        'alias_rows': len(aliases),
        'matched_rows': int(matched.notna().sum()),
        'unmatched_rows': int(matched.isna().sum()),
        'raw_alias_unmatched': int((raw['_alias'] == '').sum()),
        'duplicate_alias_keys': int(aliases['_alias_lookupkey'].duplicated(keep=False).sum()),
        'duplicate_raw_final_keys': int(raw['_raw_vendor_lookupkey2'].duplicated(keep=False).sum()),
    }
    return cost, raw, diagnostics


def make_output(df):
    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='GetGo Promo Cost')
    bio.seek(0)
    return bio.getvalue()


st.title('GetGo – Promo Cost')
st.caption('Populate Cost File Column O from Raw Vendor Store Cost Column N (endDate).')

raw_upload = st.file_uploader('Raw Vendor Store Cost', type=['csv', 'xlsx', 'xls'])
cost_upload = st.file_uploader('Cost File', type=['csv', 'xlsx', 'xls'])

alias_path = next((p for p in ALIASES_CANDIDATES if p.exists()), None)
if alias_path:
    st.success(f'Static Vendor Aliases file found: {alias_path.name}')
else:
    st.error('Static Vendor Aliases file not found in the repository root.')

if st.button('Process GetGo Promo Cost', type='primary', use_container_width=True):
    if not raw_upload or not cost_upload:
        st.error('Upload both the Raw Vendor Store Cost and Cost File.')
        st.stop()
    try:
        raw_df = read_file(raw_upload)
        cost_df = read_file(cost_upload)
        aliases_df, alias_name = read_alias_file()
        result, processed_raw, stats = process(cost_df, raw_df, aliases_df)
        st.session_state['result'] = result
        st.session_state['bytes'] = make_output(result)
        st.session_state['stats'] = stats
        st.success('Processing completed.')
    except Exception as exc:
        st.error(str(exc))
        st.exception(exc)

if 'stats' in st.session_state:
    st.subheader('Processing summary')
    s = st.session_state['stats']
    a, b, c, d = st.columns(4)
    a.metric('Cost rows', s['cost_rows'])
    b.metric('Matched rows', s['matched_rows'])
    c.metric('Unmatched rows', s['unmatched_rows'])
    d.metric('Raw rows', s['raw_rows'])
    st.write(f"Unmatched raw alias links: **{s['raw_alias_unmatched']}**")
    st.write(f"Duplicate alias keys: **{s['duplicate_alias_keys']}**")
    st.write(f"Duplicate raw final keys: **{s['duplicate_raw_final_keys']}**")
    st.subheader('Output preview')
    st.dataframe(st.session_state['result'].head(100), use_container_width=True)
    st.download_button('Download GetGo Promo Cost Output', st.session_state['bytes'], 'GetGo_Promo_Cost_Output.xlsx', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', use_container_width=True)
