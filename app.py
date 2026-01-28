import io
import re
import pandas as pd
import streamlit as st

st.set_page_config(page_title='TSC 计算模型', layout='wide')

st.title('计算模型')

HIDE_COLS = ['半成品入库量', '调整后实际量', '调整后实际额', '碎肉量', '辅助']
FMT_DISPLAY = {
    '修形前原料综合耗用单价': '{:.2f}',
    '半成品原料成本': '{:.2f}',
    '半成品修形人工成本': '{:.2f}',
    '半成品总成本': '{:.2f}',
    '修形利用率': '{:.0%}',
    '损耗率': '{:.0%}',
}

with st.sidebar:
    st.header('输入文件')
    file_compare = st.file_uploader('基础表', type=['xlsx'])
    file_rawlist = st.file_uploader('原料清单.xlsx', type=['xlsx'])
    file_q3 = st.file_uploader('系统成本', type=['xlsx'])

def _find_header_row(df, keyword):
    for i in range(min(len(df), 50)):
        row = df.iloc[i].astype(str).apply(_clean_colname)
        if row.str.contains(_clean_colname(keyword)).any():
            return i
    return None


def _normalize_mat(val):
    if pd.isna(val):
        return ''
    if isinstance(val, (int,)):
        return str(val)
    if isinstance(val, float):
        if val.is_integer():
            return str(int(val))
        return str(val).rstrip('0').rstrip('.')
    s = str(val).strip()
    if s.endswith('.0'):
        s = s[:-2]
    return s


def _clean_colname(val):
    s = str(val).strip()
    s = re.sub(r'\s+', '', s)
    return s


def _load_rawlist(rawlist_file):
    sheets = pd.read_excel(rawlist_file, sheet_name=None)
    frames = []
    for df in sheets.values():
        df = df.copy()
        df.columns = [_clean_colname(c) for c in df.columns]
        if '原料号' in df.columns and '部位' in df.columns:
            frames.append(df[['原料号', '部位']])
    if not frames:
        return pd.DataFrame(columns=['原料号', '部位'])
    raw = pd.concat(frames, ignore_index=True)
    raw['原料号'] = raw['原料号'].apply(_normalize_mat)
    raw['部位'] = raw['部位'].fillna('').astype(str).str.strip()
    return raw


def _category_from_part(part):
    if part == '腿类':
        return '腿肉'
    if part == '胸类':
        return '胸肉'
    return '其他'


def _ensure_columns(df, cols):
    cols_set = set(df.columns)
    missing = [c for c in cols if c not in cols_set]
    if missing:
        raise ValueError(f'缺少列: {missing}')


def _normalize_columns(df):
    col_map = {}
    for col in df.columns:
        key = _clean_colname(col)
        if key not in col_map:
            col_map[col] = key
    df = df.rename(columns=col_map)
    # handle duplicate names by keeping the first occurrence
    df = df.loc[:, ~df.columns.duplicated()]
    return df


def _ensure_and_map_columns(df, required):
    # map cleaned names to existing columns
    clean_map = {}
    for col in df.columns:
        clean = _clean_colname(col)
        if clean not in clean_map:
            clean_map[clean] = col

    def _find_col_by_keywords(keywords):
        for col in df.columns:
            clean = _clean_colname(col)
            for kw in keywords:
                if kw in clean:
                    return col
        return None

    alias_map = {
        '物料号': ['物料编号', '物料编码', '物料代码', '物料'],
        '物料描述': ['物料名称', '品名', '物料规格'],
        '原料号': ['原料编号', '原料编码', '原料代码', '原料'],
        '原料描述': ['原料名称', '原料规格', '原料品名'],
        '入库数量': ['入库量', '入库数量(kg)', '入库数量（kg）'],
        '入库金额': ['入库金额(元)', '入库金额（元）', '入库金额含税'],
        '实际数量': ['实际量', '实际数量(kg)', '实际数量（kg）'],
        '实际金额': ['实际金额(元)', '实际金额（元）', '实际金额含税'],
        '配方数量': ['配方量', '配方用量', '配方数量(kg)', '配方数量（kg）'],
    }

    # rename columns to required names where possible
    rename_map = {}
    for req in required:
        if req in df.columns:
            continue
        clean_req = _clean_colname(req)
        if clean_req in clean_map:
            rename_map[clean_map[clean_req]] = req
            continue
        aliases = alias_map.get(req, [])
        alias_clean = [_clean_colname(a) for a in aliases]
        col = _find_col_by_keywords(alias_clean + [clean_req])
        if col:
            rename_map[col] = req
    if rename_map:
        df = df.rename(columns=rename_map)

    # fallback: any column containing "物料号"
    if '物料号' not in df.columns:
        col = _find_col_by_keywords(['物料号', '物料'])
        if col:
            df = df.rename(columns={col: '物料号'})

    _ensure_columns(df, required)
    return df


def _to_num(s):
    return pd.to_numeric(s, errors='coerce').fillna(0.0)


def _find_tsc_value(tsc_file, sheet_name, material_no, row_label):
    df = pd.read_excel(tsc_file, sheet_name=sheet_name, header=None)
    # Find column index for 综合单价
    header_row = _find_header_row(df, '综合单价')
    if header_row is None:
        return None
    header = [_clean_colname(c) for c in df.iloc[header_row].astype(str).tolist()]
    try:
        price_col = header.index('综合单价')
    except ValueError:
        return None
    # Locate row by material no (col 3) and label (col 5)
    mat_col = 2
    label_col = 4
    for i in range(header_row + 1, len(df)):
        mat = _normalize_mat(df.iat[i, mat_col])
        label = str(df.iat[i, label_col]).strip()
        if mat == material_no and label == row_label:
            val = df.iat[i, price_col]
            try:
                return float(val)
            except Exception:
                return None
    return None


def _find_tsc_metrics(tsc_file, sheet_name, material_no, row_label):
    df = pd.read_excel(tsc_file, sheet_name=sheet_name, header=None)
    header_row = _find_header_row(df, '修形前原料综合耗用单价')
    if header_row is None:
        return None
    header = [_clean_colname(c) for c in df.iloc[header_row].astype(str).tolist()]
    names = [
        '修形前原料综合耗用单价',
        '修形利用率',
        '损耗率',
        '半成品原料成本',
        '半成品修形人工成本',
        '半成品总成本',
    ]
    cols = {}
    for name in names:
        if name not in header:
            return None
        cols[name] = header.index(name)
    mat_col = 2
    label_col = 4
    for i in range(header_row + 1, len(df)):
        mat = _normalize_mat(df.iat[i, mat_col])
        label = str(df.iat[i, label_col]).strip()
        if mat == material_no and label == row_label:
            out = {}
            for k, c in cols.items():
                val = df.iat[i, c]
                try:
                    out[k] = float(val)
                except Exception:
                    out[k] = None
            return out
    return None


def _get_tsc_raw_columns(tsc_file, sheet_name):
    df = pd.read_excel(tsc_file, sheet_name=sheet_name, header=None)
    header_row = _find_header_row(df, '综合单价')
    if header_row is None or header_row == 0:
        return [], None, None, {}
    raw_row = header_row - 1
    raw_cols = []
    spec_map = {}
    for c in range(df.shape[1]):
        val = df.iat[raw_row, c]
        if pd.isna(val):
            continue
        s = str(val).strip()
        if s.isdigit():
            raw_cols.append((c, s))
            spec_val = df.iat[header_row, c]
            spec_map[s] = str(spec_val).strip() if not pd.isna(spec_val) else ''
    comp_col = None
    for c in range(df.shape[1]):
        v = df.iat[header_row, c]
        if str(v).strip() == '综合单价':
            comp_col = c
            break
    return raw_cols, comp_col, df, spec_map


def _find_tsc_row_values(df, material_no, row_label, raw_cols, comp_col):
    mat_col = 2
    label_col = 4
    for i in range(len(df)):
        mat = _normalize_mat(df.iat[i, mat_col])
        label = str(df.iat[i, label_col]).strip()
        if mat == material_no and label == row_label:
            out = {}
            for c, code in raw_cols:
                try:
                    out[code] = float(df.iat[i, c])
                except Exception:
                    out[code] = None
            if comp_col is not None:
                try:
                    out['综合单价'] = float(df.iat[i, comp_col])
                except Exception:
                    out['综合单价'] = None
            return out
    return {}


def _load_compare_df(compare_file, sheet_name=None):
    df = pd.read_excel(compare_file, sheet_name=sheet_name or 0)
    df = _normalize_columns(df)
    if '物料号' in df.columns:
        return df
    raw = pd.read_excel(compare_file, sheet_name=sheet_name or 0, header=None)
    header_row = _find_header_row(raw, '物料号')
    if header_row is None:
        header_row = 0
    header = [_clean_colname(c) for c in raw.iloc[header_row].astype(str).tolist()]
    data = raw.iloc[header_row + 1 :].copy()
    data.columns = header
    data = _normalize_columns(data)
    if '物料号' not in data.columns and data.shape[1] > 0:
        # last resort: treat first column as 物料号
        data = data.rename(columns={data.columns[0]: '物料号'})
    return data


def _drop_hidden_cols(df):
    df = df.drop(columns=[c for c in HIDE_COLS if c in df.columns])
    df = df.drop(columns=['行序'], errors='ignore')
    return df


def compute(compare_file, rawlist_file, q3_file, sheet_name=None):
    df = _load_compare_df(compare_file, sheet_name=sheet_name)
    if '物料号' not in df.columns and df.shape[1] > 0:
        candidates = [c for c in df.columns if '物料' in str(c)]
        if candidates:
            df = df.rename(columns={candidates[0]: '物料号'})
        else:
            df = df.rename(columns={df.columns[0]: '物料号'})

    required = ['物料号', '物料描述', '原料号', '原料描述', '入库数量', '入库金额', '实际数量', '实际金额', '配方数量']
    df = _ensure_and_map_columns(df, required)
    if '物料号' not in df.columns and df.shape[1] > 0:
        df['物料号'] = df.iloc[:, 0]

    rawlist = _load_rawlist(rawlist_file)
    raw_map = dict(zip(rawlist['原料号'], rawlist['部位']))
    raw_set = set(raw_map.keys())

    df['物料号'] = df['物料号'].apply(_normalize_mat)

    df['原料号'] = df['原料号'].apply(_normalize_mat)
    df['原料描述'] = df['原料描述'].fillna('').astype(str).str.strip()
    df['物料描述'] = df['物料描述'].fillna('').astype(str).str.strip()

    df['入库数量'] = _to_num(df['入库数量'])
    df['入库金额'] = _to_num(df['入库金额'])
    df['实际数量'] = _to_num(df['实际数量'])
    df['实际金额'] = _to_num(df['实际金额'])
    df['配方数量'] = _to_num(df['配方数量'])

    # Keep all materials/rows; do not filter by raw list
    df = df.copy()

    # Build material -> category mapping from raw list (by part)
    mapping = {}
    for mat in df['物料号'].unique():
        part = raw_map.get(mat, '')
        mapping[mat] = _category_from_part(part)

    df = df[df['物料号'].isin(mapping.keys())].copy()

    agg = {}
    for _, row in df.iterrows():
        mat = row['物料号']
        if mat not in agg:
            agg[mat] = {
                '物料号': mat,
                '分类': mapping[mat],
                '物料描述': row['物料描述'],
                '入库数量': 0.0,
                '入库金额': 0.0,
                '调整后实际量': 0.0,
                '调整后实际额': 0.0,
                '碎肉量': 0.0,
                '人工费用实际额': 0.0,
            }
        if row['物料描述']:
            agg[mat]['物料描述'] = row['物料描述']

        raw = row['原料号']
        raw_desc = row['原料描述']

        if raw == '' or pd.isna(raw):
            agg[mat]['入库数量'] += row['入库数量']
            agg[mat]['入库金额'] += row['入库金额']
            continue

        if raw == '人工费用' or ('人工' in raw_desc):
            agg[mat]['人工费用实际额'] += row['实际金额']
            continue

        if row['实际数量'] < 0:
            agg[mat]['碎肉量'] += row['实际数量']
            continue

        agg[mat]['调整后实际量'] += row['实际数量']
        agg[mat]['调整后实际额'] += row['实际金额']

    records = []
    for mat, v in agg.items():
        adj_qty = v['调整后实际量']
        adj_amt = v['调整后实际额']
        in_qty = v['入库数量']
        scrap_qty = v['碎肉量']

        auxiliary = adj_qty
        unit = adj_amt / adj_qty if adj_qty != 0 else None
        scrap_ratio = (abs(scrap_qty) / auxiliary) if auxiliary != 0 else None
        util = (in_qty / auxiliary) if auxiliary != 0 else None
        loss = (1 - util - scrap_ratio) if auxiliary != 0 else None
        factor = 0.7 if v['分类'] == '胸肉' else 0.95
        raw_cost = None
        if v['分类'] == '其他':
            raw_cost = unit
        elif unit is not None and util not in (None, 0):
            raw_cost = (unit - (1 - util - loss) * unit * factor) / util
        labor = (v['人工费用实际额'] / in_qty) if in_qty != 0 else None
        total = (raw_cost + labor) if (raw_cost is not None and labor is not None) else None

        records.append({
            '物料号': v['物料号'],
            '分类': v['分类'],
            '物料描述': v['物料描述'],
            '行类型': '11月实际单价',
            '行序': 0,
            '半成品入库量': in_qty,
            '调整后实际量': adj_qty,
            '调整后实际额': adj_amt,
            '碎肉量': scrap_qty,
            '修形前原料综合耗用单价': unit,
            '修形利用率': util,
            '损耗率': loss,
            '半成品原料成本': raw_cost,
            '半成品修形人工成本': labor,
            '半成品总成本': total,
        })

    result = pd.DataFrame(records)
    # Add Q3 actual price / 11月规格占比 / Q3规格占比 rows
    extra_rows = []
    row_order = {
        '11月实际单价': 0,
        'Q3实际单价': 1,
        '11月规格占比': 2,
        'Q3规格占比': 3,
    }
    for mat, v in agg.items():
        tsc_sheet = '腿肉TSC' if v['分类'] == '腿肉' else '胸肉TSC'
        q3_ratio = _find_tsc_value(q3_file, tsc_sheet, mat, 'Q3规格占比')
        nov_ratio = None
        q3_metrics = _find_tsc_metrics(q3_file, tsc_sheet, mat, 'Q3实际单价')

        for label, val in [
            ('11月规格占比', nov_ratio),
            ('Q3规格占比', q3_ratio),
        ]:
            extra_rows.append({
                '物料号': mat,
                '分类': v['分类'],
                '物料描述': v['物料描述'],
                '行类型': label,
                '行序': row_order[label],
                '半成品入库量': None,
                '调整后实际量': None,
                '调整后实际额': None,
                '碎肉量': None,
                '修形前原料综合耗用单价': val,
                '修形利用率': None,
                '损耗率': None,
                '半成品原料成本': None,
                '半成品修形人工成本': None,
                '半成品总成本': None,
            })

        if q3_metrics:
            extra_rows.append({
                '物料号': mat,
                '分类': v['分类'],
                '物料描述': v['物料描述'],
                '行类型': 'Q3实际单价',
                '行序': row_order['Q3实际单价'],
                '半成品入库量': None,
                '调整后实际量': None,
                '调整后实际额': None,
                '碎肉量': None,
                '修形前原料综合耗用单价': q3_metrics.get('修形前原料综合耗用单价'),
                '修形利用率': q3_metrics.get('修形利用率'),
                '损耗率': q3_metrics.get('损耗率'),
                '半成品原料成本': q3_metrics.get('半成品原料成本'),
                '半成品修形人工成本': q3_metrics.get('半成品修形人工成本'),
                '半成品总成本': q3_metrics.get('半成品总成本'),
            })

    if '行序' not in result.columns:
        result['行序'] = row_order['11月实际单价']
    if extra_rows:
        extra_df = pd.DataFrame(extra_rows)
        if not extra_df.dropna(how='all').empty:
            result = pd.concat([result, extra_df], ignore_index=True)
    if '物料号' not in result.columns:
        result['物料号'] = ''
    if '分类' not in result.columns:
        result['分类'] = ''
    if '行序' not in result.columns:
        result['行序'] = 0
    legs = result[result['分类'] == '腿肉'].sort_values(['物料号', '行序'])
    breast = result[result['分类'] == '胸肉'].sort_values(['物料号', '行序'])
    other = result[
        (result['分类'] == '其他') & (result['物料号'].astype(str).str.startswith('390'))
    ].sort_values(['物料号', '行序'])

    # Build raw usage matrix (like TSC F:O)
    raw_cols_legs, comp_col_legs, df_q3_legs, spec_map_legs = _get_tsc_raw_columns(q3_file, '腿肉TSC')
    raw_cols_breast, comp_col_breast, df_q3_breast, spec_map_breast = _get_tsc_raw_columns(q3_file, '胸肉TSC')

    df_calc = df.copy()
    df_calc = df_calc[df_calc['原料号'] != '']
    df_calc['正向数量'] = df_calc['实际数量'].where(df_calc['实际数量'] > 0, 0)
    df_calc['正向金额'] = df_calc['实际金额'].where(df_calc['实际数量'] > 0, 0)

    grp = (
        df_calc.groupby(['物料号', '原料号'], as_index=False)[['正向数量', '正向金额']]
        .sum()
        .rename(columns={'原料号': '原料号_raw'})
    )

    aux_map = {k: v['调整后实际量'] for k, v in agg.items()}
    desc_map = {k: v['物料描述'] for k, v in agg.items()}

    def build_matrix_rows(materials, raw_cols, comp_col, df_q3, category_label):
        rows = []
        base_cols = ['物料号', '物料描述', '行类型', '分类']
        raw_codes = [code for _, code in raw_cols]
        for mat in materials:
            aux = aux_map.get(mat, 0)
            mat_desc = desc_map.get(mat, '')
            mat_rows = grp[grp['物料号'] == mat]
            raw_unit = {}
            raw_ratio = {}
            for _, r in mat_rows.iterrows():
                code = _normalize_mat(r['原料号_raw'])
                qty = r['正向数量']
                amt = r['正向金额']
                raw_unit[code] = (amt / qty) if qty != 0 else None
                raw_ratio[code] = (qty / aux) if aux != 0 else None

            q3_price = _find_tsc_row_values(df_q3, mat, 'Q3实际单价', raw_cols, comp_col)
            q3_ratio = _find_tsc_row_values(df_q3, mat, 'Q3规格占比', raw_cols, comp_col)

            for label, source in [
                ('11月实际单价', raw_unit),
                ('Q3实际单价', q3_price),
                ('11月规格占比', raw_ratio),
                ('Q3规格占比', q3_ratio),
            ]:
                row = {'物料号': mat, '物料描述': mat_desc, '行类型': label, '分类': category_label}
                for _, code in raw_cols:
                    row[code] = source.get(code)
                row['综合单价'] = source.get('综合单价')
                rows.append(row)
        if not rows:
            return pd.DataFrame(columns=base_cols + raw_codes + ['综合单价'])
        return pd.DataFrame(rows)

    raw_usage_legs = build_matrix_rows(legs['物料号'].unique(), raw_cols_legs, comp_col_legs, df_q3_legs, '腿肉')
    raw_usage_breast = build_matrix_rows(breast['物料号'].unique(), raw_cols_breast, comp_col_breast, df_q3_breast, '胸肉')

    # Build BB2-style detail sheets
    totals = {}
    for mat, v in agg.items():
        aux = v['调整后实际量']
        unit = v['调整后实际额'] / aux if aux != 0 else 0
        scrap_ratio = abs(v['碎肉量']) / aux if aux != 0 else 0
        totals[mat] = {
            '辅助': aux,
            '总单价': unit,
            '碎肉占比': scrap_ratio,
            '修形利用率': (v['入库数量'] / aux) if aux != 0 else 0,
            '失水率': (1 - (v['入库数量'] / aux) - scrap_ratio) if aux != 0 else 0,
        }

    bb2_rows = []
    for _, row in df.iterrows():
        mat = row['物料号']
        if mat not in totals:
            continue
        total = totals[mat]
        raw = row['原料号']
        raw_desc = row['原料描述']
        is_header = raw == ''
        is_labor = raw == '人工费用' or ('人工' in raw_desc)
        is_scrap = row['实际数量'] < 0

        adj_qty = row['实际数量'] if (not is_header and not is_labor and not is_scrap) else 0
        adj_amt = row['实际金额'] if (not is_header and not is_labor and not is_scrap) else 0
        scrap_qty = row['实际数量'] if is_scrap else 0

        bb2_rows.append({
            '物料号': mat,
            '分类': mapping.get(mat, '其他'),
            '物料描述(不含琵琶腿/全腿和无抗）': row['物料描述'],
            '入库数量': row['入库数量'] if is_header else (agg[mat]['入库数量'] if is_labor else 0),
            '原料号': raw,
            '原料描述': raw_desc,
            '实际数量': row['实际数量'] if not is_header else 0,
            '实际金额': row['实际金额'] if not is_header else 0,
            '配方数量': row['配方数量'] if not is_header else 0,
            '调整后实际量': adj_qty,
            '辅助': total['辅助'],
            '调整后实际额': adj_amt,
            '碎肉量': scrap_qty,
            '修形前原料占比': (adj_qty / total['辅助']) if total['辅助'] != 0 else 0,
            '修形前原料单价': (adj_amt / adj_qty) if adj_qty != 0 else 0,
            '碎肉占比': (-scrap_qty / total['辅助']) if (is_scrap and total['辅助'] != 0) else (total['碎肉占比'] if is_labor else 0),
            '修形利用率': total['修形利用率'] if is_labor else 0,
            '失水率': total['失水率'] if is_labor else 0,
        })

    bb2 = pd.DataFrame(bb2_rows)
    if '物料号' not in bb2.columns:
        bb2 = pd.DataFrame(columns=['物料号'])
    bb2 = bb2[bb2['物料号'].astype(str).str.startswith('3900')].copy()
    order_map = {'胸肉': 0, '腿肉': 1, '其他': 2}
    bb2['分类序'] = bb2['分类'].map(order_map).fillna(2)
    bb2 = bb2.sort_values(['分类序', '物料号'])
    bb2 = bb2.drop(columns=['分类序'])
    bb2_legs = bb2[bb2['物料号'].isin(legs['物料号'])].copy()
    bb2_breast = bb2[bb2['物料号'].isin(breast['物料号'])].copy()

    return (
        legs,
        breast,
        other,
        bb2,
        bb2_legs,
        bb2_breast,
        raw_usage_legs,
        raw_usage_breast,
        spec_map_legs,
        spec_map_breast,
    )


def to_excel_bytes(
    legs,
    breast,
    bb2_all,
    bb2_legs,
    bb2_breast,
    raw_usage_legs,
    raw_usage_breast,
    spec_map_legs,
    spec_map_breast,
    prefix,
):
    output = io.BytesIO()
    fmt_money = '{:.2f}'
    fmt_pct = '{:.0%}'
    fmt_int = '{:.0f}'

    def apply_formats(df):
        df = df.copy()
        df = df.drop(columns=[c for c in HIDE_COLS if c in df.columns])
        df = df.drop(columns=['行序'], errors='ignore')
        for col in ['修形前原料综合耗用单价', '半成品原料成本', '半成品修形人工成本', '半成品总成本']:
            if col in df.columns:
                df[col] = df[col].apply(lambda x: fmt_money.format(x) if pd.notna(x) else '')
        for col in ['修形利用率', '损耗率']:
            if col in df.columns:
                df[col] = df[col].apply(lambda x: fmt_pct.format(x) if pd.notna(x) else '')
        return df

    legs = apply_formats(legs)
    breast = apply_formats(breast)

    def build_tsc_sheet(tsc_df, raw_usage_df, spec_map):
        base_cols = ['物料号', '物料描述', '行类型', '分类']
        raw_cols = [c for c in raw_usage_df.columns if c not in base_cols + ['综合单价']]
        raw_cols = [c for c in raw_cols if c != '综合单价']

        merged = tsc_df.merge(
            raw_usage_df,
            on=['物料号', '物料描述', '行类型'],
            how='left',
            suffixes=('', '_raw'),
        )
        merged['产品族'] = ''
        merged['修行后原料'] = merged['物料号']
        merged['使用半成品规格'] = merged['物料描述']

        ordered_cols = (
            ['产品族', '修行后原料', '使用半成品规格', '行类型']
            + raw_cols
            + ['综合单价']
            + ['修形前原料综合耗用单价', '修形利用率', '损耗率', '半成品原料成本', '半成品修形人工成本', '半成品总成本']
        )
        # Keep only columns that exist
        ordered_cols = [c for c in ordered_cols if c in merged.columns]
        header1 = {c: c for c in ordered_cols}
        header2 = {c: '' for c in ordered_cols}
        for code in raw_cols:
            header2[code] = spec_map.get(code, '')
        if '综合单价' in ordered_cols:
            header2['综合单价'] = '综合单价'
        blank = {c: '' for c in ordered_cols}
        merged_vals = merged[ordered_cols].replace({0: ''})
        return pd.concat(
            [pd.DataFrame([header1, header2, blank]), merged_vals],
            ignore_index=True,
        )

    tsc_legs = build_tsc_sheet(legs, raw_usage_legs, spec_map_legs)
    tsc_breast = build_tsc_sheet(breast, raw_usage_breast, spec_map_breast)
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        bb2_all.to_excel(writer, index=False, sheet_name='半成品')
        bb2_legs.to_excel(writer, index=False, sheet_name=f'{prefix}腿肉')
        tsc_legs.to_excel(writer, index=False, sheet_name='腿肉TSC')
        bb2_breast.to_excel(writer, index=False, sheet_name=f'{prefix}胸肉')
        tsc_breast.to_excel(writer, index=False, sheet_name='胸肉TSC')
    return output.getvalue()


if file_compare and file_rawlist and file_q3:
    st.session_state.pop('download_data', None)
    try:
        (
            legs_df,
            breast_df,
            other_df,
            bb2_all,
            bb2_legs,
            bb2_breast,
            raw_usage_legs,
            raw_usage_breast,
            spec_map_legs,
            spec_map_breast,
        ) = compute(file_compare, file_rawlist, file_q3)
    except Exception as exc:
        st.error(f'计算失败: {exc}')
        # column preview removed per request
    else:
        display_legs = _drop_hidden_cols(legs_df)
        display_breast = _drop_hidden_cols(breast_df)
        display_other = _drop_hidden_cols(other_df)

        st.success('计算完成')
        if not display_legs.empty:
            st.subheader('腿肉')
            st.dataframe(display_legs.style.format(FMT_DISPLAY), use_container_width=True)
        if not display_breast.empty:
            st.subheader('胸肉')
            st.dataframe(display_breast.style.format(FMT_DISPLAY), use_container_width=True)
        if not display_other.empty:
            st.subheader('其他')
            st.dataframe(display_other.style.format(FMT_DISPLAY), use_container_width=True)
        if display_legs.empty and display_breast.empty and display_other.empty:
            st.info('本次计算没有可显示的数据。')

        try:
            prefix = (
                file_compare.name.split('_')[0]
                if '_' in file_compare.name
                else file_compare.name.split('.')[0]
            )
            data = to_excel_bytes(
                display_legs,
                display_breast,
                bb2_all,
                bb2_legs,
                bb2_breast,
                raw_usage_legs,
                raw_usage_breast,
                spec_map_legs,
                spec_map_breast,
                prefix,
            )
            has_any = (
                (not display_legs.empty)
                or (not display_breast.empty)
                or (not display_other.empty)
                or (not bb2_all.empty)
                or (not bb2_legs.empty)
                or (not bb2_breast.empty)
                or (not raw_usage_legs.empty)
                or (not raw_usage_breast.empty)
            )
            if has_any:
                st.session_state['download_data'] = data
        except Exception as exc:
            st.error(f'下载生成失败: {exc}')
else:
    st.info('请先上传三个Excel文件。')

if 'download_data' in st.session_state:
    st.download_button(
        label='下载结果 Excel',
        data=st.session_state['download_data'],
        file_name='11月比较_全部物料_参数计算.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )
