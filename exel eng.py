import pandas as pd

# === USER CONFIGURATION ===
# Set the column names exactly as they appear in your Excel file:
column_map = {
    'code': 'KOD_INTRASTAT',      # Column with CN / Intrastat code
    'btom': 'BTOM',               # Column with "low" or "medium"
    'net_weight': 'WAGA_NETTO',   # Column with net weight
    'box_count': 'LICZBA_BOXOW',  # Column with number of boxes
    'ehc': 'EHC'                  # Optional: certificate types (for medium only)
}

# === Load list of commodity code prefixes from a text file ===
def load_codes_from_file(path: str):
    with open(path, 'r', encoding='utf-8') as f:
        content = f.read()
    original_codes = [code.strip() for code in content.split(',') if code.strip()]
    code_map = {code.lstrip('0'): code for code in original_codes}
    return list(code_map.keys()), code_map

# === Match function (code starts with given prefix) ===
def match_code(kod, code_prefix):
    if not isinstance(kod, str):
        kod = str(kod)
    kod = kod.replace(' ', '').split('.')[0].lstrip('0')
    return kod.startswith(code_prefix)

# === Main processing function ===
def summarize_by_code_and_btom_with_ehc(file_path: str, sheet_name: str, code_starts: list):
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    df.columns = df.columns.astype(str)

    # Prepare KOD column
    df[column_map['code']] = df[column_map['code']].astype(str).str.replace(' ', '').str.lstrip('0')

    # Normalize BTOM column
    if column_map['btom'] in df.columns:
        df[column_map['btom']] = df[column_map['btom']].astype(str).str.lower()
    else:
        df[column_map['btom']] = ''

    results_by_btom = {}

    # === LOW CATEGORY ===
    df_low = df[df[column_map['btom']] == 'low']
    low_results = {}
    for code in code_starts:
        matched_df = df_low[df_low[column_map['code']].apply(lambda x: match_code(x, code))]
        if not matched_df.empty:
            low_results[code] = {
                'net_weight': round(matched_df[column_map['net_weight']].sum(), 2),
                'box_count': round(matched_df[column_map['box_count']].sum(), 2)
            }
    if low_results:
        results_by_btom['low'] = low_results

    # === MEDIUM CATEGORY (by EHC) ===
    df_medium = df[df[column_map['btom']] == 'medium'].copy()
    if column_map['ehc'] in df.columns:
        df_medium[column_map['ehc']] = df_medium[column_map['ehc']].fillna('none')
    else:
        df_medium[column_map['ehc']] = 'none'

    df_medium[column_map['ehc']] = df_medium[column_map['ehc']].astype(str).str.split(',')
    df_medium = df_medium.explode(column_map['ehc'])
    df_medium[column_map['ehc']] = df_medium[column_map['ehc']].str.strip()

    medium_results_by_ehc = {}
    for ehc_val in df_medium[column_map['ehc']].unique():
        df_ehc = df_medium[df_medium[column_map['ehc']] == ehc_val]
        ehc_group = {}
        for code in code_starts:
            matched_df = df_ehc[df_ehc[column_map['code']].apply(lambda x: match_code(x, code))]
            if not matched_df.empty:
                ehc_group[code] = {
                    'net_weight': round(matched_df[column_map['net_weight']].sum(), 2),
                    'box_count': round(matched_df[column_map['box_count']].sum(), 2)
                }
        if ehc_group:
            medium_results_by_ehc[ehc_val] = ehc_group
    if medium_results_by_ehc:
        results_by_btom['medium'] = medium_results_by_ehc

    # === FALLBACK if no BTOM column exists ===
    if 'low' not in results_by_btom and 'medium' not in results_by_btom:
        overall_results = {}
        for code in code_starts:
            matched_df = df[df[column_map['code']].apply(lambda x: match_code(x, code))]
            if not matched_df.empty:
                overall_results[code] = {
                    'net_weight': round(matched_df[column_map['net_weight']].sum(), 2),
                    'box_count': round(matched_df[column_map['box_count']].sum(), 2)
                }
        if overall_results:
            results_by_btom['total'] = overall_results

    return results_by_btom

# === ENTRY POINT ===
if __name__ == "__main__":
    file_path = "Spizarnia_78384_EUR_new.xls"              # Excel file with source data
    sheet_name = "Sheet0"                    # Sheet name inside the file
    code_file = "CommodityCodes.txt"         # File with list of commodity codes
    output_excel = "grouped_result.xlsx"     # Output file name

    codes, code_map = load_codes_from_file(code_file)
    results = summarize_by_code_and_btom_with_ehc(file_path, sheet_name, codes)

    # === SAVE TO EXCEL ===
    excel_data = {}
    for category, data in results.items():
        if category == 'medium':
            for ehc, group in data.items():
                rows = []
                for code, vals in group.items():
                    restored_code = code_map.get(code, code)
                    rows.append({
                        'CN_Code': restored_code,
                        'SUM_NET_WEIGHT': vals['net_weight'],
                        'BOX_COUNT': vals['box_count']
                    })
                sheet = 'medium' if ehc == 'none' else ehc[:31]
                excel_data[sheet] = pd.DataFrame(rows)
        elif category in ['low', 'total']:
            rows = []
            for code, vals in data.items():
                restored_code = code_map.get(code, code)
                rows.append({
                    'CN_Code': restored_code,
                    'SUM_NET_WEIGHT': vals['net_weight'],
                    'BOX_COUNT': vals['box_count']
                })
            excel_data[category] = pd.DataFrame(rows)

    with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
        for name, df_out in excel_data.items():
            df_out.to_excel(writer, index=False, sheet_name=name)

    print(f"\n✅ Results saved to file: {output_excel}")
