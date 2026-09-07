"""Competitor classification, parent brands, and SOS calculations."""




def is_competitor(series):
    return series.astype(str).str.contains('COMPETITOR', case=False, na=False)


def extract_parent_brand(brand_name, is_competitor=False):
    """Ekstrak nama brand induk dari nama brand detail.
    Competitor brands tetap pakai nama lengkap.
    Contoh Indofood: 'INDOMIE SOTO MIE-SM' -> 'INDOMIE'
    """
    brand = str(brand_name).strip().upper()
    if is_competitor:
        return brand
    KNOWN_PARENTS = ['SARIMI GELAS', 'CAP 3 AYAM', 'POP MIE', 'IND MIE']
    for parent in KNOWN_PARENTS:
        if brand.startswith(parent):
            return 'INDOMIE' if parent == 'IND MIE' else parent
    return brand.split()[0] if brand else brand


def calc_sos(df, groupby_cols):
    """Hitung SOS% untuk sembarang groupby. Return df dengan fi, fk, total, SOS%."""
    df_i = df[~is_competitor(df['Produsen'])]
    df_k = df[ is_competitor(df['Produsen'])]

    fi = df_i.groupby(groupby_cols)['Facing'].sum().reset_index().rename(columns={'Facing': 'fi'})
    fk = df_k.groupby(groupby_cols)['Facing'].sum().reset_index().rename(columns={'Facing': 'fk'})

    merged = fi.merge(fk, on=groupby_cols, how='outer').fillna(0)
    merged['total'] = merged['fi'] + merged['fk']
    merged['SOS%']  = (
        merged['fi'] / merged['total'].replace(0, float('nan')) * 100
    ).round(1).fillna(0)
    return merged


def hitung_sos(df, groupby_cols):
    """Untuk STORE DETAIL — kolom facing_indofood/kompetitor/SOS_% per Period."""
    g = groupby_cols + ['Period']
    df_i = df[~is_competitor(df['Produsen'])]
    df_k = df[ is_competitor(df['Produsen'])]

    fi = df_i.groupby(g)['Facing'].sum().reset_index().rename(columns={'Facing': 'facing_indofood'})
    fk = df_k.groupby(g)['Facing'].sum().reset_index().rename(columns={'Facing': 'facing_kompetitor'})

    merged = fi.merge(fk, on=g, how='outer').fillna(0)
    merged['total_facing'] = merged['facing_indofood'] + merged['facing_kompetitor']
    merged['SOS_%'] = (
        merged['facing_indofood'] / merged['total_facing'].replace(0, float('nan')) * 100
    ).round(1).fillna(0)
    return merged
