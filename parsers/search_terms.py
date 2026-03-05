import pandas as pd


# These are the thresholds that determine what we flag as actionable.
# Adjust these based on your target ACoS and acceptable spend levels.
MIN_CLICKS_FOR_NEGATIVE = 8     # flag a term as wasted spend if it got more than this many clicks with no sales
TARGET_ACOS = 0.40              # 40% — terms above this are underperforming
MIN_CVR_FOR_EXACT = 0.10        # 10% conversion rate — strong enough to add as exact match


def load_search_terms(filepath):
    df = pd.read_excel(filepath)

    # Strip whitespace from column names — Amazon likes to add trailing spaces
    df.columns = df.columns.str.strip()

    # Rename the columns we care about to cleaner names
    df = df.rename(columns={
        'Customer Search Term': 'search_term',
        'Spend': 'spend',
        '7 Day Total Orders (#)': 'orders',
        '7 Day Total Sales': 'sales',
        'Total Advertising Cost of Sales (ACOS)': 'acos',
        'Total Return on Advertising Spend (ROAS)': 'roas',
        'Clicks': 'clicks',
        'Impressions': 'impressions',
        '7 Day Conversion Rate': 'cvr',
        'Campaign Name': 'campaign',
        'Ad Group Name': 'ad_group',
        'Match Type': 'match_type',
    })

    # When there are zero sales, ACOS comes in as NaN — set it to a high number
    # so these terms naturally float to the top of our waste analysis
    df['acos'] = df['acos'].fillna(999)
    df['sales'] = df['sales'].fillna(0)
    df['orders'] = df['orders'].fillna(0)
    df['cvr'] = df['cvr'].fillna(0)

    return df


def find_negative_keyword_candidates(df):
    """
    Terms where we spent real money but got zero orders.
    These are draining your budget with nothing to show for it.
    """
    wasted = df[
        (df['clicks'] >= MIN_CLICKS_FOR_NEGATIVE) &
        (df['orders'] == 0)
    ].copy()

    wasted = wasted[['search_term', 'spend', 'clicks', 'campaign', 'ad_group', 'match_type']]
    wasted = wasted.sort_values('spend', ascending=False)

    return wasted


def find_exact_match_candidates(df):
    """
    Terms with strong conversion rates — worth adding as exact match
    keywords so you can bid on them directly and control spend more precisely.
    """
    strong = df[
        (df['cvr'] >= MIN_CVR_FOR_EXACT) &
        (df['orders'] >= 1)
    ].copy()

    strong = strong[['search_term', 'cvr', 'orders', 'spend', 'acos', 'campaign', 'ad_group']]
    strong = strong.sort_values('cvr', ascending=False)

    return strong


def find_high_spend_low_performance(df):
    """
    Terms where you're spending but ACoS is way above target.
    These need a bid reduction or closer review before cutting entirely.
    """
    underperforming = df[
        (df['clicks'] >= MIN_CLICKS_FOR_NEGATIVE) &
        (df['acos'] > TARGET_ACOS) &
        (df['acos'] < 999)  # exclude the zero-sale terms, those are handled above
    ].copy()

    underperforming = underperforming[['search_term', 'spend', 'acos', 'orders', 'campaign', 'ad_group']]
    underperforming = underperforming.sort_values('acos', ascending=False)

    return underperforming


def find_low_acos_terms(df):
    """
    Terms performing well under 20% ACoS — these are your winners.
    Consider increasing bids to get more volume out of them.
    """
    low_acos = df[
        (df['acos'] < 0.20) &
        (df['acos'] > 0) &     # exclude zero-sale rows where acos was set to 999
        (df['orders'] >= 1)
    ].copy()

    low_acos = low_acos[['search_term', 'spend', 'acos', 'orders', 'campaign', 'ad_group']]
    low_acos = low_acos.sort_values(['campaign', 'ad_group', 'spend'], ascending=[True, True, False])

    return low_acos


def find_very_high_acos_terms(df):
    """
    Terms above 60% ACoS — likely burning money faster than they're converting.
    Review before cutting in case they're volume drivers.
    """
    very_high = df[
        (df['acos'] > 0.60) &
        (df['acos'] < 999) &   # exclude zero-sale rows
        (df['orders'] >= 1)
    ].copy()

    very_high = very_high[['search_term', 'spend', 'acos', 'orders', 'campaign', 'ad_group']]
    very_high = very_high.sort_values(['campaign', 'ad_group', 'spend'], ascending=[True, True, False])

    return very_high


def summarize(filepath):
    df = load_search_terms(filepath)

    negatives = find_negative_keyword_candidates(df)
    exact_matches = find_exact_match_candidates(df)
    underperforming = find_high_spend_low_performance(df)
    low_acos = find_low_acos_terms(df)
    very_high_acos = find_very_high_acos_terms(df)

    total_spend = df['spend'].sum()

    summary = {
        'total_spend': round(total_spend, 2),
        'wasted_spend': round(negatives['spend'].sum(), 2),
        'wasted_pct': round(negatives['spend'].sum() / total_spend * 100, 1) if total_spend > 0 else 0,
        'negative_candidates': negatives.to_dict(orient='records'),
        'exact_match_candidates': exact_matches.to_dict(orient='records'),
        'high_acos_terms': underperforming.to_dict(orient='records'),
        'low_acos_terms': low_acos.to_dict(orient='records'),
        'very_high_acos_terms': very_high_acos.to_dict(orient='records'),
    }

    return summary


def export_to_xlsx(result, output_path='output/search_term_analysis.xlsx'):
    import os
    os.makedirs('output', exist_ok=True)

    def sort_rows(rows):
        return sorted(rows, key=lambda x: (x['campaign'], x['ad_group'], -x['spend']))

    def make_df(rows, columns):
        return pd.DataFrame([
            {col: row.get(col) for col in columns}
            for row in rows
        ], columns=columns)

    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:

        # --- Summary tab ---
        summary_data = {
            'Metric': ['Total Spend', 'Wasted Spend (0 orders)', 'Wasted % of Budget',
                       'Negative Candidates', 'Exact Match Candidates',
                       'High ACoS Terms (40-60%+)', 'Low ACoS Terms (<20%)', 'Very High ACoS Terms (>60%)'],
            'Value': [
                f"${result['total_spend']}",
                f"${result['wasted_spend']}",
                f"{result['wasted_pct']}%",
                len(result['negative_candidates']),
                len(result['exact_match_candidates']),
                len(result['high_acos_terms']),
                len(result['low_acos_terms']),
                len(result['very_high_acos_terms']),
            ]
        }
        pd.DataFrame(summary_data).to_excel(writer, sheet_name='Summary', index=False)

        # --- Negative Keywords tab ---
        neg_rows = sort_rows(result['negative_candidates'])
        neg_df = make_df(neg_rows, ['campaign', 'ad_group', 'search_term', 'spend', 'clicks', 'orders'])
        neg_df.columns = ['Campaign', 'Ad Group', 'Search Term', 'Spend', 'Clicks', 'Orders']
        neg_df['Spend'] = neg_df['Spend'].apply(lambda x: f"${x:.2f}")
        neg_df.to_excel(writer, sheet_name='Negative Keywords', index=False)

        # --- Exact Match Candidates tab ---
        exact_rows = sort_rows(result['exact_match_candidates'])
        exact_df = make_df(exact_rows, ['campaign', 'ad_group', 'search_term', 'cvr', 'orders', 'spend', 'acos'])
        exact_df.columns = ['Campaign', 'Ad Group', 'Search Term', 'CVR', 'Orders', 'Spend', 'ACoS']
        exact_df['CVR'] = exact_df['CVR'].apply(lambda x: f"{x*100:.1f}%")
        exact_df['Spend'] = exact_df['Spend'].apply(lambda x: f"${x:.2f}")
        exact_df['ACoS'] = exact_df['ACoS'].apply(lambda x: f"{x*100:.1f}%" if x < 999 else 'N/A')
        exact_df.to_excel(writer, sheet_name='Exact Match Candidates', index=False)

        # --- Low ACoS tab (<20%) ---
        low_df = make_df(result['low_acos_terms'], ['campaign', 'ad_group', 'search_term', 'spend', 'acos', 'orders'])
        low_df.columns = ['Campaign', 'Ad Group', 'Search Term', 'Spend', 'ACoS', 'Orders']
        low_df['Spend'] = low_df['Spend'].apply(lambda x: f"${x:.2f}")
        low_df['ACoS'] = low_df['ACoS'].apply(lambda x: f"{x*100:.1f}%")
        low_df.to_excel(writer, sheet_name='Low ACoS (under 20%)', index=False)

        # --- Very High ACoS tab (>60%) ---
        high_df = make_df(result['very_high_acos_terms'], ['campaign', 'ad_group', 'search_term', 'spend', 'acos', 'orders'])
        high_df.columns = ['Campaign', 'Ad Group', 'Search Term', 'Spend', 'ACoS', 'Orders']
        high_df['Spend'] = high_df['Spend'].apply(lambda x: f"${x:.2f}")
        high_df['ACoS'] = high_df['ACoS'].apply(lambda x: f"{x*100:.1f}%")
        high_df.to_excel(writer, sheet_name='Very High ACoS (over 60%)', index=False)

        # --- High ACoS tab (40-60%) ---
        mid_df = make_df(sort_rows(result['high_acos_terms']), ['campaign', 'ad_group', 'search_term', 'spend', 'acos', 'orders'])
        mid_df.columns = ['Campaign', 'Ad Group', 'Search Term', 'Spend', 'ACoS', 'Orders']
        mid_df['Spend'] = mid_df['Spend'].apply(lambda x: f"${x:.2f}")
        mid_df['ACoS'] = mid_df['ACoS'].apply(lambda x: f"{x*100:.1f}%")
        mid_df.to_excel(writer, sheet_name='High ACoS (40-60%)', index=False)

    print(f"\nExcel file exported to: {output_path}")


if __name__ == '__main__':
    result = summarize('data/Sponsored_Products_Search_term_report_60_days.xlsx')

    print(f"\nTotal Spend: ${result['total_spend']}")
    print(f"Wasted Spend (no orders): ${result['wasted_spend']} ({result['wasted_pct']}% of budget)")

    print(f"\n--- Negative Keyword Candidates ({len(result['negative_candidates'])}) ---")
    for row in result['negative_candidates'][:10]:
        print(f"  '{row['search_term']}'")
        print(f"    ${row['spend']:.2f} spent | {row['clicks']} clicks | 0 orders")
        print(f"    Campaign: {row['campaign']}")
        print(f"    Ad Group: {row['ad_group']}")

    print(f"\n--- Exact Match Candidates ({len(result['exact_match_candidates'])}) ---")
    for row in result['exact_match_candidates'][:10]:
        print(f"  '{row['search_term']}'")
        print(f"    {row['cvr']*100:.1f}% CVR | {row['orders']} orders | ${row['spend']:.2f} spent")
        print(f"    Campaign: {row['campaign']}")
        print(f"    Ad Group: {row['ad_group']}")

    print(f"\n--- High ACoS Terms ({len(result['high_acos_terms'])}) ---")
    for row in result['high_acos_terms'][:10]:
        print(f"  '{row['search_term']}'")
        print(f"    {row['acos']*100:.1f}% ACoS | ${row['spend']:.2f} spent | {row['orders']} orders")
        print(f"    Campaign: {row['campaign']}")
        print(f"    Ad Group: {row['ad_group']}")

    export_to_xlsx(result)