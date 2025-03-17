from rapidfuzz import process, fuzz
import pandas as pd
from data_io import import_data, export_data

def find_best_match(hotel_name:str, competitor_hotels:str):
    if pd.isna(hotel_name) or not isinstance(hotel_name, str):  # Handle NaN or non-string values
        return pd.Series([None, 0])
    
    match = process.extractOne(hotel_name, competitor_hotels, scorer=fuzz.token_sort_ratio)
    
    if match:
        return pd.Series([match[0], match[1]])  # Extract match name and score
    return pd.Series([None, 0])


# Enter data
ro_df, summary = import_data('Test.xlsx', 'Test.xlsx', 'RO', 'SAMO')

# Apply fuzzy matching
ro_df[['matched_hotel', 'match_score']] = ro_df['ro_hotel'].apply(
    find_best_match, competitor_hotels=summary['samo_hotel'].dropna().tolist()
)

# Merge with competitor prices
df_merged = ro_df.merge(summary, left_on='matched_hotel', right_on='samo_hotel', how='left')

export_data(df_merged, 'matched_hotels.xlsx')