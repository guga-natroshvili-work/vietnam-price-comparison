from rapidfuzz import process, fuzz
import pandas as pd
import re
from data_io import import_data, export_data

ro_df, samo_df = import_data('CatSAMORO.xlsx', 'CatSAMORO.xlsx', 'RO', 'SAMO')


# Function to extract text inside parentheses
def extract_parentheses(text):
    return " | ".join(re.findall(r"\((.*?)\)", text)) or None  # Join multiple occurrences


# Function to remove text inside parentheses
def remove_parentheses(text):
    return re.sub(r"\s*\([^)]*\)", "", text).strip()


ro_df['ro_roomcat_stripped'] = ro_df['ro_roomcat'].apply(remove_parentheses) 
samo_df['samo_roomcat_stripped'] = samo_df['samo_roomcat'].apply(remove_parentheses)
ro_df['ro_roomcat_parentheses'] = ro_df['ro_roomcat'].apply(extract_parentheses)
samo_df['samo_roomcat_parentheses'] = samo_df['samo_roomcat'].apply(extract_parentheses)

# Apply fuzzy matching

# Define function to find best match
def find_best_match(hotel_name, competitor_hotels):
    if pd.isna(hotel_name) or not isinstance(hotel_name, str):  # Handle NaN or non-string values
        return pd.Series([None, 0])
    
    match = process.extractOne(hotel_name, competitor_hotels, scorer=fuzz.token_sort_ratio)
    
    if match:
        return pd.Series([match[0], match[1]])  # Extract match name and score
    return pd.Series([None, 0])


ro_df[['matched_room', 'match_score']] = ro_df['ro_roomcat_stripped'].apply(
    find_best_match, competitor_hotels=samo_df['samo_roomcat_stripped'].dropna().tolist()
)

# Merge with competitor prices and include city and id
df_merged = ro_df.merge(samo_df[['samo_roomcat_stripped','samo_roomcat_parentheses','samo_roomcat_id',]], 
                        left_on='matched_room', 
                        right_on='samo_roomcat_stripped', 
                        how='left')

# Drop redundant column and rename for clarity
df_merged = df_merged.drop(columns=['samo_roomcat_stripped'])

# Save the merged DataFrame to an Excel file
df_merged.to_excel('rooms_strippe.xlsx', index=False)


def find_best_match_parentheses(df_merged, threshold=85):
    """
    Finds the best match for `ro_roomcat_parentheses` using `samo_roomcat_parentheses`, 
    but only if they belong to the same `matched_room` category.
    
    Parameters:
    - df_merged: DataFrame containing room category data.
    - threshold: Minimum match score to consider a valid match.
    
    Returns:
    - DataFrame with a new column `best_match_parentheses` and `match_score_parentheses`.
    """
    df = df_merged.copy()


    # Group by `matched_room` to ensure we are comparing similar room categories
    def match_within_group(group):
        # Get unique lists of parentheses values for both columns
        ro_parentheses = group["ro_roomcat_parentheses"].dropna().unique()
        samo_parentheses = group["samo_roomcat_parentheses"].dropna().unique()

        match_results = {}

        for ro_value in ro_parentheses:
            if not isinstance(ro_value, str) or ro_value.strip() == "":
                continue  # Skip empty values

            best_match = process.extractOne(ro_value, samo_parentheses, scorer=fuzz.token_sort_ratio)
            if best_match and best_match[1] >= threshold:
                match_results[ro_value] = (best_match[0], best_match[1])
            else:
                match_results[ro_value] = (None, 0)  # No good match found

        # Apply matches back to the group
        group["best_match_parentheses"] = group["ro_roomcat_parentheses"].map(lambda x: match_results.get(x, (None, 0))[0])
        group["match_score_parentheses"] = group["ro_roomcat_parentheses"].map(lambda x: match_results.get(x, (None, 0))[1])

        return group

    df = df.groupby("matched_room", group_keys=False).apply(match_within_group)

    return df

# Example usage
df_updated = find_best_match_parentheses(df_merged)

export_data(df_updated, 'matched_categories.xlsx')