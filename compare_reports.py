import pandas as pd
import requests
import json
import re
import os

# =====================================================
# AUTHENTICATION
# =====================================================


# =====================================================
# CONFIGURATION
# =====================================================

PREVIOUS_FILE = "input/previous_month.xlsx"
CURRENT_FILE = "input/current_month.xlsx"
OUTPUT_FILE = "output/entity_changes.xlsx"

ENTITY_ID = "CorporationID"
REG_COLUMN = "RegistrationNumber"

MAIN_REPORT_API_URL = "https://your-main-report-api-url"
STATUS_API_URL = "https://your-status-api-url"

AUTH_TOKEN = "Bearer YOUR_TOKEN_HERE"

ADDRESS_COLUMNS = [
    "SubsidiaryEntityRel.CurrentMainRegisteredAddress_SNum",
    "SubsidiaryEntityRel.CurrentMainRegisteredAddress_SN",
    "SubsidiaryEntityRel.CurrentMainRegisteredAddress_NativeSN",
    "SubsidiaryEntityRel.CurrentMainRegisteredAddress_E",
    "SubsidiaryEntityRel.CurrentMainRegisteredAddress_NativeE",
    "SubsidiaryEntityRel.CurrentMainRegisteredAddress_City",
    "SubsidiaryEntityRel.CurrentMainRegisteredAddress_NativeCity",
    "SubsidiaryEntityRel.CurrentMainRegisteredAddress_Country",
]

COMPARE_COLUMNS = [
    "EntityName",
    "RegistrationNumber"
]

# =====================================================
# NORMALIZATION FUNCTIONS
# =====================================================

def normalize_text(value):
    if pd.isna(value):
        return ""
    value = str(value)
    value = re.sub(r"\s+", " ", value)
    value = re.sub(r",+", ",", value)
    value = value.strip(" ,")
    return value.lower().strip()

def merge_address(row):
    parts = []
    for col in ADDRESS_COLUMNS:
        val = row.get(col, "")
        if pd.notna(val) and str(val).strip() != "":
            parts.append(str(val).strip())
    merged = ", ".join(parts)
    return normalize_text(merged)

# =====================================================
# LOAD EXCEL FILES
# =====================================================

print("Loading Excel files...")

prev_df = pd.read_excel(PREVIOUS_FILE)
curr_df = pd.read_excel(CURRENT_FILE)

# Remove blank or null CorporationID rows
prev_df = prev_df[prev_df[ENTITY_ID].notna()]
curr_df = curr_df[curr_df[ENTITY_ID].notna()]

prev_df = prev_df[prev_df[ENTITY_ID].astype(str).str.strip() != ""]
curr_df = curr_df[curr_df[ENTITY_ID].astype(str).str.strip() != ""]

# Normalize ID
prev_df[ENTITY_ID] = prev_df[ENTITY_ID].astype(str).str.strip()
curr_df[ENTITY_ID] = curr_df[ENTITY_ID].astype(str).str.strip()

# Check duplicates
if prev_df[ENTITY_ID].duplicated().any():
    raise Exception("Duplicate CorporationID found in previous file.")

if curr_df[ENTITY_ID].duplicated().any():
    raise Exception("Duplicate CorporationID found in current file.")

# =====================================================
# CALL MAIN REPORT API (FULL DATASET)
# =====================================================

print("Calling Main Report API...")

main_response = requests.get(
    MAIN_REPORT_API_URL,
    headers={"Authorization": AUTH_TOKEN},
    verify=False
)

main_data = main_response.json()

if isinstance(main_data, str):
    main_data = json.loads(main_data)

main_df = pd.DataFrame(main_data)

main_df[ENTITY_ID] = main_df[ENTITY_ID].astype(str).str.strip()
main_df = main_df[[ENTITY_ID, REG_COLUMN]]

# Merge and recover blank RegistrationNumber
curr_df = curr_df.merge(
    main_df,
    on=ENTITY_ID,
    how="left",
    suffixes=("", "_API")
)

curr_df[REG_COLUMN] = curr_df[REG_COLUMN].fillna(
    curr_df[f"{REG_COLUMN}_API"]
)

curr_df.loc[
    curr_df[REG_COLUMN].astype(str).str.strip() == "",
    REG_COLUMN
] = curr_df[f"{REG_COLUMN}_API"]

curr_df.drop(columns=[f"{REG_COLUMN}_API"], inplace=True)

# =====================================================
# NORMALIZE FIELDS
# =====================================================

prev_df["NormalizedAddress"] = prev_df.apply(merge_address, axis=1)
curr_df["NormalizedAddress"] = curr_df.apply(merge_address, axis=1)

for col in COMPARE_COLUMNS:
    prev_df[col] = prev_df[col].apply(normalize_text)
    curr_df[col] = curr_df[col].apply(normalize_text)

# =====================================================
# DETECT CHANGES
# =====================================================

print("Detecting changes...")

prev_ids = set(prev_df[ENTITY_ID])
curr_ids = set(curr_df[ENTITY_ID])

added_ids = curr_ids - prev_ids
deleted_ids = prev_ids - curr_ids
common_ids = prev_ids & curr_ids

added = curr_df[curr_df[ENTITY_ID].isin(added_ids)].copy()
deleted = prev_df[prev_df[ENTITY_ID].isin(deleted_ids)].copy()

modified_list = []

for entity_id in common_ids:
    prev_row = prev_df[prev_df[ENTITY_ID] == entity_id].iloc[0]
    curr_row = curr_df[curr_df[ENTITY_ID] == entity_id].iloc[0]

    changed = False

    for col in COMPARE_COLUMNS:
        if prev_row[col] != curr_row[col]:
            changed = True
            break

    if prev_row["NormalizedAddress"] != curr_row["NormalizedAddress"]:
        changed = True

    if changed:
        modified_list.append(curr_row)

modified = pd.DataFrame(modified_list)

# =====================================================
# CALL STATUS API (FULL DATASET)
# =====================================================

print("Calling Status API...")

status_response = requests.get(
    STATUS_API_URL,
    headers={"Authorization": AUTH_TOKEN},
    verify=False
)

status_data = status_response.json()

if isinstance(status_data, str):
    status_data = json.loads(status_data)

status_df = pd.DataFrame(status_data)

status_df[ENTITY_ID] = status_df[ENTITY_ID].astype(str).str.strip()
status_df = status_df[[ENTITY_ID, "Status", "WhenModified"]]

# =====================================================
# ENRICH WITH STATUS
# =====================================================

def enrich(df):
    if df.empty:
        return df
    return df.merge(status_df, on=ENTITY_ID, how="left")

added = enrich(added)
deleted = enrich(deleted)
modified = enrich(modified)

# =====================================================
# FINALIZE OUTPUT
# =====================================================

added["ChangeType"] = "ADDED"
deleted["ChangeType"] = "DELETED"
modified["ChangeType"] = "MODIFIED"

final_df = pd.concat([added, deleted, modified], ignore_index=True)

change_order = {"ADDED": 1, "DELETED": 2, "MODIFIED": 3}
final_df["SortOrder"] = final_df["ChangeType"].map(change_order)

final_df = final_df.sort_values(["SortOrder", ENTITY_ID])
final_df.drop(columns=["SortOrder"], inplace=True)

# =====================================================
# SAVE OUTPUT
# =====================================================

os.makedirs("output", exist_ok=True)
final_df.to_excel(OUTPUT_FILE, index=False)

print("====================================")
print("Reconciliation Complete")
print(f"Added: {len(added)}")
print(f"Deleted: {len(deleted)}")
print(f"Modified: {len(modified)}")
print(f"Output saved to: {OUTPUT_FILE}")
print("====================================")