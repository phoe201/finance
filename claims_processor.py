# claims_processor.py

import pandas as pd
from io import BytesIO

def process_claims_excel(file_bytes):
    xls = pd.ExcelFile(BytesIO(file_bytes))
    sheet1_df = pd.read_excel(xls, sheet_name=0)

    start_row = sheet1_df[sheet1_df.iloc[:, 0] == 'Sub A/c Code'].index[0]
    sheet1_df = pd.read_excel(xls, sheet_name=0, skiprows=start_row + 2)
    sheet1_df = sheet1_df[~sheet1_df["TC"].str.lower().str.startswith("fx", na=False)]

    # Extract Document Dates and Codes
    codes, dates, current_code = [], [], None
    for item in sheet1_df["Document Date"]:
        if str(item).startswith("600"):
            current_code = item
        elif item == "":
            current_code = None
        else:
            if current_code:
                codes.append(current_code)
                dates.append(item)
    new_df = pd.DataFrame({"Document Date": dates, "Sub A/c Code": codes}).dropna(subset=["Document Date"])

    # Extract TC and Sub A/c Name
    names, tcs, current_name = [], [], None
    for item in sheet1_df["TC"]:
        if not str(item).startswith(("FX", "PV", "CL", "RV", "JV")):
            current_name = item
        elif item == "":
            current_name = None
        else:
            if current_name:
                names.append(current_name)
                tcs.append(item)
    new_df["Sub A/c Name"] = names
    new_df["TC"] = tcs

    # Currency Code and Document Number
    current_code, curr_code_list, doc_number_list = None, [], []
    for _, row in sheet1_df.iterrows():
        doc_date = str(row["Document Date"])
        doc_number = row["Document Number"]
        if doc_date.startswith("600"):
            current_code = doc_number
        elif current_code is not None:
            curr_code_list.append(current_code)
            doc_number_list.append(doc_number)
    curr_df = pd.DataFrame({"Currency Code": curr_code_list, "Document Number": doc_number_list}).dropna()
    new_df["Currency Code"] = curr_df["Currency Code"]
    new_df["Document Number"] = curr_df["Document Number"]

    # Other columns: Reference, Division, Dept, Narration, Debits, Credits
    for col in ["Reference", "Division", "Dept", "Narration", "Debits", "Credits"]:
        temp_df = sheet1_df.dropna(subset=[col]).reset_index(drop=True)
        new_df = new_df.reset_index(drop=True)
        new_df[col] = temp_df[col]

    # Calculate Balance
    new_df["Balance"] = new_df["Debits"] - new_df["Credits"]

    # Filter valid rows
    required_columns = ["Sub A/c Code", "Reference", "Narration", "Balance"]
    if not all(col in new_df.columns for col in required_columns):
        raise ValueError("Missing required columns")

    # Match Claims by Reference + Narration
    new_df["Matching"] = ""
    new_df["Reference"] = new_df["Reference"].fillna("").astype(str)
    new_df["Narration"] = new_df["Narration"].fillna("").astype(str)
    new_df["Sub A/c Code"] = new_df["Sub A/c Code"].astype(str)

    grouped = new_df.groupby("Sub A/c Code")
    for sub_code, group in grouped:
        temp_df = group.copy().reset_index()
        matched_indexes = set()
        for i, row in temp_df.iterrows():
            ref = row["Reference"]
            if not ref:
                continue
            for j, other_row in temp_df.iterrows():
                if i == j:
                    continue
                if ref in other_row["Narration"]:
                    total_balance = row["Balance"] + other_row["Balance"]
                    if abs(total_balance) < 1e-2:
                        matched_indexes.update([i, j])
        new_df.loc[temp_df.loc[list(matched_indexes), "index"], "Matching"] = "Matching Zero Customer"

    return new_df.to_dict(orient="records")
