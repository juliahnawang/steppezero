#%% LOAD PACKAGES

import win32com.client
from win32com.client import GetActiveObject
import pandas as pd


#%% CONNECT TO LEAP

def connect_to_leap():
    leap = GetActiveObject("LEAP.LEAPApplication")
    leap.ActiveView = "Analysis"
    leap.EnableControls()
    leap.Visible = True

    print("Connected!")
    print("Active area:", leap.ActiveArea)

    return leap

leap = connect_to_leap()


#%% LOAD GOOGLE SHEETS

GOOGLE_SHEET_ID = "1u8pPUcBDZaqUYiQd-C2uEHluAH_yhs2KcldS81_jaVM"


def load_sheet(sheet_name):
    url = (
        f"https://docs.google.com/spreadsheets/d/{GOOGLE_SHEET_ID}/"
        f"gviz/tq?tqx=out:csv&headers=1&sheet={sheet_name}"
    )
    return pd.read_csv(url)


# Sheet 1: branch structure
branch_df = load_sheet("demand_branches_stocks")

# Sheet 2: yearly stock/mileage observations
data_df = load_sheet("transport_data")

#%% CLEAN AND VALIDATE BRANCH SHEET

required_branch_columns = {
    "Branch", "Type", "Scale", "Units", "Fuel"
}

missing = required_branch_columns - set(branch_df.columns)

if missing:
    raise ValueError(
        f"Branch sheet is missing columns: {sorted(missing)}. "
        "Add a Fuel column; leave it empty for Category rows."
    )


def clean_branch_sheet(df):
    df = df.copy()

    for col in ["Branch", "Type", "Scale", "Units", "Fuel"]:
        df[col] = df[col].fillna("").astype(str).str.strip()

    # Preserve spaces within branch names.
    # Only normalize accidental leading/trailing whitespace.
    df["Branch"] = df["Branch"].str.strip("\\")
    df = df[df["Branch"] != ""]

    if df["Branch"].duplicated().any():
        duplicates = df.loc[
            df["Branch"].duplicated(keep=False), "Branch"
        ].tolist()
        raise ValueError(f"Duplicate branch paths in sheet: {duplicates}")

    return df


branch_df = clean_branch_sheet(branch_df)


#%% GET EXISTING BRANCH PATHS

def get_all_branches(leap):
    rows = []
    seen_ids = set()

    def traverse(parent):
        children = parent.Children

        for i in range(1, children.Count + 1):
            child = children(i)

            if child is None:
                continue

            branch_id = child.ID

            if branch_id not in seen_ids:
                seen_ids.add(branch_id)
                rows.append({
                    "Name": child.Name,
                    "FullName": child.FullName,
                    "ID": branch_id
                })

            traverse(child)

    for i in range(1, leap.Branches.Count + 1):
        branch = leap.Branches(i)

        if branch is None:
            continue

        branch_id = branch.ID

        if branch_id not in seen_ids:
            seen_ids.add(branch_id)
            rows.append({
                "Name": branch.Name,
                "FullName": branch.FullName,
                "ID": branch_id
            })

        traverse(branch)

    return pd.DataFrame(rows, columns=["Name", "FullName", "ID"])


#%% CREATE BRANCHES IN PARENT-FIRST ORDER

def create_branches(leap, df, dry_run=True):
    """
    dry_run=True: report planned actions only.
    dry_run=False: create missing branches.

    Note: the technology creation call must match the intended LEAP
    branch methodology and the correct activity/fuel/energy units.
    """

    existing = set(get_all_branches(leap)["FullName"])
    work = df.copy()

    work["Depth"] = work["Branch"].str.count("\\\\")
    work = work.sort_values(["Depth", "Branch"], kind="stable")

    for _, row in work.iterrows():
        path = row["Branch"]
        branch_type = row["Type"]
        scale = row["Scale"]
        units = row["Units"]
        fuel = row["Fuel"]

        if path in existing:
            print(f"EXISTS:  {path}")
            continue

        if "\\" not in path:
            raise ValueError(
                f"Cannot determine parent for {path!r}. "
                "Provide a full path including its parent."
            )

        parent_path, child_name = path.rsplit("\\", 1)

        if parent_path not in existing:
            raise ValueError(
                f"Parent does not exist yet: {parent_path!r} "
                f"(required for {path!r})"
            )

        if dry_run:
            print(
                f"WOULD CREATE: {path} | type={branch_type} | "
                f"fuel={fuel or '(none)'}"
            )
            # Treat as planned so children can be checked in the same run.
            existing.add(path)
            continue

        parent = leap.Branch(parent_path)

        if parent is None:
            raise RuntimeError(f"LEAP returned no parent for {parent_path}")

        if branch_type.casefold() == "category":
            leap.AddCategory(
                parent.ID,
                child_name,
                scale,
                units
            )

        elif "transport technology (stock turnover method)" in branch_type.casefold():
            if not fuel:
                raise ValueError(
                    f"Technology {path!r} has no Fuel value."
                )

            raise NotImplementedError(
                "Confirm the correct LEAP API method for creating a "
                "transport Stock Turnover technology in this LEAP version. "
                "Do not create it as an ordinary Activity Analysis technology."
            )

        else:
            raise NotImplementedError(
                f"Branch type {branch_type!r} needs a specific creation "
                "method. No method was guessed for {path!r}."
            )

        # Refresh the path set after a successful creation.
        existing = set(get_all_branches(leap)["FullName"])
        print(f"CREATED: {path}")

    return existing


# First inspect planned actions without changing the model.
planned_branches = create_branches(leap, branch_df, dry_run=True)

# After confirming the plan and the technology-creation method:
# create_branches(leap, branch_df, dry_run=False)


#%% VALIDATE DATA SHEET FORMAT

required_data_columns = {
    "Branch", "Variable", "Year", "Value", "Region", "Scenario"
}

missing = required_data_columns - set(data_df.columns)

if missing:
    raise ValueError(
        f"Data sheet is missing columns: {sorted(missing)}"
    )

data_df = data_df.copy()

for col in ["Branch", "Variable", "Region", "Scenario"]:
    data_df[col] = data_df[col].fillna("").astype(str).str.strip()

data_df["Year"] = pd.to_numeric(data_df["Year"], errors="raise").astype(int)
data_df["Value"] = pd.to_numeric(data_df["Value"], errors="raise")

if data_df[
    ["Branch", "Variable", "Region", "Scenario"]
].eq("").any().any():
    raise ValueError(
        "Data sheet has empty Branch, Variable, Region, or Scenario values."
    )

key_columns = ["Branch", "Variable", "Year", "Region", "Scenario"]

if data_df.duplicated(key_columns).any():
    duplicates = data_df.loc[
        data_df.duplicated(key_columns, keep=False), key_columns
    ]
    raise ValueError(
        "Duplicate data records found:\n" + duplicates.to_string(index=False)
    )

known_branches = set(branch_df["Branch"])
unknown_data_branches = sorted(set(data_df["Branch"]) - known_branches)

if unknown_data_branches:
    print(
        "WARNING: data rows refer to branches not listed in the "
        "branch sheet:\n" + "\n".join(unknown_data_branches)
    )

print("Branch sheet validated:", len(branch_df), "rows")
print("Data sheet validated:", len(data_df), "rows")
print(data_df.head(10).to_string(index=False))