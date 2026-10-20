#%% TEST CONNECTION TO INTERACTIVE WINDOW

print("Test Interactive Window")

#%% LOAD PACKAGES

import win32com.client
from win32com.client import GetActiveObject
import pandas as pd
import os

# %% INITIALIZE LEAP

def connect_to_leap():

    # Option 1 (open leap area first):
    leap = GetActiveObject("LEAP.LEAPApplication")

    print("Connected!")
    print("Active area:", leap.ActiveArea)

    # Option 2 (pull active area up): 
    # leap = win32com.client.gencache.EnsureDispatch("LEAP.LEAPApplication")
    # leap.ActiveArea =  AreaName  # must match the area name exactly

    leap.ActiveView = "Analysis"
    leap.EnableControls()
    leap.Visible = True

    return leap

# area_name = "Transport_Draft1" # for option 2
leap = connect_to_leap()


#%% LOAD DATA

def load_demands(Url, SheetName):
    
    url = Url.replace("edit?usp=sharing", "gviz/tq?tqx=out:csv&headers=1&sheet=")
    url = f"{url}{SheetName}"
    demand_branches = pd.read_csv(url)

    return demand_branches
    
url = "https://docs.google.com/spreadsheets/d/1u8pPUcBDZaqUYiQd-C2uEHluAH_yhs2KcldS81_jaVM/edit?usp=sharing"
sheet_name = "demand_branches_stocks"
demand_branches = load_demands(url, sheet_name)

#%% CLEAN NAMES

def clean_names(DemandBranches):

    DemandBranches['Branch'] = DemandBranches['Branch'].str.replace(' ', '')
    DemandBranches['Branch'] = DemandBranches['Branch'].str.replace('_','')

    Types = {name: group for name, group in DemandBranches.groupby('Type')}
    category = Types['Category']
    # br_name = category.iloc[:,0]
    # br_scale = category.iloc[:,3]
    # br_units = category.iloc[:,4]
    br_name = category['Branch']
    br_scale = category['Scale']
    br_units = category['Units']
    return br_name, br_scale, br_units

br_name, br_scale, br_units = clean_names(demand_branches)


#%% GET EXISTING BRANCHES

def get_all_branches(leap):
    branches = []
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
                branches.append({
                    "Name": child.Name,
                    "FullName": child.FullName,
                    "ID": branch_id,
                    "Depth": child.FullName.count("\\")
                })

            traverse(child)

    # Traverse the top-level branches
    for i in range(1, leap.Branches.Count + 1):
        branch = leap.Branches(i)

        if branch is None:
            continue

        branch_id = branch.ID

        if branch_id not in seen_ids:
            seen_ids.add(branch_id)
            branches.append({
                "Name": branch.Name,
                "FullName": branch.FullName,
                "ID": branch_id,
                "Depth": branch.FullName.count("\\")
            })

        traverse(branch)

    return pd.DataFrame(
        branches,
        columns=["Name", "FullName", "ID", "Depth"]
    )


# Get Demand Branches
all_branches = get_all_branches(leap)

all_branch_names = all_branches["FullName"]

demand_branch_names = all_branch_names[
    all_branch_names.str.contains("Demand", case=False, na=False)
]

exist_demand_branches = set(demand_branch_names)

print(all_branches.to_string(index=False))


#%% CREATES A SINGLE DEMAND BRANCH

def create_demand_branch(BranchIndex,ExistDemandBranches):

    name = br_name.iloc[BranchIndex]
    parent, child = name.rsplit("\\", 1)
    print("TESTING:     " ,name)

    if br_scale.iloc[BranchIndex] == None:
        scale = ""
    else:
        scale = br_scale.iloc[BranchIndex]

    if br_units.iloc[BranchIndex] == None:
        units = ""
    else: 
        units = br_units.iloc[BranchIndex]

    # Test to make sure parent branch exists and branch has not already been created
    # If either condition isn't satisfied, skips to next line
    if parent in ExistDemandBranches and name not in ExistDemandBranches:

        parent = leap.Branches(parent)

        new_cat = leap.AddCategory(
            parent.ID,
            child,
            scale,
            units
        )

        print("Branch created: ", name) 

        ExistDemandBranches.add(name)
    else:
        new_cat = f"{"Skipping:  "}{name}"
        print(new_cat)

    return new_cat


#%% CREATES ALL DEMAND BRANCHES

def create_all_demand():
    
    BranchIndex = 0 
    while BranchIndex <= len(br_name):
        print(BranchIndex)
        new_cat = create_demand_branch(BranchIndex, exist_demand_branches)
        BranchIndex = BranchIndex + 1 
    return BranchIndex

BranchIndex = create_all_demand()

