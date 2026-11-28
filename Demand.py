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

def load_branches(Url, SheetName):
    
    url = Url.replace("edit?usp=sharing", "gviz/tq?tqx=out:csv&headers=1&sheet=")
    url = f"{url}{SheetName}"
    demand_branches = pd.read_csv(url)

    return demand_branches


# OLD    
# url = "https://docs.google.com/spreadsheets/d/1u8pPUcBDZaqUYiQd-C2uEHluAH_yhs2KcldS81_jaVM/edit?usp=sharing"
# sheet_name = "demand_branches_stocks"

# NEW
# url = "https://docs.google.com/spreadsheets/d/1vUwZEsMsWO4dPNd7wPUkMoQXzTn2MAqGBbk1eelF74o/edit?usp=sharing"
# sheet_name = "demand_branches"

# CORRECTED
url = "https://docs.google.com/spreadsheets/d/1dnkUAZv-ZUIfi0NCcSjTSN5pypKsyzDA7xKBEYKQ1ck/edit?usp=sharing"
sheet_name = "demand_branches"


demand_branches = load_branches(url, sheet_name)

#%% CLEAN NAMES
def clean_names(DemandBranches):

    DemandBranches['Branch'] = DemandBranches['Branch'].str.replace(' ', '')
    DemandBranches['Branch'] = DemandBranches['Branch'].str.replace('_','')

    return DemandBranches

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

# Types = {name: group for name, group in demand_branches.groupby('Type')}

branch_type = demand_branches['Type']
branch_name = demand_branches['Branch']
branch_scale = demand_branches['Scale']
branch_activity_units = demand_branches['Activity Units']
branch_fuel = demand_branches['Fuel']
branch_energy_units = demand_branches['Energy Units']

def create_demand_branch( BranchIndex,ExistDemandBranches):

    name = branch_name.iloc[BranchIndex]
    parent, child = name.rsplit("\\", 1)

    print("TESTING:     " ,name)

    scale = branch_scale.iloc[BranchIndex]
    activity_units = branch_activity_units.iloc[BranchIndex]
    fuel = branch_fuel.iloc[BranchIndex]
    energy_units = branch_energy_units.iloc[BranchIndex]
    type = branch_type.iloc[BranchIndex]

    scale = "" if pd.isna(scale) else str(scale).strip()
    activity_units = "" if pd.isna(activity_units) else str(activity_units).strip()
    fuel = "" if pd.isna(fuel) else str(fuel).strip()
    energy_units = "" if pd.isna(energy_units) else str(energy_units).strip()
    type = "" if pd.isna(type) else str(type).strip()
    
    # Test to make sure parent branch exists and branch has not already been created
    # If either condition isn't satisfied, skips to next line
    if parent in ExistDemandBranches and name not in ExistDemandBranches:

        parent_id = leap.Branch(parent).ID

        if type == "Category":
            
            new_cat = leap.AddCategory(
                parent_id,
                child,
                scale,
                activity_units
            )

            new_branch = new_cat

        elif type == "Technology" or type == "Technology with Energy Intensity (Activity Analysis)" or "Transport Technology (Stock Turnover Method)":

            new_tech = leap.AddTechnology(
                parent_id,
                child,
                scale,
                activity_units,
                fuel,
                energy_units
            )

            new_branch = new_tech

        created_branch = leap.Branch(name)

        activity = created_branch.Variable("Activity Level")
        activity.DataUnitID = leap.Units.Item(activity_units).ID

        activity.Scale
        print("Requested scale:", repr(scale))
        print("Scale currently stored:", activity.Scale)
        print("Activity unit:", activity.DataUnitText)
        print("Requested:", repr(scale))
        print("Stored scale:", activity.Scale)
        print("Stored unit:", activity.DataUnitText)

        
        print("Branch created: ", name) 

        ExistDemandBranches.add(name)
    else:
        new_branch = f"{"Skipping:  "}{name}"
        print(new_branch)

    return new_branch


# a = create_demand_branch(5, exist_demand_branches)

#%% CREATES ALL DEMAND BRANCHES

def create_all_demand():
    
    BranchIndex = 0 
    while BranchIndex <= len(branch_name):
        print(BranchIndex)
        new_branch = create_demand_branch(BranchIndex, exist_demand_branches)
        BranchIndex = BranchIndex + 1 
    return BranchIndex

BranchIndex = create_all_demand()


# %%
