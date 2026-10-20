#%%

print("Test Interactive Window")

#%% LOAD PACKAGES

import win32com.client
from win32com.client import GetActiveObject
import pandas as pd
import os

#%%





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

area_name = "Transport_Draft1"
leap = connect_to_leap()


#%% LOAD CSV

def load_demands(Url, SheetName):
    
    url = Url.replace("edit?usp=sharing", "gviz/tq?tqx=out:csv&headers=1&sheet=")
    url = f"{url}{SheetName}"
    demand_branches = pd.read_csv(url)

    return demand_branches
    
url = "https://docs.google.com/spreadsheets/d/1u8pPUcBDZaqUYiQd-C2uEHluAH_yhs2KcldS81_jaVM/edit?usp=sharing"
sheet_name = "demand_branches"
demand_branches = load_demands(url, sheet_name)

#%% 

def clean_names(DemandBranches):
    Types = {name: group for name, group in DemandBranches.groupby('Type')}
    category = Types['Category']
    br_name = category.iloc[:,0]
    br_scale = category.iloc[:,3]
    br_units = category.iloc[:,4]










#%%

def get_all_branches(leap):
    branches = []

    def traverse(parent):
        for i in range(1, parent.Children.Count + 1):
            child = parent.Children(i)

            branches.append({
                "Name": child.Name,
                "FullName": child.FullName,
                "ID": child.ID,
                "Depth": child.FullName.count("\\")
            })

            traverse(child)

    # Start from the top-level branches
    for i in range(1, leap.Branches.Count + 1):
        branch = leap.Branches(i)

        branches.append({
            "Name": branch.Name,
            "FullName": branch.FullName,
            "ID": branch.ID,
            "Depth": branch.FullName.count("\\")
        })

        traverse(branch)

    return pd.DataFrame(branches)


# Run
all_branches = get_all_branches(leap)
all_branch_names = all_branches.FullName
# all_branch_names= pd.DataFrame(all_branch_names)

demand_branch_names = all_branch_names[all_branch_names.str.contains('Demand')]
demand_branch_names = pd.DataFrame(demand_branch_names)
exist_demand_branches = set(demand_branch_names["FullName"])

print(all_branches.to_string(index=False))



#%% GET ALL BRANCHES NAMES

def create_demand_branch(BranchIndex):

    name = br_name[BranchIndex]
    parent, child = name.rsplit("\\", 1)
    print("TESTING:     " ,name)

    if br_scale[BranchIndex] == None:
        scale = ""
    else:
        scale = br_scale[BranchIndex]

    if br_units[BranchIndex] == None:
        units = ""
    else: 
        units = br_units[BranchIndex]

    # Test to make sure parent branch exists and branch has not already been created
    # If either condition isn't satisfied, skips to next line
    if parent in exist_demand_branches and name not in exist_demand_branches:

        parent = leap.Branches(parent)

        new_cat = leap.AddCategory(
            parent.ID,
            child,
            scale,
            units
        )

        print("Branch created: ", name) 
    else:
        new_cat = f"{"Skipping:  "}{name}"
        print(new_cat)

    return new_cat


#%%

def create_all_demand():
    
    BranchIndex = 0 
    while BranchIndex <= len(br_name):
        print(BranchIndex)
        new_cat = create_demand_branch(BranchIndex)
        BranchIndex = BranchIndex + 1 
    return BranchIndex

BranchIndex = create_all_demand()






#%%
lrt = leap.Branch(r"Demand\LRT")


LRT = leap.Branch(r"Demand\LRT")
print("LRT ID:", LRT.ID, "| BranchID:", LRT.BranchID)

leap.AddCategory(LRT.BranchID, "TestFolder")
leap.RefreshBranches()

# %%
# Add a folder under Demand
leap.Branch("Demand").Children.Add("Transport")   # check the Add signature and branch-type argument in the help

b = leap.Branch(r"Demand\Transport")
b.Variable("Activity Level").Expression = "7000000000"

#%%
leap.Calculate()
v = leap.Branch(r"Demand\Transport").Variable("Activity Level")
print(v.Value(2030))