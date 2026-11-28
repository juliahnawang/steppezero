
#%%
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

#%%
def load_branches(Url, SheetName):
    
    url = Url.replace("edit?usp=sharing", "gviz/tq?tqx=out:csv&headers=1&sheet=")
    url = f"{url}{SheetName}"
    demand_branches = pd.read_csv(url)

    return demand_branches

#%%
def clean_names(DemandBranches):

    DemandBranches['Branch'] = DemandBranches['Branch'].str.replace(' ', '')
    DemandBranches['Branch'] = DemandBranches['Branch'].str.replace('_','')

    return DemandBranches

#%%
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

#%%

def create_demand_branch(BranchIndex,ExistDemandBranches):

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

#%%
#%%
def create_all_demand():

    BranchIndex = 0

    while BranchIndex < len(branch_name):
        print(BranchIndex)
        new_branch = create_demand_branch(BranchIndex, leap_branch_names)
        BranchIndex = BranchIndex + 1

    return BranchIndex


#%%
def create_transformation_branch(BranchIndex):

    row = transformation_branches.iloc[BranchIndex]
    name = row["Branch"].strip()
    parent, child = name.rsplit("\\", 1)

    print("TESTING:", name)

    if row["Type"] == "Module":

        if not leap.Branches.Exists(name):

            coproduct = row["CoProduct Fuel"]
            coproduct = leap.Fuels.Item(coproduct).ID if coproduct else 0

            leap.AddModule(
                child,
                bool(int(row["IsSimple"])),
                bool(int(row["UseEfficiencies"])),
                bool(int(row["UseCapacities"])),
                bool(int(row["UseLoadCurve"])),
                bool(int(row["UseCoProd"])),
                bool(int(row["UseOutputShares"])),
                bool(int(row["MeetAuxFromOutputs"])),
                coproduct,
                1,
                leap.Units.Item(row["OutputUnit"]).ID,
                1,
                leap.Units.Item(row["CapacityUnit"]).ID
            )

            print("Module created:", name)

        module = leap.Branch(name)
        if bool(int(row["IsSimple"])):
            return module

        # BranchType 6 = Transformation output category.
        output = next(b for b in module.Children if b.BranchType == 6)
        fuel = row["Output Fuel"].strip()

        if not any(b.Name == fuel for b in output.Children):

            leap.AddOutput(
                output.ID,
                fuel,
                int(row["ShortfallImport"]),
                int(row["SurplusExport"]),
                int(row["DomesticPriority"]),
                bool(int(row["IsPriorityFuel"]))
            )

            print("Output fuel added:", fuel)

        return module

    elif row["Type"] == "Process":

        module = leap.Branch(parent)

        processes = next(
            b for b in module.Children if b.BranchType == 5
        )

        fuels = [
            f.strip() for f in row["Feedstock Fuels"].split(",")
            if f.strip()
        ]

        if not fuels:
            raise ValueError(f"No feedstock fuel specified: {name}")

        module_row = transformation_branches[
            transformation_branches["Branch"] == parent
        ].iloc[0]

        is_simple = bool(int(module_row["IsSimple"]))

        process_path = processes.FullName + "\\" + child

        if not leap.Branches.Exists(process_path):

            if is_simple:

                leap.AddSimpleProcess(
                    int(processes.ID),
                    child,
                    leap.Fuels.Item(fuels[0]).ID,
                    leap.Fuels.Item(module_row["Output Fuel"]).ID
                )

            else:

                leap.AddProcess(
                    int(processes.ID),
                    child,
                    leap.Fuels.Item(fuels[0]).ID,
                    int(row["DispatchRule"])
                )

            print("Process created:", name)

        process = leap.Branch(process_path)

        # The first fuel is already assigned during creation.
        if len(fuels) == 1:
            return process

        # Add additional feedstock fuels.
        feedstocks = next(
            (b for b in process.Children if b.BranchType == 32),
            None
        )

        if feedstocks is None:
            raise ValueError(f"No feedstock category found: {name}")

        for fuel in fuels[1:]:

            if not any(b.Name == fuel for b in feedstocks.Children):

                leap.AddFeedstock(
                    int(feedstocks.ID),
                    leap.Fuels.Item(fuel).ID
                )

                print("Feedstock added:", fuel)

        return process

    else:
        raise ValueError(f"Unsupported branch type: {row['Type']}")

#%%
def create_all_transformation():

    # Check fuel names before creating any branches.
    for _, row in transformation_branches.iterrows():

        fuels = (str(row["Output Fuel"]) + "," +
                 str(row["Feedstock Fuels"]) + "," +
                 str(row["CoProduct Fuel"]))

        for fuel in fuels.split(","):
            fuel = fuel.strip()
            if fuel and not leap.Fuels.Exists(fuel):
                raise ValueError(f"Fuel not found in LEAP: {fuel}")

    BranchIndex = 0

    while BranchIndex < len(transformation_branches):

        create_transformation_branch(BranchIndex)

        BranchIndex += 1

    return BranchIndex



#%% CREATES A SINGLE KEY ASSUMPTION BRANCH

def create_key_assumption_branch(BranchIndex):

    row = key_assumptions_branches.iloc[BranchIndex]

    name = row["Branch"].strip()
    parent, child = name.rsplit("\\", 1)

    print("TESTING:", name)

    if leap.Branches.Exists(name):
        print("Skipping:", name)
        return None

    parent_id = leap.Branch(parent).ID

    if row["Type"] == "Category":

        new_branch = leap.AddKeyAssumptionCategory(
            parent_id,
            child
        )

    elif row["Type"] == "Key Assumption":

        new_branch = leap.AddKeyAssumption(
            parent_id,
            child,
            row["Scale"],
            row["Units"]
        )

    else:
        raise ValueError(f"Unsupported branch type: {row['Type']}")

    print("Branch created:", name)

    return new_branch

#%% CREATES ALL KEY ASSUMPTION BRANCHES

def create_all_key_assumptions():

    BranchIndex = 0

    while BranchIndex < len(key_assumptions_branches):

        create_key_assumption_branch(BranchIndex)

        BranchIndex += 1

    return BranchIndex


#%%
# Test interactive window
print("Test Interactive Window")

import win32com.client
from win32com.client import GetActiveObject
import pandas as pd
import os

# area_name = "Transport_Draft1" # for option 2
leap = connect_to_leap()

url = "https://docs.google.com/spreadsheets/d/1dnkUAZv-ZUIfi0NCcSjTSN5pypKsyzDA7xKBEYKQ1ck/edit?usp=sharing"

### Demands ###


demand_branches = load_branches(url, "demand_branches")
# demand_branches = clean_names(demand_branches)

leap_branches = get_all_branches(leap)

leap_branch_names = leap_branches["FullName"]

leap_branch_names = leap_branch_names[
    leap_branch_names.str.contains("Demand", case=False, na=False)
]

leap_branch_names = set(leap_branch_names)

print(demand_branches.to_string(index=False))

branch_type = demand_branches['Type']
branch_name = demand_branches['Branch']
branch_scale = demand_branches['Scale']
branch_activity_units = demand_branches['Activity Units']
branch_fuel = demand_branches['Fuel']
branch_energy_units = demand_branches['Energy Units']

BranchIndex = create_all_demand()

### Transformations ###

transformation_branches = load_branches(
    url, "transformation_branches_test"
    ).fillna("")

leap.ShowTransformation = True

BranchIndex = create_all_transformation()

### Key Assumptions ###

key_assumptions_branches = load_branches(
    url, "key_assumptions_branches_test"
).fillna("")

BranchIndex = create_all_key_assumptions()

# %%

all_branch_names= pd.concat([
    transformation_branches['Branch'],
    key_assumptions_branches['Branch'],
    demand_branches['Branch'],
], ignore_index=True).fillna("")

# %%
