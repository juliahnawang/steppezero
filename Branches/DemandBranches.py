import pandas as pd

def create_demand_branch(BranchIndex, ExistDemandBranches):

    name = branch_name.iloc[BranchIndex]
    parent, child = name.rsplit("\\", 1)

    print("TESTING: ", name)

    scale = branch_scale.iloc[BranchIndex]
    units = branch_units.iloc[BranchIndex]
    fuel = branch_fuel.iloc[BranchIndex]
    energy_units = branch_energy_units.iloc[BranchIndex]
    type = branch_type.iloc[BranchIndex]

    scale = "" if pd.isna(scale) else str(scale).strip()
    units = "" if pd.isna(units) else str(units).strip()
    fuel = "" if pd.isna(fuel) else str(fuel).strip()
    energy_units = "" if pd.isna(energy_units) else str(energy_units).strip()
    type = "" if pd.isna(type) else str(type).strip()

    if parent in ExistDemandBranches and name not in ExistDemandBranches:

        parent_id = leap.Branch(parent).ID

        if type == "Category":

            new_branch = leap.AddCategory(
                parent_id,
                child,
                scale,
                units
            )

        elif type in (
            "Technology",
            "Technology with Energy Intensity (Activity Analysis)"
        ):

            new_branch = leap.AddTechnology(
                parent_id,
                child,
                scale,
                units,
                fuel,
                energy_units
            )

        else:
            print("Unsupported branch type:", type)
            return None

        ExistDemandBranches.add(name)

        print("Branch created:", name)

        return new_branch

    else:
        print("Skipping:", name)
        return None


# %% CREATES ALL DEMAND BRANCHES

def create_all_demand():

    BranchIndex = 0

    while BranchIndex < len(branch_name):

        print(BranchIndex)

        create_demand_branch(
            BranchIndex,
            exist_demand_branches
        )

        BranchIndex += 1

    return BranchIndex


BranchIndex = create_all_demand()