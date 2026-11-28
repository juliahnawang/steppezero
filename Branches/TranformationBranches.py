#%% LOAD TRANSFORMATION BRANCHES

import pandas as pd

# Put the CSV in the steppezero project folder.
# transformation_branches = pd.read_csv(
#     "transformation_branches.csv", keep_default_na=False, dtype=str
# )

# Alternatively, after importing the CSV into your Google Sheet:
transformation_branches = load_branches(
    url, "transformation_branches_test"
    ).fillna("")

leap.ShowTransformation = True


#%% CREATES A SINGLE TRANSFORMATION BRANCH

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


#%% CREATES ALL TRANSFORMATION BRANCHES

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


#%% RUN TRANSFORMATION BRANCH CREATION

BranchIndex = create_all_transformation()

# %%
