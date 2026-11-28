
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
#%% LOAD AND CREATE KEY ASSUMPTIONS

key_assumptions_branches = load_branches(
    url, "key_assumptions_branches_test"
).fillna("")

#%%
BranchIndex = create_all_key_assumptions()