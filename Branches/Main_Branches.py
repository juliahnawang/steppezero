
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
