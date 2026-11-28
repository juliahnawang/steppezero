
#%% LOAD SCENARIO METADATA

metadata_scenarios = load_branches(
    url, "metadata_scenarios"
).fillna("")

scenario_years = dict(zip(
    metadata_scenarios["Scenario"].str.strip(),
    metadata_scenarios["Year"]
))

base_year = int(scenario_years["Current Accounts"])

#%% SET LEAP MODEL YEARS

leap.BaseYear = base_year
leap.FirstScenarioYear = 2026

print("Base Year:", leap.BaseYear)
print("First Scenario Year:", leap.FirstScenarioYear)


#%% LOAD ALL DATA

key_assumptions_data = load_branches(url, "key_assumptions_data")
electricity_data = load_branches(url, "electricity_data_test")
heating_data = load_branches(url, "heating_data_test")
transportation_data = load_branches(url, "transport_data_test")

print("Key Assumptions rows:", len(key_assumptions_data))
print("Electricity rows:", len(electricity_data))
print("Heating rows:", len(heating_data))
print("Transportation rows:", len(transportation_data))

print("\nTransportation scenarios:")
# print(transportation_data['Scenario'].value_counts())

data = pd.concat([
    key_assumptions_data,
    electricity_data,
    heating_data,
    transportation_data
], ignore_index=True).fillna("")

#%% PREPARE DATA

def prepare_data(data, scenario="Current Accounts"):

    data = data.copy().fillna("")

    # Clean column values.
    for col in ["Branch", "Variable", "Scenario",
                "Data status", "Units", "Value"]:

        data[col] = data[col].astype(str).str.strip()

    # Select only the requested scenario and valid inputs.
    data = data[
        (data["Scenario"] == scenario) &
        (data["Data status"] != "Not a LEAP input") &
        (data["Variable"] != "Reference") &
        (data["Value"] != "")
    ].copy()

    # Preserve original year for data prioritization.
    data["Original Year"] = pd.to_numeric(
        data["Year"], errors="coerce"
    )

    # Assign model year from scenario metadata.
    if scenario not in scenario_years or scenario_years[scenario] == "":
        raise ValueError(f"No year defined for scenario: {scenario}")

    data["Year"] = int(scenario_years[scenario])

    # Standardize Key Assumption paths.
    data["Branch"] = data["Branch"].str.replace(
        "Key Assumption\\", "Key\\", regex=False
    )

    data["Branch"] = data["Branch"].str.replace(
        "Key Assumptions\\", "Key\\", regex=False
    )

    # Correct branch names modified by LEAP.
    data["Branch"] = data["Branch"].str.replace(
        "Vehicle Stock", "Vehicle_Stock", regex=False
    )

    data["Branch"] = data["Branch"].str.replace(
        "Residential Floor Area", "Residential_Floor Area", regex=False
    )

    data["Branch"] = data["Branch"].str.replace(
        "Total Heated Floor Area", "Total Heated_Floor Area", regex=False
    )

    # Correct flat transportation paths to nested paths.
    for vehicle in [
        "Passenger Vehicles",
        "Cargo Vehicles",
        "Buses",
        "Motorcycles"
    ]:

        for category in ["Fuel Economy", "Electrical Intensity"]:

            data["Branch"] = data["Branch"].str.replace(
                f"Key\\{category}\\{vehicle} ",
                f"Key\\{category}\\{vehicle}\\",
                regex=False
            )

    # Key Assumptions use Activity Level.
    data.loc[
        data["Branch"].str.startswith("Key\\"),
        "Variable"
    ] = "Activity Level"

    # Convert formatted values to numbers.
    data["Value"] = pd.to_numeric(
        data["Value"]
        .str.replace(",", "", regex=False)
        .str.replace("%", "", regex=False),
        errors="coerce"
    )

    # Report invalid numerical values.
    invalid_values = data[data["Value"].isna()]

    if not invalid_values.empty:
        print("\nInvalid numerical values:")
        print(invalid_values[["Branch", "Variable"]])

    data = data.dropna(subset=["Value"])

    # Primary data quality priority.
    data["Priority"] = data["Data status"].map({
        "Model setting": 0,
        "Local data": 1,
        "Check": 2
    }).fillna(3)

    # Secondary data quality priority.
    data["Estimate Priority"] = data["Data status"].map({
        "Proxy": 0,
        "Placeholder": 1
    }).fillna(2)

    # Prioritize quality, most recent year, then estimate quality.
    data = data.sort_values(
        ["Priority", "Original Year", "Estimate Priority"],
        ascending=[True, False, True],
        na_position="last",
        kind="stable"
    )

    # Retain the best input for each branch and variable.
    data = data.drop_duplicates(
        subset=["Branch", "Variable", "Scenario", "Units"],
        keep="first"
    )

    return data.reset_index(drop=True)


#%% SET A SINGLE BRANCH EXPRESSION

def set_branch_expression(branch, variable, scenario, expression, units):

    print("TESTING:", branch, variable, scenario)

    try:

        # Check that branch exists.
        if not leap.Branches.Exists(branch):
            print("SKIPPED: Branch not found:", branch)
            return False

        # Check that variable exists.
        var = leap.Branch(branch).Variable(variable)

        if var is None:
            print("SKIPPED: Variable not found:", variable)
            return False

        if not var.IsData:
            print("SKIPPED: Not an input variable:", variable)
            return False

        # Display units for verification.
        print("Spreadsheet units:", units)
        print("LEAP units:", var.DataUnitText)

        # Avoid automatically assigning unverified fuel economy units.
        if variable == "Fuel Economy":
            print("SKIPPED: Fuel economy units need verification")
            return False

        var.Expression = expression

        print("Expression assigned:", var.Expression)

        return True

    except Exception as e:

        print("ERROR:", branch, variable)
        print(e)

        return False


#%% SET ALL EXPRESSIONS

def set_all_expressions(data, scenario="Current Accounts"):

    data = prepare_data(data, scenario)

    # Check model years.
    if scenario == "Current Accounts":

        invalid_years = data[
            (data["Year"] < leap.BaseYear) |
            (data["Year"] >= leap.FirstScenarioYear)
        ]

        if not invalid_years.empty:
            raise ValueError("Data years outside Current Accounts")

    leap.ActiveScenario = scenario

    print("\nScenario:", scenario)
    print("Total inputs to import:", len(data))

    successful = 0
    unsuccessful = 0

    # Import Key Assumptions first.
    data["Branch Priority"] = (
        ~data["Branch"].str.startswith("Key\\")
    )

    data = data.sort_values(
        "Branch Priority", kind="stable"
    )

    for _, row in data.iterrows():

        branch = row["Branch"]
        variable = row["Variable"]
        expression = str(row["Value"])
        units = row["Units"]

        # Fuel-specific inputs require separate handling.
        if variable in ["Feedstock Fuel Share", "Emission Factor"]:

            print("SKIPPED: Fuel-specific input:", branch, variable)

            unsuccessful += 1
            continue

        result = set_branch_expression(
            branch,
            variable,
            scenario,
            expression,
            units
        )

        if result:
            successful += 1
        else:
            unsuccessful += 1

    print("\nIMPORT COMPLETE")
    print("Successful:", successful)
    print("Unsuccessful or skipped:", unsuccessful)

    return successful


#%% TEST PREPARED DATA

prepared_data = prepare_data(data, "Current Accounts")

print("\nTotal prepared rows:", len(prepared_data))

print("Key Assumptions:",
      prepared_data["Branch"].str.startswith("Key\\").sum())

print("Demand:",
      prepared_data["Branch"].str.startswith("Demand\\").sum())

print("Transformation:",
      prepared_data["Branch"].str.startswith("Transformation\\").sum())

print(prepared_data[[
    "Branch",
    "Variable",
    "Units",
    "Data status",
    "Original Year",
    "Year",
    "Value"
]].to_string(index=False))


#%% IMPORT CURRENT ACCOUNTS

set_all_expressions(data, "Current Accounts")