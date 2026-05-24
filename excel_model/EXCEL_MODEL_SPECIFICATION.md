# ============================================================================
# ENERGY SYSTEM OPTIMIZATION MODEL IN MS EXCEL
# Complete Model Documentation and Setup Guide
# ============================================================================

## OVERVIEW

This energy system optimization model is built entirely in Microsoft Excel.
It can be solved using:
- **OpenSolver** (Free, recommended)
- **Frontline Solver** (Commercial, faster)
- **Excel Solver** (Built-in, limited)

## MODEL STRUCTURE

The Excel workbook contains the following sheets:

### SHEET 1: INSTRUCTIONS
- Overview and purpose
- How to use the model
- Solver setup instructions
- Troubleshooting guide

### SHEET 2: INPUT DATA
- Model parameters (years, regions, discount rate)
- Demand data (by type and region)
- Technology parameters (efficiency, cost, lifetime)
- Resource availability (fossil fuels, renewable potential)
- Constraint limits (capacity, emissions)

### SHEET 3: DECISION VARIABLES
- New capacity installations (MW) by technology, region, year
- Technology dispatch (energy output in MWh)
- Storage operations (charge, discharge, level)
- Energy flows between regions

### SHEET 4: CONSTRAINTS
- Energy balance constraints (supply = demand)
- Capacity constraints (min/max installed capacity)
- Transmission constraints (flow limits)
- Emission constraints (total allowable emissions)
- Ramp constraints (change in output per time period)
- Storage constraints (level limits, round-trip efficiency)

### SHEET 5: OBJECTIVE FUNCTION
- Investment costs (capacity × unit cost, annualized)
- Fixed operating costs (capacity × annual O&M)
- Variable operating costs (output × variable cost)
- Emission penalties (if any)
- TOTAL SYSTEM COST to minimize

### SHEET 6: RESULTS
- Optimal capacity by technology, region, year
- Energy dispatch by technology, year
- Annual costs (investment, operating, total)
- Emissions by technology and pollutant
- Summary statistics and key metrics

## SETTING UP THE MODEL

### Step 1: Open Excel and Create Workbook Structure
Create a new Excel workbook with sheets named:
1. Instructions
2. Input_Data
3. Variables
4. Constraints
5. Objective
6. Results

### Step 2: Install OpenSolver
1. Download from: https://opensolver.com
2. Follow installation instructions for your Excel version
3. Restart Excel
4. Verify: Data tab should show "OpenSolver" button

### Step 3: Fill in Input Data
Complete INPUT_DATA sheet with:
- Model horizon (years 1-20)
- Regions (North, Central, South, etc.)
- Technologies and their parameters
- Demand values
- Resource availability
- Constraint limits

### Step 4: Set Up Solver
1. Click Data > OpenSolver > Open and Solve
2. Configure:
   - Objective Cell: Objective!$D$50 (Total Cost)
   - To: Minimize
   - Changing Cells: Variables!$B$5:$G$25 (Capacity decisions)
   - Subject To: Add constraints from Constraints sheet

### Step 5: Run Optimization
1. Click Solve
2. Wait for solver to complete (seconds to minutes)
3. Accept solution
4. View results in Results sheet

## KEY EQUATIONS AND FORMULAS

### Energy Balance Constraint
```
For each demand, region, year:
Sum(Energy Output from all technologies) >= Demand

Formula: =SUMIF(Variables!Technology_Range, Criteria) >= Input_Data!Demand
```

### Capacity Linking
```
Total Capacity = Existing Capacity + New Capacity

Formula: =Input_Data!ExistingCapacity + Variables!NewCapacity
```

### Energy Output Limit
```
Energy Output <= Installed Capacity × Availability × Capacity Factor × Period Duration

Formula: =IF(Variables!Capacity=0, 0, Variables!Capacity * Parameters!Availability * Parameters!CapacityFactor * PeriodDuration)
```

### Investment Cost (Annualized)
```
Annual Investment = New Capacity × Unit Cost × CRF(r, n)

Where CRF = (r(1+r)^n) / ((1+r)^n - 1)

Formula: =Variables!NewCapacity * Input_Data!InvestmentCost * (DiscountRate*(1+DiscountRate)^Life) / ((1+DiscountRate)^Life - 1)
```

### Fixed Operating Cost
```
Annual Fixed Cost = Total Capacity × Fixed O&M Cost

Formula: =Variables!TotalCapacity * Input_Data!FixedOMCost
```

### Emission Calculation
```
Emissions = Energy Output × Emission Factor

Formula: =Variables!EnergyOutput * Input_Data!EmissionFactor
```

### Total System Cost (Objective)
```
Total Cost = Sum(Investment Costs) + Sum(Fixed Operating Costs) + Sum(Variable Costs)

Formula: =SUM(Objective!InvestmentCosts) + SUM(Objective!FixedCosts) + SUM(Objective!VariableCosts)
```

## FILE: ENERGY_SYSTEM_OPTIMIZATION_MODEL.xlsx

This is the complete model ready to use. See individual sheet documentation below.

---

## SHEET: INSTRUCTIONS

Provides step-by-step guidance for using the model.

---

## SHEET: INPUT_DATA

Contains all user inputs organized in sections:

### SECTION A: MODEL PARAMETERS
- Start Year: 2024
- End Year: 2044
- Number of Years: 21
- Time Resolution: Hourly / Seasonal / Annual
- Annual Time Periods: 8760 (for hourly)
- Discount Rate: 5.0%
- Currency: USD

### SECTION B: REGIONS
- Region 1: North
- Region 2: Central  
- Region 3: South
- (Can add more regions as needed)

### SECTION C: TECHNOLOGY PARAMETERS
Table with columns:
- Technology Name
- Type (Thermal, Renewable, Storage, Conversion)
- Existing Capacity (MW)
- Efficiency (%)
- Lifetime (years)
- Investment Cost ($/MW)
- Fixed O&M Cost ($/MW/year)
- Variable Cost ($/MWh)
- Capacity Factor (%)
- Min Load Fraction (%)
- Ramp-up Rate (%/hour)
- Ramp-down Rate (%/hour)
- CO2 Emission Factor (kg CO2/MWh)

Technologies included:
- Coal Plant: 0.35 efficiency, $2500/MW cost, 40 year life
- Natural Gas: 0.50 efficiency, $1800/MW cost, 30 year life
- Solar PV: 1.0 efficiency, $1500/MW cost, 25 year life
- Wind Turbine: 1.0 efficiency, $1800/MW cost, 25 year life
- Hydro Plant: 0.90 efficiency, $3000/MW cost, 50 year life
- Nuclear: 0.33 efficiency, $4000/MW cost, 60 year life
- Battery Storage: 0.85 efficiency, $2000/MW cost, 10 year life
- Heat Pump: 3.0 efficiency, $1200/MW cost, 20 year life

### SECTION D: DEMAND DATA
Table with columns:
- Demand Type
- Region
- Annual Demand (MWh)
- Time Profile (1 = uniform, <1 = peaks, >1 = valley)

Demand types:
- Electricity: 50,000 MWh/year (high seasonal variation)
- Space Heating: 30,000 MWh/year (winter peak)
- Water Heating: 10,000 MWh/year (mild seasonal)
- Transport: 20,000 MWh/year (uniform)
- Industrial: 25,000 MWh/year (uniform)

### SECTION E: RESOURCE AVAILABILITY
Table with columns:
- Resource Type
- Available Quantity
- Cost ($/unit)
- CO2 Emission Factor

Fossil Resources:
- Coal: 1,000,000 tonnes, $50/tonne, 2400 kg CO2/tonne
- Natural Gas: 500,000 MWh equiv, $100/unit, 2200 kg CO2/unit
- Oil: 300,000 MWh equiv, $150/unit, 3150 kg CO2/unit

Renewable Potential:
- Solar: 10,000 MW potential, 25% capacity factor
- Wind: 8,000 MW potential, 35% capacity factor
- Hydro: 5,000 MW potential, 60% capacity factor

### SECTION F: CONSTRAINT LIMITS
Table with columns:
- Constraint Type
- Min Limit
- Max Limit
- Unit

Capacity Constraints:
- Coal: 0-5000 MW
- Natural Gas: 0-4000 MW
- Solar: 0-10000 MW
- Wind: 0-8000 MW
- Hydro: 0-5000 MW
- Nuclear: 0-3000 MW
- Battery: 0-2000 MW

Emission Limits:
- CO2: 3,000,000 tonnes/year
- NOx: 50,000 tonnes/year
- SO2: 100,000 tonnes/year

Other:
- Renewable minimum: 30% of demand
- Nuclear maximum: 50% of capacity
- Demand satisfaction: 100% required

---

## SHEET: VARIABLES (Decision Variables to Optimize)

### Table 1: NEW CAPACITY INSTALLATIONS (MW)
Columns: Technology, Year 1, Year 2, ..., Year 21
Rows: Coal, Natural Gas, Solar, Wind, Hydro, Nuclear, Battery, Heat Pump

Initial values: All 0 (solver will optimize)

### Table 2: TECHNOLOGY DISPATCH (Energy Output in MWh)
Columns: Technology, Year 1, Year 2, ..., Year 21
Rows: Coal, Natural Gas, Solar, Wind, Hydro, Nuclear, Battery, Heat Pump

Initial values: All 0 (calculated by constraints)

### Table 3: STORAGE OPERATIONS (if Battery technology included)
Columns: Storage Tech, Year 1, Year 2, ..., Year 21
Rows:
- Charge (MWh)
- Discharge (MWh)
- Level (MWh)

### Table 4: REGIONAL ENERGY FLOWS (if transmission modeled)
Columns: From Region, To Region, Year 1, Year 2, ..., Year 21
Rows: (links between regions)

---

## SHEET: CONSTRAINTS (Constraint Definitions)

### ENERGY BALANCE CONSTRAINTS
Format: For each demand, region, year
```
Sum(Energy Output) >= Demand
Constraint: Variables!Output >= Input_Data!Demand
```

Implementation:
- Column A: Constraint Name (e.g., "Demand_Electricity_North_Year1")
- Column B: Left Side Formula (=SUM of technology outputs)
- Column C: Operator (">=" for supply >= demand)
- Column D: Right Side Value (=Input_Data!Demand)
- Column E: Status (Formula: =IF(B>=D, "OK", "VIOLATED"))

### CAPACITY CONSTRAINTS  
Format: For each technology, year
```
New Capacity <= Max Capacity
Existing Capacity + New Capacity >= Min Capacity
```

Implementation:
- Column A: Constraint Name
- Column B: Left Side (=Variables!NewCapacity)
- Column C: Operator ("<=")
- Column D: Right Side (=Input_Data!MaxCapacity)

### EMISSION CONSTRAINTS
Format: For each pollutant, year
```
Sum(Emissions from all technologies) <= Emission Limit
```

Implementation:
- Column A: "CO2_Emissions_Year_1" through "Year_21"
- Column B: =SUM(Variables!Output * Input_Data!EmissionFactor)
- Column C: "<="
- Column D: =Input_Data!CO2Limit

### STORAGE CONSTRAINTS
Format: Battery storage operations
```
Charge + Discharge <= Available Energy
Storage Level <= Storage Capacity
Storage Level >= 0
```

Implementation:
- Charge Limit: Charge[t] <= StorageCapacity
- Discharge Limit: Discharge[t] <= StorageCapacity
- Level Balance: Level[t] = Level[t-1] + Charge[t]*Efficiency - Discharge[t]

### MINIMUM RENEWABLE CONSTRAINT
Format:
```
Sum(Renewable Output) >= 30% * Total Demand
```

Implementation:
- Column A: "Renewable_Minimum_Year_1" through "Year_21"
- Column B: =SUM(Solar+Wind+Hydro output)
- Column C: ">="
- Column D: =0.3 * SUM(All Demands)

---

## SHEET: OBJECTIVE (Cost Calculation)

### SECTION 1: INVESTMENT COSTS
Table with columns:
- Technology
- New Capacity (Year 1) [from Variables!]
- New Capacity (Year 2) [from Variables!]
- ...
- New Capacity (Year 21) [from Variables!]
- Unit Cost [from Input_Data!]
- Annualization Factor (CRF) [=Calculated]
- Annual Cost [=Capacity × Unit Cost × CRF]

Annualization Formula:
```
CRF = r(1+r)^n / ((1+r)^n - 1)
where r = discount rate, n = technology lifetime

Formula: =(DiscountRate*(1+DiscountRate)^Lifetime) / ((1+DiscountRate)^Lifetime - 1)
```

Example for Coal (40-year life, 5% discount rate):
```
CRF = 0.05(1.05)^40 / ((1.05)^40 - 1) = 0.0583
Annual Investment Cost = NewCapacity × $2,500,000 × 0.0583
```

### SECTION 2: FIXED OPERATING COSTS
Table with columns:
- Technology
- Total Capacity (Year 1) [=Existing + New]
- Total Capacity (Year 2)
- ...
- Total Capacity (Year 21)
- Fixed O&M Cost ($/MW/year) [from Input_Data!]
- Annual Fixed Cost [=Capacity × Fixed O&M]

Example for Coal:
```
Total Capacity = 2000 MW (existing) + 500 MW (new) = 2500 MW
Fixed O&M Cost = 2500 × $50,000 = $125,000,000/year
```

### SECTION 3: VARIABLE OPERATING COSTS
Table with columns:
- Technology
- Energy Output (Year 1) [from Variables!]
- Energy Output (Year 2)
- ...
- Energy Output (Year 21)
- Variable Cost ($/MWh) [from Input_Data!]
- Annual Variable Cost [=Output × Variable Cost]

Example for Coal:
```
Energy Output = 5,000,000 MWh/year
Variable Cost = 5,000,000 × $25 = $125,000,000/year
```

### SECTION 4: TOTAL ANNUAL COSTS
Table with columns:
- Year
- Total Investment Cost [=SUM(Investment costs for all techs)]
- Total Fixed Cost [=SUM(Fixed O&M for all techs)]
- Total Variable Cost [=SUM(Variable costs for all techs)]
- TOTAL ANNUAL COST [=Investment + Fixed + Variable]

### SECTION 5: DISCOUNTED PRESENT VALUE
Table with columns:
- Year
- Total Annual Cost (undiscounted)
- Discount Factor [=1/(1+DiscountRate)^(Year-StartYear)]
- Discounted Cost [=Annual Cost × Discount Factor]

Example:
```
Year 1: Cost = $500M, DF = 1/(1.05)^0 = 1.000, PV = $500M
Year 2: Cost = $510M, DF = 1/(1.05)^1 = 0.952, PV = $486M
Year 3: Cost = $520M, DF = 1/(1.05)^2 = 0.907, PV = $472M
...
```

### SECTION 6: OBJECTIVE FUNCTION
Cell D$50: =SUM(Discounted Costs for all years)

**This is the cell to MINIMIZE in OpenSolver**

Value: Total system cost in present value terms (USD)

Example result: $2,500,000,000 (2.5 billion)

---

## SHEET: RESULTS (Optimization Results)

### SECTION 1: OPTIMIZATION STATUS
- Date/Time Solved: [Timestamp]
- Solver Used: OpenSolver / Frontline / Excel Solver
- Status: Optimal / Feasible / Infeasible
- Solution Time: [Seconds]
- Objective Value: [Total Cost]

### SECTION 2: OPTIMAL CAPACITY BY TECHNOLOGY
Table with columns:
- Technology
- Existing Capacity (Year 1)
- New Capacity (Year 1)
- Total Capacity (Year 1)
- Total Capacity (Year 5)
- Total Capacity (Year 10)
- Total Capacity (Year 21)

Formulas:
```
Existing: =Input_Data!ExistingCapacity
New: =Variables!NewCapacity
Total: =Existing + New
```

### SECTION 3: ENERGY OUTPUT BY TECHNOLOGY
Table with columns:
- Technology
- Output (Year 1) [MWh]
- Output (Year 5) [MWh]
- Output (Year 10) [MWh]
- Output (Year 21) [MWh]
- % of Total Supply

Formula for % of Total:
```
=Output / SUM(All Outputs)
```

### SECTION 4: ANNUAL COSTS BY CATEGORY
Table with columns:
- Year
- Investment Cost ($M)
- Fixed O&M Cost ($M)
- Variable Cost ($M)
- Total Cost ($M)
- Cumulative Cost ($B)

### SECTION 5: EMISSIONS BY TECHNOLOGY
Table with columns:
- Technology
- CO2 Emissions (Year 1) [tonnes]
- CO2 Emissions (Year 5) [tonnes]
- CO2 Emissions (Year 21) [tonnes]
- Avg Emissions per MWh

### SECTION 6: SYSTEM SUMMARY
Key metrics:
- Total System Cost (PV): $[Value]
- Total Capacity Installed: [MW]
- Total Energy Supply: [MWh]
- Peak Demand: [MW]
- Renewable Energy %: [Percentage]
- Average CO2 Intensity: [kg CO2/MWh]
- System Efficiency: [Percentage]

### SECTION 7: DEMAND SATISFACTION
Table with columns:
- Demand Type
- Total Demand (MWh)
- Total Supplied (MWh)
- Satisfaction %
- Unserved Energy (MWh)

---

## HOW TO USE THE MODEL

### For Beginners:

1. Open the Excel file
2. Review Instructions sheet
3. Check Input_Data sheet and understand structure
4. Open OpenSolver: Data tab > OpenSolver button
5. Click "Open and Solve"
6. Specify:
   - Objective: Objective!$D$50
   - Minimize
   - Changing Cells: Variables!$B$5:$G$25
   - Add constraints from Constraints sheet
7. Click Solve
8. Review Results sheet

### For Advanced Users:

1. Customize Input_Data sheet with your own parameters
2. Add new technologies or demands as needed
3. Modify constraint limits
4. Add additional constraints (e.g., ramp rates, minimum load)
5. Configure solver for faster solving (e.g., Gurobi, CPLEX)
6. Run scenario analysis (e.g., high carbon price, more renewables)
7. Analyze sensitivity (e.g., change discount rate, technology costs)

### Optimization Tips:

1. **Start Simple**: Test with few technologies first
2. **Check Constraints**: Ensure they're not contradictory
3. **Warm Start**: Provide initial feasible solution
4. **Solver Selection**: Use GLPK for free, Gurobi for fast
5. **Scaling**: Use appropriate units ($M, not $) to improve numerics
6. **Gap Tolerance**: Set to 0.01 (1% gap) for faster solving

---

## TROUBLESHOOTING

### Model Won't Solve
- Check constraints aren't contradictory (e.g., min > max)
- Verify capacity is sufficient for demand
- Ensure all formulas reference correct cells
- Try relaxing some constraints

### Results Don't Make Sense
- Check objective function formula
- Verify constraint definitions
- Review input parameters
- Check formulas in Results sheet

### Solving Takes Too Long
- Reduce model size (fewer years, technologies)
- Use commercial solver (Gurobi, CPLEX)
- Increase solver optimality gap
- Simplify constraints

### Solver Not Available
- Download and install OpenSolver
- Alternative: Use Frontline Solver or Excel Solver
- Ensure Excel is recent version

---

## EXTENDING THE MODEL

### Add New Technology:
1. Add row in Input_Data!Technologies table
2. Add column in Variables!NewCapacity table
3. Add constraints in Constraints sheet
4. Add cost calculations in Objective sheet
5. Add results formulas in Results sheet

### Add Time Periods:
1. Change Input_Data!TimeResolution
2. Add columns to Variables sheet
3. Expand Objective function
4. Add time-dependent constraints

### Add Regions:
1. Add region names in Input_Data!Regions
2. Add transmission links in Variables
3. Add flow constraints in Constraints
4. Repeat demand and capacity by region

### Add Storage:
1. Add storage parameters in Input_Data
2. Add charge/discharge/level variables in Variables
3. Add balance equations in Constraints
4. Add cost calculations in Objective

---

## REFERENCES

- OpenSolver: https://opensolver.com
- Energy system optimization: Standard LP/MIP formulation
- Linear Programming: https://en.wikipedia.org/wiki/Linear_programming
- Excel Solver: https://support.microsoft.com/en-us/office

---

This model is ready to use. Generate Energy_System_Optimization_Model.xlsx
by following the sheet templates and instructions above.
