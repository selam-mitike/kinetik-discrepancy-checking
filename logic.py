import pandas as pd

INPUT_FILE = "april-may.csv"
OUTPUT_FILE = "april-may_rebilling.xlsx"

SHORT_TRIP_MAX = 2.5
SHORT_TRIP_CHARGE = 10.0
RATE_PER_UNIT = 4.0

df = pd.read_csv(INPUT_FILE, dtype=str)
df.columns = df.columns.str.strip()
print(df.columns.tolist())

def parse_value(cell):
    """ 'A0390:9.7' -> 9.7 """
    if pd.isna(cell) or str(cell).strip() == "":
        return None, None
    text = str(cell).strip()
    code,_,unit = text.rpartition(":")
    try:
        return (code.strip() or None), float(unit)
    except ValueError:
        return (code.strip() or None), None

def expected_charge(units):
    if units is None:
        return None
    if units <= SHORT_TRIP_MAX:
        return SHORT_TRIP_CHARGE
    return units * RATE_PER_UNIT

service_nums = sorted(
    int(m.group(1))
    for c in df.columns
    if (m := re.fullmatch(r"Service (\d+) Units", c))
)

main_rows = []
flagged_rows = []

for _, row in df.iterrows():
    out = {
        "Claim ID": row["Claim ID"],
        "Name": row["Name"],
    }
    row_status = "OK"
    any_service = False

    for n in service_nums:
        code, units = parse_value(row.get(f"Service {n} Units"))
        _, charge = parse_value(row.get(f"Service {n} Charge Amount"))

        if code is None and units is None and charge is None:
            out[f"Service {n} Unit"] = None
            out[f"Service {n} Calculated"] = None
            out[f"Service {n} Charge"] = None
            continue
        
        any_service = True
        out[f"Service {n} Unit"] = units
        out[f"Service {n} Charge"] = charge

        if units is None or charge is None:
            out[f"Service {n} Calculated"] = None
            if row_status == "OK":
                row_status = "REVIEW (could not parse)"
            continue

        calc = expected_charge(units)
        out[f"Service {n} Calculated"] = calc
        if calc != charge:
            row_status = "NEED TO REBILL"
            flagged_rows.append({
                "Claim ID": row["Claim ID"],
                "Name": row["Name"],
                "Service #": n,
                "Unit": units,
                "Calculated": calc,
                "Charge": charge,
                "Status": row_status
            })
        if not any_service:
            row_status = "NO SERVICES"
        out["Status"] = row_status
        main_rows.append(out)
        
"""
Read from the csv file
    look at "Service 1 Units" and get the unit/milage: it's in form of A0390:9.7 where 9.7 is the unit
    multiply the unit by 4 and compare with the charge in the "Service 1 Charge Amount" column. The charge is in the form A0390:35.4 where 35.4 is the charge.
    if the calculated doesn't match the charge then flag it for review and add a column in the output file the calculated value(mulitpled by 4)
    also add a column for the charge value from the "Service 1 Charge Amount" column for comparison

    for any milage that's 2.5 or less, the charge should be 10$
do this for all services in each row (one row may have multiple services and others less so there would be empty columns for the missing services)

output file columns to have: "Claim ID", "Name", "Service 1 unit", "Service 1 Calculated", "Service 1 Charge", "Service 2 unit", "Service 2 Calculated", "Service 2 Charge", etc. , "Status"
use "Claim ID" and "Name" columns to create the rows. Sheet 1
Sheet 2: Only services flagged NEED TO REBILL in Status column appear here.
"""