import pandas as pd

REQUIRED_COLUMNS = [
    "Custom Task Name",
    "Worker",
    "Reported Date",
    "Hours",
    "Worker Cost Center"
]

def read_input_excel(file_path):
    """
    Read input Excel file with dynamic header detection.
    Handles the Concentrix export format (Project Actual Hours...).
    """
    raw_df = pd.read_excel(file_path, header=None)
    header_row_index = None

    # Detect header row dynamically (look for "Time Block" and "Hours")
    for i in range(15):
        row_values = raw_df.iloc[i].astype(str).str.lower().tolist()
        if "time block" in row_values and "hours" in row_values:
            header_row_index = i
            break

    if header_row_index is None:
        raise ValueError("Could not detect header row automatically. Expected 'Time Block' and 'Hours' columns.")

    df = pd.read_excel(file_path, header=header_row_index)

    # Normalize column names and remove duplicates
    df.columns = [str(c).strip() for c in df.columns]
    df = df.loc[:, ~df.columns.duplicated()]

    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(f"Missing required columns: {missing}")

    # Clean worker names: "Name (123)" -> "Name"
    df["Worker"] = df["Worker"].astype(str).str.split("(").str[0].str.strip()

    # Parse dates/hours
    df["Reported Date"] = pd.to_datetime(df["Reported Date"], errors="coerce")
    df["Hours"] = pd.to_numeric(df["Hours"], errors="coerce").fillna(0)

    # Remove invalid rows
    df = df.dropna(subset=["Reported Date", "Custom Task Name"])
    df = df[df["Hours"] > 0]

    return df
