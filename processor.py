import pandas as pd
from excel_writer import write_task_file

def process_data(df, output_dir):
    """
    Groups by Custom Task Name and generates one Concentrix-style Excel per task.
    Returns list of generated file paths.
    """
    df = df.copy()
    df["Custom Task Name"] = df["Custom Task Name"].astype(str).str.strip()
    df["Worker"] = df["Worker"].astype(str).str.strip()
    df["Reported Date"] = pd.to_datetime(df["Reported Date"], errors="coerce")
    df["Hours"] = pd.to_numeric(df["Hours"], errors="coerce").fillna(0)

    df = df[df["Hours"] > 0]
    df = df.dropna(subset=["Reported Date", "Custom Task Name"])

    generated_files = []
    for task_name, task_df in df.groupby("Custom Task Name", dropna=False):
        task_str = str(task_name).strip()
        if task_str.lower() in ("nan", ""):
            continue

        file_path = write_task_file(task_name, task_df, output_dir)
        generated_files.append(file_path)

    return generated_files
