import pandas as pd
import glob
import os


def list_files_in_folder(folder_path):
    """
    List all files in the specified folder.
    Returns full paths.
    """
    return [os.path.join(folder_path, f) for f in os.listdir(folder_path)
            if os.path.isfile(os.path.join(folder_path, f))]

def loadTelesignalData(file_path):
    """
    Load telesignal data from a single .ods file, apply filtering.
    Returns a pandas DataFrame.
    """
    if not os.path.exists(file_path):
        print(f"File not found: {file_path}")
        return pd.DataFrame()  # return empty DataFrame

    df = pd.read_excel(file_path, engine='odf')
    # Filter rows where 'Operator' contains '/Main' and 'Tag' is not 'NT'
    filtered_df = df[df['Operator'].astype(str).str.contains('/Main', na=False) & (df['Tag'] != 'NT')]
    return filtered_df


