from IPython.display import display
import matplotlib.pyplot as plt
import numpy as np
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter
import os
import pandas as pd
import re
import seaborn as sns


def add_fuel_dfs_missing_columns(dfs, period):
    """Add missing columns to the DataFrames.
    Args:
        dfs: List of DataFrames
    Returns:
        List of DataFrames with missing columns added
    """
    print("Adding missing columns to the DataFrames...\n")
    
    # Validate input is a list
    is_list = True
    if not isinstance(dfs, list):
        is_list = False
        dfs = [dfs]
    
    # Add missing columns
    for df in dfs:
        if 'PurchaseMonth' not in df.columns:
            df['PurchaseMonth'] = period
            print(f"Added PurchaseMonth column to {df.supplier} with value {period}")
        else:
            print(f"PurchaseMonth column already exists in {df.supplier}")
        
        if 'Buyer' not in df.columns:
            df['Buyer'] = df.supplier
            print(f"Added Buyer column to {df.supplier} with value {df.supplier}")
        else:
            print(f"Buyer column already exists in {df.supplier}")
        print("-"*50 + "\n")
    
    if is_list:
        return dfs
    else:
        return dfs[0]


def merge_disclosed_fuel_data(dfs, required_columns):
    """Merge multiple DataFrames after validating structure and column requirements.
    Ensures:
    - Input is a list of DataFrames
    - All DataFrames contain specified required_columns
    - Columns are reordered with required_columns first
    - Data is concatenated into a single DataFrame
    Args:
        dfs: List of pandas DataFrames to merge
        required_columns: List of column names that must exist in all DataFrames
    Returns:
        pd.DataFrame: Merged result with consistent column order
    Raises:
        ValueError: If input validation fails (not a list or missing columns)
    """
    print("Beginning disclosed fuel data merge...")

    if not isinstance(dfs, list):
        raise ValueError("dfs is not a list. Pass a list of DataFrames.")
    
    ## Check for EXISTENCE of required columns
    missing_columns = {}
    for df_index, df in enumerate(dfs):
        mis_cols = set(required_columns) - set(df.columns)
        if mis_cols:
            missing_columns[df_index] = mis_cols
    
    if missing_columns:
        for df_index, missing_cols in missing_columns.items():
            print(f"DataFrame at index {df_index} is missing columns: {missing_cols}")
        raise ValueError("Some DataFrames are missing required columns")
    else:
        print("All DataFrames have the same columns")
    
    return pd.concat(dfs, ignore_index=True)


def rename_existing_columns(dfs, columns_to_rename):
    """Rename columns of a DataFrame based on a mapping dictionary.
    For each column in the DataFrame, if the column name is not a key in
    `columns_to_rename`, it checks if the column name exists in the values
    (which are lists of strings) of the dictionary. If a match is found,
    the column is renamed to the corresponding dictionary key. Columns that
    are already keys in `columns_to_rename` are kept as is.
    Args:
        dfs: Input DataFrame or list of DataFrames.
        columns_to_rename: A dictionary where keys are the desired new column names
                           and values are lists of possible old column names.
                           Example: {'new_name': ['old_name1', 'old_name2']}.
    Returns:
        DataFrame or list of DataFrames with renamed columns.
    """
    print("Renaming existing columns...")
    # Check if dfs is a list
    is_list = True
    if not isinstance(dfs, list):
        is_list = False
        dfs = [dfs]
    
    # Rename columns
    for i, df in enumerate(dfs):
        name = df.name
        rename_map = {}
        for col in df.columns:
            if col not in columns_to_rename.keys():
                for new_name, old_names in columns_to_rename.items():
                    if col in old_names:
                        rename_map[col] = new_name
                        print(f"Renamed {col} to: {new_name}")
                        break  # Found a match, move to the next column
        if rename_map:
            dfs[i] = df.rename(columns=rename_map)
            dfs[i].name = name

    if is_list:
        return dfs
    else:
        return dfs[0]


def order_existing_columns(dfs, cols_to_rename):
    """Reorder DataFrame columns to put required columns first, keeping others after.
    
    Args:
        dfs: Input DataFrame or list of DataFrames to process.
        cols_to_rename: Dictionary where keys are column names that should appear first.
                         Only the keys are used; values are ignored.
    Returns:
        Processed DataFrame(s) with reordered columns. Returns same type as input
        (single DataFrame or list of DataFrames).
    """
    print("\nOrdering existing columns...")
    # Check if dfs is a list
    is_list = True
    if not isinstance(dfs, list):
        is_list = False
        dfs = [dfs]

    # Order columns
    for i, df in enumerate(dfs):
        name = df.name
        req_cols = [col for col in cols_to_rename.keys() if col in df.columns]
        dfs[i] = df[req_cols + [col for col in df.columns if col not in req_cols]]
        dfs[i].name = name
    
    if is_list:
        return dfs
    else:
        return dfs[0]


def fill_na_fuel_columns(dfs, columns_to_fill):
    """Replace NaN values with empty strings in specified DataFrame columns.
    Args:
        df: Input DataFrame
        columns_to_fill: List of column names to process
    Returns:
        DataFrame with processed columns (unchanged if no valid columns found)
    """
    print("\nFilling NaN values to certain columns of the DataFrames...")
    
    # Check if dfs is a list
    is_list = True
    if not isinstance(dfs, list):
        is_list = False
        dfs = [dfs]
    
    # Fill the Na values of the columns in the dataframe
    for i, df in enumerate(dfs):
        name = df.name
        col_in_df = [col for col in columns_to_fill if col in df.columns]
        # Fill the Na values of the columns in the dataframe
        if col_in_df:
            dfs[i] = df.assign(**{col: df[col].fillna('') for col in col_in_df})
        dfs[i].name = name
    
    if is_list:
        return dfs
    else:
        return dfs[0]


def convert_date_to_string_format(dfs, date_column, date_format):
    """Convert a date column to a string format.
    Args:
        df: Input DataFrame
        date_column: Column name to convert
        date_format: Format to convert to
    """
    print("\nConverting date column to string format...")
    # Check if dfs is a list
    is_list = True
    if not isinstance(dfs, list):
        is_list = False
        dfs = [dfs]
    
    # Convert the date column to the desired format
    for i, df in enumerate(dfs):
        dfs[i][date_column] = pd.to_datetime(df[date_column], format=date_format).dt.strftime(date_format)
    
    if is_list:
        return dfs
    else:
        return dfs[0]


def set_content_to_uppercase(dfs, columns_to_uppercase):
    """Convert specified columns to uppercase.
    
    Processes either a single DataFrame or list of DataFrames. For each valid column found,
    converts all string values to uppercase. Non-string columns will remain unchanged.
    
    Args:
        dfs: Input DataFrame or list of DataFrames to process
        columns_to_uppercase: List of column names to convert to uppercase
        
    Returns:
        Union[pd.DataFrame, List[pd.DataFrame]]: 
            Processed DataFrame(s) with specified columns converted to uppercase.
            Returns same type as input (list if input was list, single DataFrame otherwise).
            Unchanged if no valid columns found.
    """
    print("\nSetting content to uppercase...")
    # Check if dfs is a list
    is_list = True
    if not isinstance(dfs, list):
        is_list = False
        dfs = [dfs]
    
    # Set the content of the specified columns to uppercase
    for i, df in enumerate(dfs):
        name = df.name
        col_in_df = [col for col in columns_to_uppercase if col in df.columns]
        if col_in_df:
            dfs[i] = df.assign(**{col: df[col].str.upper() for col in col_in_df})
        dfs[i].name = name
    
    if is_list:
        return dfs
    else:
        return dfs[0]


def remove_eds_from_flag(dfs, flag_column):
    """Remove EDS from flag column.
    Args:
        dfs: Input DataFrame or list of DataFrames
        flag_column: Column name to remove EDS from
    """
    print("\nRemoving EDS from flag column...")
    # Check if dfs is a list
    is_list = True
    if not isinstance(dfs, list):
        is_list = False
        dfs = [dfs]
    
    # Remove EDS from flag column
    for i, df in enumerate(dfs):
        dfs[i][flag_column] = df[flag_column].str.replace(r'\bEDS\b', '', regex=True)
    
    if is_list:
        return dfs
    else:
        return dfs[0]


def remove_spaces_n_strip(dfs, columns_to_strip):
    """Remove multiple spaces and strip spaces from specified columns.
    Args:
        dfs: Input DataFrame or list of DataFrames
        columns_to_strip: List of column names to process
    Returns:
        Union[pd.DataFrame, List[pd.DataFrame]]: 
            Processed DataFrame(s) with specified columns cleaned.
            Returns same type as input (list if input was list, single DataFrame otherwise).
            Unchanged if no valid columns found.
    """
    print("\nRemoving multiple spaces and stripping spaces from specified columns...")
    # Check if dfs is a list
    is_list = True
    if not isinstance(dfs, list):
        is_list = False
        dfs = [dfs]
    
    # Remove multiple spaces and strip spaces from specified columns
    for i, df in enumerate(dfs):
        name = df.name
        col_in_df = [col for col in columns_to_strip if col in df.columns]
        if col_in_df:
            dfs[i] = df.assign(**{col: df[col].str.replace(r'\s+', ' ', regex=True).str.strip() for col in col_in_df})
        dfs[i].name = name
    
    if is_list:
        return dfs
    else:
        return dfs[0]


def get_fuel_unique_buyers(dfs):
    """Get unique buyers from a list of DataFrames.
    Args:
        dfs: Input DataFrame or list of DataFrames
    Returns:
        List of unique buyers
    """
    print("\nGetting unique buyers from the DataFrames...")
    # Check if dfs is a list
    if not isinstance(dfs, list):
        dfs = [dfs]
    
    # Get unique buyers
    unique_buyers = []
    for df in dfs:
        unique_buyers.extend(df['Buyer'].unique().tolist())
    unique_buyers = sorted(set(unique_buyers))
    
    print(f"Number of unique buyers: {len(unique_buyers)}")
    print(f"Unique buyers: {unique_buyers}")
    
    return unique_buyers


def get_unique_flags(dfs):
    """
    Extract unique flag values from one or more DataFrames.
    Args:
        dfs: DataFrame or list of DataFrames, each containing a 'Flag' column
    Returns: 
        List of sorted unique non-empty flag strings
    """
    # Check dfs is a list
    if not isinstance(dfs, list):
        dfs = [dfs]
    
    flags = []
    for df in dfs:
        flags.extend(df['Flag'].unique().tolist())
    flags = sorted(set(flags))
    flags = [x for x in flags if x != '']  # Exclude empty strings
    
    print(f"Unique flags len ({len(flags)}): {flags}")
    return flags


def lematize_serv_st(df, serv_station_lemat_colname, unique_flags):
    """
    Lemmatizes service station names by running several transformations.
    Args:
        df: DataFrame containing service station data
        serv_station_lemat_colname: Column name to lemmatize
        unique_flags: List of unique flag names to remove
    """
    # Set values in uppercase
    df[serv_station_lemat_colname] = df[serv_station_lemat_colname].str.upper()
    
    # Make a list of terms to replace (some will be removed later)
    terms_to_replace = {
        'DUITAMA-OT': 'DUITAMA',
        'DUITAMAOT': 'DUITAMA',
        'TEPEL': 'TERPEL',
        'TRINFO': 'TRIUNFO',
        'LA MONTAÑA GIRARDOTA': 'LA MONTAÑA ITAGUI',
    }
    for old_term, new_term in terms_to_replace.items(): # Replace the words in the list
        df[serv_station_lemat_colname] = df[serv_station_lemat_colname].str.replace(old_term, new_term, regex=True)
    
    # Remove some specific words using regex pattern
    words_to_remove = r'\b(EDS|OT|AEROPUERTO|SODEXO|COESCO|EL|LA|LOS|LAS|ALIADA)\b'  # \b is a word boundary (start/end of a word)
    # Replace with space to avoid errors in the following steps. We will remove the spaces later
    df[serv_station_lemat_colname] = df[serv_station_lemat_colname].str.replace(words_to_remove, ' ', regex=True)
    
    # Remove unique_flags from ServStLematized when included in the ServiceStation (ServStLematized) except if flag is 'ICEBERG'
    for flag in unique_flags:
        mask = df[serv_station_lemat_colname] != 'ICEBERG'
        # Replace with space to avoid errors in the following steps. We will remove the spaces later
        df.loc[mask, serv_station_lemat_colname] = df.loc[mask, serv_station_lemat_colname].str.replace(flag, ' ')

    # Remove content to the right of the first parenthesis
    # Replace with space to avoid errors in the following steps. We will remove the spaces later
    df[serv_station_lemat_colname] = df[serv_station_lemat_colname].str.replace(r'\(.*', ' ', regex=True)
    
    # Apply a simple phonetical lematizer
    df[serv_station_lemat_colname] = df[serv_station_lemat_colname] \
        .str.replace('C', 'S', regex=False) \
        .str.replace('CH', 'S', regex=False) \
        .str.replace('K', 'S', regex=False) \
        .str.replace('Z', 'S', regex=False) \
        .str.replace('H', '', regex=False) \
        .str.replace('G', 'J', regex=False) \
        .str.replace('RR', 'R', regex=False) \
        .str.replace('Ñ', 'N', regex=False) \
        .str.replace('LL', 'Y', regex=False) \
        .str.replace('QUE', 'KE', regex=False) \
        .str.replace('Q', 'K', regex=False) \
        .str.replace('V', 'B', regex=False) \
        .str.replace('W', 'V', regex=False) \
        .str.replace('Á', 'A', regex=False) \
        .str.replace('É', 'E', regex=False) \
        .str.replace('Í', 'I', regex=False) \
        .str.replace('Ó', 'O', regex=False) \
        .str.replace('Ú', 'U', regex=False) \
        .str.replace('Ü', 'U', regex=False)
    
    # Only pass letters and numbers. REMOVE other characters AND SPACES
    df[serv_station_lemat_colname] = df[serv_station_lemat_colname].astype(str).str.replace(r'[^\w]', '', regex=True)


def lematize_discl_serv_st(dfs, unique_flags):
    """
    Lemmatizes service station names by running several transformations on multiple DataFrames.
    Args:
        dfs: List of DataFrames containing service station data
        unique_flags: List of unique flag names to remove
    Returns:
        None (modifies DataFrames in place)
    """
    print("\nLemmatizing service station names...")
    # Check if dfs is a list
    if not isinstance(dfs, list):
        dfs = [dfs]
    
    # Lemmatize service station names
    for df in dfs:
        # Copy ServiceStation column data to ServStLematized column
        df['ServStLematized'] = df['ServiceStation']
        # Run the lematization function on the ServStLematized column
        lematize_serv_st(df, 'ServStLematized', unique_flags)

        display(df.head(3))
        print(f"{len(df.ServiceStation.unique())} ServiceStations")
        print(f"{len(df.ServStLematized.unique())} ServStLematized")


def check_serv_st_lemat_to_serv_st(dfs):
    """Find cases where multiple ServiceStations map to same lematized value.
    Identifies ServStLematized values that correspond to multiple different
    ServiceStation names, indicating potential lematization collisions.
    Args:
        dfs: DataFrame or list of DataFrames with 'ServStLematized' and 'ServiceStation' columns
    """
    print("\nChecking the ServStLematized with more than one ServiceStation...")
    # Check if dfs is a list
    if not isinstance(dfs, list):
        dfs = [dfs]
    
    # Check the ServStLematized with more than one ServiceStation
    for df in dfs:
        ## Group by ServStLematized and get unique ServiceStation matches
        srv_st_matches = df.groupby('ServStLematized')['ServiceStation'].unique()
        ## Filter for ServiceStations with more than one match
        multi_lematized = srv_st_matches[srv_st_matches.apply(len) > 1]
        print(f"\nServiceStation matches by ServStLematized (more than one):\n{multi_lematized}")
        print("-"*50 + "\n")


def fill_flag_by_frequency_on_serv_st_lematized(dfs):
    """Fills empty 'Flag' values with the most frequent non-blank flag 
    for the corresponding 'ServStLematized' across all DataFrames.
    
    Args:
        dfs: List of DataFrames containing 'Flag' and 'ServStLematized' columns. 
             Can also be a single DataFrame.
        
    Returns:
        list or pd.DataFrame: The modified list of DataFrames (or single DataFrame) 
                              with the 'Flag' column filled.
    """
    print("\nFilling empty Flags based on ServStLematized frequency...")
    # Check if dfs is a list
    is_list = True
    if not isinstance(dfs, list):
        is_list = False
        dfs = [dfs]

    all_pairs = []
    # Consolidate all 'Flag' and 'ServStLematized' pairs from all dataframes.
    for df in dfs:
        df['Flag'] = df['Flag'].replace(r'^\s*$', np.nan, regex=True)
        all_pairs.extend(df[['Flag', 'ServStLematized']].values.tolist())
    
    if all_pairs:
        pairs_df = pd.DataFrame(all_pairs, columns=['Flag', 'ServStLematized'])
        # Drop rows with no 'Flag' to build a clean mapping.
        valid_pairs_df = pairs_df.dropna(subset=['Flag'])
        
        if not valid_pairs_df.empty:
            # Create a map from each service station to its most frequent flag (brand).
            serv_st_to_flag_map = valid_pairs_df.groupby('ServStLematized')['Flag'].agg(lambda x: x.mode()[0] if not x.mode().empty else np.nan).to_dict()
            
            # Use the map to fill missing 'Flag' values in each dataframe.
            for df in dfs:
                df['Flag'] = df['Flag'].fillna(df['ServStLematized'].map(serv_st_to_flag_map))
                # Convert any remaining NaNs to empty strings to avoid downstream errors.
                df['Flag'] = df['Flag'].fillna('')

    if is_list:
        return dfs
    else:
        return dfs[0]


def create_flag_serv_st_lematized(dfs):
    """Creates a combined identifier column by concatenating flag and lemmatized service station names.
    Generates a new 'FlagServStLematized' column by combining values from 'Flag' and 'ServStLematized' columns.
    Handles both single DataFrame and list of DataFrames inputs. Modifies DataFrames in place.
    Args:
        dfs (pd.DataFrame or list[pd.DataFrame]): Input DataFrame(s) containing flag and service station data.
            Must include 'Flag' and 'ServStLematized' columns.
    Returns:
        None: Modifies input DataFrame(s) by adding 'FlagServStLematized' column.
    """
    print("\nCreating FlagServStLematized column...")
    
    # Check if dfs is a list
    if not isinstance(dfs, list):
        dfs = [dfs]
    
    for df in dfs:
        # Create the combined column
        df['FlagServStLematized'] = df['Flag'] + df['ServStLematized']
        
        # Print sample and count of unique values
        display(df[['Flag', 'ServStLematized', 'FlagServStLematized']].head(3))
        print(f"Number of unique FlagServStLematized: {len(df['FlagServStLematized'].unique())}")
    
    print("-"*50 + "\n")
    
    # Extract the unique FlagServStLematized values
    unique_flag_serv_st_lematized = []
    for df in dfs:
        unique_flag_serv_st_lematized.extend(df['FlagServStLematized'].unique().tolist())
    unique_flag_serv_st_lematized = sorted(set(unique_flag_serv_st_lematized))
    print(f"Sample of Total Unique FlagServStLematized: {unique_flag_serv_st_lematized[:5]}")
    print(f"Number of Total Unique FlagServStLematized: {len(unique_flag_serv_st_lematized)}")
    
    return unique_flag_serv_st_lematized


def clean_fuel_data(dfs):
    """Clean fuel data by removing empty/invalid rows and resetting index.
    Handles both single DataFrame and list of DataFrames inputs.
    Args:
        dfs: DataFrame or list of DataFrames to clean
    Returns:
        Cleaned DataFrame(s) with reset index
    """
    print("\nCleaning fuel data...")
    # Check if dfs is a list
    is_list = True
    if not isinstance(dfs, list):
        is_list = False
        dfs = [dfs]
    
    cleaned_dfs = []
    for df in dfs:
        name = df.name
        print(f"\nDataframe name: {name}")
        print(f"Input shape before cleaning: {df.shape}")
        
        # Delete empty rows
        df = df.dropna(how='all')
        print(f"After dropping empty rows: {df.shape}")
        
        # Delete rows where the ServiceStation is NaN
        df = df[df['ServiceStation'].notna()]
        print(f"After dropping NaN ServiceStation: {df.shape}")
        
        # Delete rows with NaN (Empty or 'NAME?' text) in the Gls or PricePerGl column
        for col in ['Gls', 'PricePerGl']:
            df = df[df[col].notna()]
        print(f"After dropping rows where Gls or PricePerGl are NaN: {df.shape}")
        
        # Delete rows with non-positive Gls
        df = df[df['Gls'] > 0]
        print(f"After dropping rows where Gls is not greater than 0: {df.shape}")
        
        # Delete rows with non-positive PricePerGl
        df = df[df['PricePerGl'] > 0]
        print(f"After dropping rows where PricePerGl is not greater than 0: {df.shape}")
        
        # Reset index
        df = df.reset_index(drop=True)
        df.name = name
        
        # Add the cleaned dataframe to the list
        cleaned_dfs.append(df)
    
    if is_list:
        return cleaned_dfs
    else:
        return cleaned_dfs[0]


def preprocess_department_column(df):
    """Standardizes and validates department names in a DataFrame.
    Performs the following transformations:
    1. Normalizes spacing and removes accents from department names
    2. Maps common variations to standardized department names
    3. Validates against a set of allowed department names
    4. Creates a processed column with cleaned department names
    Args:
        df (pd.DataFrame): Input DataFrame containing a 'Department' column
    Returns:
        pd.DataFrame: Original DataFrame with added 'DepartmentProcessed' column
    Raises:
        ValueError: If any department names remain unhandled after processing
    """
    # Check if the Department column exists in the dataframe
    if not 'Department' in df.columns:
        print("No Department column found in the dataframe. Nothing to lematize.")
        print("-"*50 + "\n")
        return df
    
    # Preprocess the Department column
    # 1. Replace multiple spaces with a single space and strip spaces from Department
    df['Department'] = df['Department'].str.replace('  ', ' ').str.strip()
    # Remove accents from Department
    df['Department'] = df['Department'].str.normalize('NFKD').str.encode('ascii', errors='ignore').str.decode('utf-8')

    # Get the unique departments
    unique_departments = sorted(df['Department'].astype(str).unique().tolist())
    print(f"Unique departments: {unique_departments}")
    print(f"Number of unique departments: {len(unique_departments)}")

    # 2. Create a conversion dictionary for the Department column
    department_conversion_dict = {
        'BOGOTA, D.C.': 'BOGOTA',
        'C/MARCA': 'CUNDINAMARCA',
        'CASANARES': 'CASANARE',
        'N. SANTANDER': 'NORTE DE SANTANDER',
        'NORTE DEL SANTANDER': 'NORTE DE SANTANDER',
        'RISARALDAS': 'RISARALDA',
        'VALLE': 'VALLE DEL CAUCA',
        'MALAMBO': 'ATLANTICO', # Malambo is a municipality in Atlantico
        'SAHAGUN': 'CORDOBA' # Sahagun is a municipality in Cordoba
    }

    # 3. Apply the conversion and clean the Department column
    df['DepartmentProcessed'] = df['Department'].replace(department_conversion_dict)

    # Handle NaN values by filling them with empty strings
    df['DepartmentProcessed'] = df['DepartmentProcessed'].fillna('')

    # 4. Define the set of all valid final departments
    valid_departments = {
        '', 'ANTIOQUIA', 'ATLANTICO', 'BOGOTA', 'BOLIVAR', 'BOYACA', 'CALDAS', 'CASANARE', 
        'CAUCA', 'CESAR', 'CORDOBA', 'CUNDINAMARCA', 'HUILA', 'MAGDALENA', 'META', 
        'NARINO', 'NORTE DE SANTANDER', 'QUINDIO', 'RISARALDA', 'SANTANDER', 'TOLIMA', 
        'VALLE DEL CAUCA'
    }

    # 5. Validate the final department names against the allowed set.
    unique_departments_after = set(df['DepartmentProcessed'].unique())
    unhandled_departments = unique_departments_after - valid_departments

    if unhandled_departments:
        raise ValueError(f"Unhandled department names found after cleaning: {list(unhandled_departments)}")

    # 6. Delete spaces in the DepartmentProcessed column
    df['DepartmentProcessed'] = df['DepartmentProcessed'].str.replace(' ', '')

    print("All department names were successfully standardized and validated.")
    final_list = sorted(set(df['DepartmentProcessed'].unique()))
    print(f"\nFinal unique departments: {final_list}")
    print(f"Number of unique departments after conversion: {len(final_list)}")
    print("-"*50 + "\n")
    
    return df


def preprocess_department_fuel(dfs):
    """Preprocess the department column for the fuel data.
    Args:
        dfs: DataFrame or list of DataFrames to preprocess
    Returns:
        Preprocessed DataFrame(s)
    """
    print("\nPreprocessing the department column for the fuel data...")
    # Check if dfs is a list
    is_list = True
    if not isinstance(dfs, list):
        is_list = False
        dfs = [dfs]
    
    for i, df in enumerate(dfs):
        print(f"\nProcessing DataFrame {i+1} of {len(dfs)}")
        dfs[i] = preprocess_department_column(df)
    
    if is_list:
        return dfs
    else:
        return dfs[0]


def preprocess_municipality_column(df):
    """Standardizes and validates municipality names in a DataFrame.
    Performs the following transformations:
    1. Normalizes spacing and removes accents from municipality names
    2. Maps common variations to standardized municipality names
    3. Validates against a set of allowed municipality names
    4. Creates a processed column with cleaned municipality names
    Args:
        df (pd.DataFrame): Input DataFrame containing a 'Municipality' column
    Returns:
        pd.DataFrame: Original DataFrame with added 'MunicipalityProcessed' column
    Raises:
        ValueError: If any municipality names remain unhandled after processing
    """
    # Check if the Municipality column exists in the dataframe
    if not 'Municipality' in df.columns:
        print("No Municipality column found in the dataframe. Nothing to lematize.")
        print("-"*50 + "\n")
        return df

    # Preprocess the Municipality column
    # 1. Replace multiple spaces with a single space and strip spaces from Municipality
    df['Municipality'] = df['Municipality'].str.replace('  ', ' ').str.strip()
    # Remove accents from Municipality
    df['Municipality'] = df['Municipality'].str.normalize('NFKD').str.encode('ascii', errors='ignore').str.decode('utf-8')

    # Get the unique municipalities
    unique_municipalities = sorted(df['Municipality'].astype(str).unique().tolist())
    print(f"Unique municipalities: {unique_municipalities}")
    print(f"Number of unique municipalities: {len(unique_municipalities)}")

    # 2. Create a conversion dictionary for the Municipality column
    municipality_conversion_dict = {
        'BOGOTA, D.C.': 'BOGOTA',
        'CARTAGENA (DISTRITO TURISTICO Y CULTURAL DE CARTAGENA)': 'CARTAGENA',
        'GUADALAJARA DE BUGA': 'BUGA',
        'SITIONUEVO': 'SITIO NUEVO'
    }

    # 3. Apply the conversion and clean the Municipality column
    df['MunicipalityProcessed'] = df['Municipality'].replace(municipality_conversion_dict)

    # Handle cases where some Municipalities appears in Department column by moving it to MunicipalityProcessed
    df.loc[df['Department'] == 'MALAMBO', 'MunicipalityProcessed'] = 'MALAMBO'
    df.loc[df['Department'] == 'SAHAGUN', 'MunicipalityProcessed'] = 'CORDOBA'

    # Handle NaN values by filling them with empty strings
    df['MunicipalityProcessed'] = df['MunicipalityProcessed'].fillna('')

    # 4. Define the set of all valid final municipalities
    valid_municipalities = {
        '', 'AGUACHICA', 'AGUAZUL', 'AIPE', 'BARBOSA', 'BARRANCA DE UPIA', 'BARRANCABERMEJA', 
        'BARRANQUILLA', 'BELLO', 'BOGOTA', 'BOSCONIA', 'BRICENO', 'BUCARAMANGA', 'BUENAVENTURA', 'BUENAVISTA', 
        'BUGA', 'BUGALAGRANDE', 'CACERES', 'CALARCA', 'CALDAS', 'CANDELARIA', 'CAREPA', 'CARTAGENA', 
        'CAUCASIA', 'CHACHAGUI', 'CHIA', 'CHIGORODO', 'CHIQUINQUIRA', 'CIENAGA', 'COELLO', 'CORDOBA', 'COTA', 
        'CUCUTA', 'CURUMANI', 'DON MATIAS', 'DUITAMA', 'FUNZA', 'FUSAGASUGA', 'GIRARDOTA', 'GIRON', 
        'GUARNE', 'HONDA', 'IBAGUE', 'ITAGUI', 'LA FORTUNA', 'LA GLORIA', 'LA ISLA', 'LA URIBE', 
        'LOS PATIOS', 'MADRID', 'MALAMBO', 'MANIZALES', 'MARIQUITA', 'MELGAR', 'MONTERIA', 'MOSQUERA', 
        'MUTATA', 'NATAGAIMA', 'PALERMO', 'PAMPLONA', 'PARATEBUENO', 'PASTO', 'PELAYA', 'PEREIRA', 
        'PIEDECUESTA', 'PLANETA RICA', 'PONEDERA', 'POPAYAN', 'PUERTO BOYACA', 'PUERTO LIBRE', 'PUERTO LOPEZ',  
        'PUERTO SALGAR', 'PUERTO TRIUNFO', 'SABANALARGA', 'SABANETA', 'SAHAGUN', 'SAN ALBERTO', 
        'SAN GIL', 'SAN JERONIMO', 'SAN JOSE DE PARE', 'SAN JUAN NEPOMUCENO', 
        'SANTA FE DE ANTIOQUIA', 'SESQUILE', 'SIBATE', 'SIBERIA', 'SIMACOTA', 'SITIO NUEVO', 
        'SOACHA', 'SOCORRO', 'SOLEDAD', 'SOPO', 'SOTAQUIRA', 'TARAZA', 'TOCANCIPA', 'TURBANA', 'VALDIVIA', 
        'VALLEDUPAR', 'VILLANUEVA', 'VILLAVICENCIO', 'YOPAL', 'YUMBO', 'ZARZAL', 'ZIPAQUIRA'
    }

    # 5. Validate the final municipality names against the allowed set.
    unique_municipalities_after = set(df['MunicipalityProcessed'].unique())
    unhandled_municipalities = unique_municipalities_after - valid_municipalities

    if unhandled_municipalities:
        raise ValueError(f"Unhandled municipality names found after cleaning: {list(unhandled_municipalities)}")

    # 6. Delete spaces in the MunicipalityProcessed column
    df['MunicipalityProcessed'] = df['MunicipalityProcessed'].str.replace(' ', '')

    print("All municipality names were successfully standardized and validated.")
    final_municipalities_list = sorted(set(df['MunicipalityProcessed'].unique()))
    print(f"\nFinal unique municipalities: {final_municipalities_list}")
    print(f"Number of unique municipalities after conversion: {len(final_municipalities_list)}")
    print("-"*50 + "\n")

    return df


def preprocess_municipality_fuel(dfs):
    """Preprocess the municipality column for the fuel data.
    Args:
        dfs: DataFrame or list of DataFrames to preprocess
    Returns:
        Preprocessed DataFrame(s)
    """
    print("\nPreprocessing the municipality column for the fuel data...")
    # Check if dfs is a list
    is_list = True
    if not isinstance(dfs, list):
        is_list = False
        dfs = [dfs]
    
    for i, df in enumerate(dfs):
        print(f"\nProcessing DataFrame {i+1} of {len(dfs)}")
        dfs[i] = preprocess_municipality_column(df)
    
    if is_list:
        return dfs
    else:
        return dfs[0]


def add_flag_to_serv_st_lematized(dfs, required_columns):
    """
    Validates service station data by checking for multiple location mappings per flagged station.
    This function:
    1. Checks for required columns (Flag, ServStLematized, DepartmentProcessed, MunicipalityProcessed)
    2. Creates a composite key combining these columns (FlgSrvStLemDepMun)
    3. Identifies any FlagServStLematized values that map to multiple locations
    4. Reports any validation issues found
    Args:
        dfs: DataFrame or list of DataFrames to process
        required_columns: List of required column names
    Returns:
        None: Prints validation results but doesn't return anything
    """
    # Check if dfs is a list
    is_list = True
    if not isinstance(dfs, list):
        is_list = False
        dfs = [dfs]
    
    for i, df in enumerate(dfs):
        print(f"\nProcessing DataFrame {i+1} of {len(dfs)}")
        # Check if the ServStLematized column exists in the dataframe
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            print(f"Missing required columns: {missing_columns}.\nCannot add flag to ServStLematized column.")
            print("-"*50 + "\n")
            continue
            
        # Find records with multiple FlgSrvStLemDepMun per each FlagServStLematized
        df['FlgSrvStLemDepMun'] = df['Flag'] + df['ServStLematized'] + df['DepartmentProcessed'] + df['MunicipalityProcessed']

        # Group by FlagServStLematized and count unique FlgSrvStLemDepMun values
        duplicate_station_check = df.groupby('FlagServStLematized')['FlgSrvStLemDepMun'].unique()

        # Filter for ServiceStations with more than one match
        multi_sv_stations = duplicate_station_check[duplicate_station_check.apply(len) > 1]

        if not multi_sv_stations.empty:
            print(f"Found {len(multi_sv_stations)} FlagServStLematized with multiple FlgSrvStLemDepMun values:")
            # Set the option to display full content for column width
            with pd.option_context('display.max_colwidth', None):
                print(multi_sv_stations)
        else:
            print("No FlagServStLematized with multiple FlgSrvStLemDepMun values found")
        
        print("-"*50 + "\n")


def fix_multi_sv_stations(df):
    """
    Corrects multiple FlgSrvStLemDepMun mappings for service stations by standardizing them to single values.
    This function:
    1. Checks for the existence of required FlgSrvStLemDepMun column
    2. Uses a predefined mapping to consolidate multiple FlgSrvStLemDepMun values to their canonical forms
    3. Verifies the corrections were successful by checking for remaining duplicates
    Args:
        df (pd.DataFrame): Input dataframe containing service station data with FlgSrvStLemDepMun column
    Returns:
        pd.DataFrame: Dataframe with corrected FlgSrvStLemDepMun values
    Raises:
        ValueError: If multiple FlgSrvStLemDepMun values remain after correction for any FlagServStLematized
    """
    # Check if the FlgSrvStLemDepMun column exists in the dataframe
    if not 'FlgSrvStLemDepMun' in df.columns:
        print("No FlgSrvStLemDepMun column found in the dataframe. Nothing to fix.")
        print("-"*50 + "\n")
        return df

    # Create a dictionary to fix the multiple FlgSrvStLemDepMun values
    # The keys are the final record and the values are the records to match
    fix_multi_sv_stations = {
        'DISTRACOMAUTOPISTAANTIOQUIAGUARNE': ['DISTRACOMAUTOPISTAANTIOQUIAGUARNE', 'DISTRACOMAUTOPISTA'],
        'DISTRACOMBAJOSAUSAANTIOQUIACAUCASIA': ['DISTRACOMBAJOSAUSAANTIOQUIACAUCASIA', 'DISTRACOMBAJOSAUSACORDOBAMONTERIA'],
        'DISTRACOMRUTADELSOLCESARBOSCONIA': ['DISTRACOMRUTADELSOLCESARBOSCONIA', 'DISTRACOMRUTADELSOLCESARCURUMANI'],
        'DISTRACOMSANMARINOMAGDALENASITIONUEVO': ['DISTRACOMSANMARINOMAGDALENASITIONUEVO', 'DISTRACOMSANMARINOATLANTICOBARRANQUILLA'],
        'DISTRACOMURIBEVALLEDELCAUCABUGALAGRANDE': ['DISTRACOMURIBEVALLEDELCAUCABUGALAGRANDE', 'DISTRACOMURIBEVALLEDELCAUCABUGA', 'DISTRACOMURIBEVALLEDELCAUCALAURIBE'],
        'DISTRACOMYECORDOBASAHAGUN': ['DISTRACOMYECORDOBASAHAGUN', 'DISTRACOMYECORDOBACORDOBA'],
        'ECOMTRANSBESOTECESARAGUACHICA': ['ECOMTRANSBESOTECESARAGUACHICA', 'ECOMTRANSBESOTECESARLAGLORIA'],
        'TERPELDAJOTASANTANDERBARRANCABERMEJA': ['TERPELDAJOTASANTANDERBARRANCABERMEJA', 'TERPELDAJOTASANTANDERLAFORTUNA'],
        'TERPELMELANYCUNDINAMARCAPUERTOSALGAR': ['TERPELMELANYCUNDINAMARCAPUERTOSALGAR', 'TERPELMELANYCUNDINAMARCAPUERTOLIBRE'],
        'TERPELMULERANORTEDESANTANDERLOSPATIOS': ['TERPELMULERANORTEDESANTANDERLOSPATIOS', 'TERPELMULERANORTEDESANTANDERCUCUTA'],
        'TERPELSARISADACUNDINAMARCASOPO': ['TERPELSARISADACUNDINAMARCASOPO', 'TERPELSARISADACUNDINAMARCABRICENO'],
    }

    # Find in the FlgSrvStLemDepMun column the records that are in the values of the fix_multi_sv_stations dictionary and replace them with the keys
    # Invert the dictionary to map old values to the new correct value
    value_to_key_map = {old_val: key for key, old_vals in fix_multi_sv_stations.items() for old_val in old_vals}

    # Replace all occurrences of keys found in the `value_to_key_map`.
    df['FlgSrvStLemDepMun'] = df['FlgSrvStLemDepMun'].replace(value_to_key_map)

    # Verification Step
    # Re-run the check to confirm that the duplicates have been resolved.
    print("Verifying the changes...")
    duplicate_station_check_after_fix = df.groupby('FlagServStLematized')['FlgSrvStLemDepMun'].unique()
    multi_sv_stations_after_fix = duplicate_station_check_after_fix[duplicate_station_check_after_fix.apply(len) > 1]

    if multi_sv_stations_after_fix.empty:
        print("Success! All FlgSrvStLemDepMun entries have been corrected.")
        print(f"Number of unique FlgSrvStLemDepMun values after correction: {len(duplicate_station_check_after_fix)}")
    else:
        print(f"Found {len(multi_sv_stations_after_fix)} FlagServStLematized with multiple FlgSrvStLemDepMun values after correction:")
        # Set the option to display full content for column width to see the full IDs
        with pd.option_context('display.max_colwidth', None):
            print(multi_sv_stations_after_fix)
        # Raise an error to stop the execution
        raise ValueError("Multiple FlgSrvStLemDepMun values found after correction.")
    print("-"*50 + "\n")

    return df


def fix_multi_sv_stations_disclosed(dfs):
    """
    Corrects multiple FlgSrvStLemDepMun mappings for service stations by standardizing them to single values.
    This wrapper function handles both single DataFrames and lists of DataFrames by:
    1. Checking input type and normalizing to list format
    2. Applying fix_multi_sv_stations to each DataFrame
    3. Returning results in the same format as input
    Args:
        dfs (pd.DataFrame or list[pd.DataFrame]): Input dataframe(s) containing service station data
    Returns:
        pd.DataFrame or list[pd.DataFrame]: Corrected dataframe(s) in same format as input
    Raises:
        ValueError: If multiple FlgSrvStLemDepMun values remain after correction (from fix_multi_sv_stations)
    """
    print("\nFixing multiple FlgSrvStLemDepMun mappings for service stations by standardizing them to single values...")
    # Check if dfs is a list
    is_list = True
    if not isinstance(dfs, list):
        is_list = False
        dfs = [dfs]
    
    for i, df in enumerate(dfs):
        print(f"\nProcessing DataFrame {i+1} of {len(dfs)}")
        dfs[i] = fix_multi_sv_stations(df)
    
    if is_list:
        return dfs
    else:
        return dfs[0]


def box_plot_gls_by_serv_st(dfs):
    """
    Create box plots for Gls and PricePerGl grouped by FlgSrvStLemDepMun for each passed DataFrame
    Handles both single DataFrame and list of DataFrames
    Args:
        dfs (pd.DataFrame or list[pd.DataFrame]): Input dataframe(s) to plot
    """
    # Normalize input to list format
    if not isinstance(dfs, list):
        dfs = [dfs]
    
    for i, df in enumerate(dfs):
        print(f"\nPlotting DataFrame {i+1} of {len(dfs)}")
        
        if not 'FlgSrvStLemDepMun' in df.columns:
            print("No FlgSrvStLemDepMun column found in the dataframe. Nothing to plot.")
            print("-"*50)
            continue
            
        percentile = 0.95
        plt.figure(figsize=(16, 8))

        gls_cutoff = df[df['Gls'] > 0]['Gls'].quantile(percentile)
        price_cutoff = df[df['PricePerGl'] > 0]['PricePerGl'].quantile(percentile)

        plt.subplot(1, 2, 1)
        sns.boxplot(data=df[(df['Gls'] > gls_cutoff)], y='Gls', x='FlgSrvStLemDepMun')
        plt.title(f'Gallons Purchased by Service Station (>{(percentile*100):.0f}th Percentile)')
        plt.ylabel('Gallons')
        plt.xlabel('Service Station')
        plt.xticks(rotation=90, fontsize=8)
        plt.gca().yaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: '{:,.0f}'.format(x)))

        plt.subplot(1, 2, 2)
        sns.boxplot(data=df[(df['PricePerGl'] > price_cutoff)], y='PricePerGl', x='FlgSrvStLemDepMun') 
        plt.title(f'Price Per Gallon by Service Station (>{(percentile*100):.0f}th Percentile)')
        plt.ylabel('Price ($)')
        plt.xlabel('Service Station')
        plt.xticks(rotation=90, fontsize=8)
        plt.gca().yaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: '{:,.0f}'.format(x)))

        plt.tight_layout()
        plt.show()

        print("-"*50)


def outlier_threshold_detection(df):
    """
    Detect the outliers in the PricePerGl column
    """
    if not 'FlagServStLematized' in df.columns:
        print("No FlagServStLematized column found in the dataframe. No outliers to detect.")
        print("-"*50 + "\n")
        return df
    
    # Detect and remove outliers in the PricePerGl column
    ## Calculate Q1, Q3, and the IQR for 'PricePerGl'. Consider only values above 0.
    q1 = df[df['PricePerGl'] > 0]['PricePerGl'].quantile(0.25)
    q3 = df[df['PricePerGl'] > 0]['PricePerGl'].quantile(0.75)

    ## Filter the DataFrame to values within the IQR
    iqr_df = df[(df['PricePerGl'] >= q1) & (df['PricePerGl'] <= q3)]

    ## Calculate the mean of 'PricePerGl' within the IQR
    iqr_mean = iqr_df['PricePerGl'].mean()

    ## Define the outlier threshold
    outlier_threshold = iqr_mean * 3

    return outlier_threshold, iqr_mean


def detect_outliers_price_per_gl(dfs):
    """
    Detect outliers in the PricePerGl column across multiple dataframes
    Args:
        dfs: Dictionary of dataframes to process
    Returns:
        dict: Dictionary of processed dataframes with outlier information
    """
    print("\nDetecting outliers in the PricePerGl column...")
    if not isinstance(dfs, list):
        dfs = [dfs]
    
    for i, df in enumerate(dfs):
        print(f"Processing DataFrame {i+1} of {len(dfs)}. Index: {i}")

        # Get the outlier threshold
        outlier_threshold, iqr_mean = outlier_threshold_detection(df)

        ## Get outliers and their values
        outliers = df[df['PricePerGl'] >= outlier_threshold]
        outliers_count = len(outliers)
        outliers_values = outliers['PricePerGl'].values
        outliers_stations = outliers['FlagServStLematized'].values

        if outliers_count > 0:
            print(f"Found {outliers_count} outliers above the threshold of {outlier_threshold:.2f} with values: {outliers_values}")
            print(f"Affected service stations: {outliers_stations}") 
        else:
            print("No outliers found in the PricePerGl column.")
        print("-"*50 + "\n")
    

def handle_outliers_price_per_gl(dfs, index_val_to_replace_dict):
    """
    Change the outliers values in the PricePerGl column
    """
    print("Handling outliers in the PricePerGl column...")
    is_list = True
    if not isinstance(dfs, list):
        is_list = False
        dfs = [dfs]
    
    for i, val_to_replace in index_val_to_replace_dict.items():
        # Get the outlier threshold
        outlier_threshold, iqr_mean = outlier_threshold_detection(dfs[i])

        # If no replacement value is provided, use the IQR mean
        if val_to_replace is None or val_to_replace == '':
            val_to_replace = iqr_mean

        if (dfs[i]['PricePerGl'] >= outlier_threshold).any():
            dfs[i]['PricePerGl'] = np.where(dfs[i]['PricePerGl'] >= outlier_threshold, val_to_replace, dfs[i]['PricePerGl'])
            print(f"From index {i}, replaced outliers with value: {val_to_replace:.2f}")
        else:
            print(f"No outliers found to replace in index {i}.")
        print("-"*50 + "\n")
    
    if is_list:
        return dfs
    else:
        return dfs[0]


def calculate_total_purchase_value(dfs, galons_col_name, price_per_gl_col_name):
    """
    Calculate total purchase value by ensuring numeric columns and multiplying Gls * PricePerGl
    Args: df (pd.DataFrame): Input dataframe containing fuel purchase data
    Returns: pd.DataFrame: Dataframe with added TotalPurchaseValue column and cleaned numeric columns
    """
    is_list = True
    if not isinstance(dfs, list):
        is_list = False
        dfs = [dfs]
    
    for i, df in enumerate(dfs):
        print(f"Processing DataFrame {i+1} of {len(dfs)}")
        
        # Ensure 'Gls' and 'PricePerGl' are numeric, coercing errors if necessary
        df[galons_col_name] = pd.to_numeric(df[galons_col_name], errors='coerce')
        df[price_per_gl_col_name] = pd.to_numeric(df[price_per_gl_col_name], errors='coerce')
        
        # Drop rows where conversion failed if necessary, or handle them appropriately
        print(f"Shape before dropping {df.shape}")
        df.dropna(subset=[galons_col_name, price_per_gl_col_name], inplace=True)
        print(f"Shape after dropping na values {df.shape}")
        
        # Calculate TotalPurchaseValue
        df['TotalPurchaseValue'] = df[galons_col_name] * df[price_per_gl_col_name]
        
        print("-"*50 + "\n")

    if is_list:
        return dfs
    else:
        return dfs[0]


def calculate_period_gl_weighted_avg(df, month):
    """Calculate weighted average fuel price for a period (month).
    Args:
        df: DataFrame with fuel purchase data
        month: Month identifier (e.g. '01/01/25')
    Returns:
        float: Weighted average price per gallon
    """
    total_gl = df['Gls'].sum()
    gls_weight = df['Gls'] / total_gl
    weighted_price = (gls_weight * df['PricePerGl']).sum()
    print(f"-Weighted average price per gl for month ({month}): ${int(weighted_price):,.2f}")
    
    return weighted_price


def calculate_weighted_avg_by_buyer(df, month):
    """
    Calculate weighted average price per gallon by buyer for a given month.
    Args:
        df (pd.DataFrame): DataFrame containing fuel purchase data
        month (str): Month identifier to include in print statements
    Returns:
        dict: Dictionary containing buyer results with weighted averages, 
              sorted by total gallons in descending order
    """
    buyer_groups = df.groupby('Buyer')
    buyer_results = {}

    for buyer, group in buyer_groups:
        total_gl_buyer = group.Gls.sum()
        gls_weight_buyer = group.Gls / total_gl_buyer
        weighted_price_buyer = (gls_weight_buyer * group.PricePerGl).sum()
        buyer_results[buyer] = {
            'weighted_avg_price': weighted_price_buyer,
            'total_gallons': total_gl_buyer,
            'num_stations': len(group)
        }

    # Sort by total gallons in descending order
    buyer_results = dict(sorted(buyer_results.items(), key=lambda x: x[1]['total_gallons'], reverse=True))

    print(f"-Weighted average price per gallon by buyer ({month}). Organized by total gallons:")
    for buyer, data in buyer_results.items():
        print(f"    ${data['weighted_avg_price']:,.2f} for {buyer} (from {data['num_stations']} stations, {data['total_gallons']:,.2f} gallons)")
    
    return buyer_results


def get_period_weighted_average(dfs):
    """Calculate weighted average fuel price for a period.
    Args:
        df: DataFrame with fuel purchase data
    Returns:
        float: Weighted average price per gallon
    Raises:
        ValueError: If duplicate months are found across different DataFrames
    """
    weighted_prices_periods = {}
    unique_months = set()
    
    for i, df in enumerate(dfs):
        print(f"\nProcessing DataFrame {i+1} of {len(dfs)}. Index: {i}")
        weighted_prices_periods[i] = {}
        months = df['PurchaseMonth'].unique()
        
        for month in months:
            # Check for duplicate months
            if month in unique_months:
                raise ValueError(f"Duplicate month found: {month}. Each month must be unique across all DataFrames.")
            unique_months.add(month)
            
            # Calculate weighted average price per gallon for the period and by buyer
            df_month = df[df['PurchaseMonth'] == month]
            weighted_price = calculate_period_gl_weighted_avg(df_month, month)
            buyer_results = calculate_weighted_avg_by_buyer(df_month, month)
            
            weighted_prices_periods[i][month] = {'weighted_price': weighted_price,
                                                 'buyer_results': buyer_results}

    return weighted_prices_periods


def render_scatter_plot_weighted_avg_price_by_total_gallons(buyer_results, month, weighted_price):
    """Render a scatter plot of the weighted average price per gallon by total gallons"""
    plt.figure(figsize=(12, 6))

    # Extract x (total_gallons) and y (weighted_avg_price) values
    x_values = [v['total_gallons'] for v in buyer_results.values()]
    y_values = [v['weighted_avg_price'] for v in buyer_results.values()]
    buyers = list(buyer_results.keys())

    plt.scatter(
        x_values,  # x values: total_gallons
        y_values,   # y values: weighted_avg_price
        s=80,
        color='blue'
    )

    # Handle single data point case for x-axis
    if len(x_values) == 1:
        padding = x_values[0] * 0.1 if x_values[0] > 0 else 10
        plt.xlim(x_values[0] - padding, x_values[0] + padding)

    # Add point values (Buyer and Price) to scatter plot
    vertical_offset = 8
    for i, buyer in enumerate(buyers):
        if buyer == 'CARGA' or buyer == 'TRANSER':
            plt.text(x_values[i], y_values[i] + vertical_offset, f"{buyer}\n${y_values[i]:,.0f}",
                    ha='center', va='bottom', fontsize=10, fontweight='bold')
        else:
            plt.text(x_values[i], y_values[i] + vertical_offset, f"{buyer}\n${y_values[i]:,.0f}",
                    ha='center', va='bottom', fontsize=10)

    # Horizontal Line for Overall Weighted Average
    plt.axhline(y=weighted_price, color='blue', linestyle='--', alpha=0.5, linewidth=2)
    
    # Position text near the left edge, slightly above the line
    xmin, xmax = plt.gca().get_xlim()
    text_x_pos = xmin + (xmax - xmin) * 0.05  # 5% from the left edge
    text_y_pos = weighted_price + vertical_offset / 2 # Position slightly above the line
    plt.text(text_x_pos, text_y_pos, f"Total Average (${weighted_price:,.0f})",
            ha='left', va='bottom', color='blue', fontsize=10, alpha=0.7)

    plt.ylabel('Weighted Average Price per Gallon')
    plt.xlabel('Total Gallons Purchased')
    plt.title(f'Weighted Average Price per Gl vs. Total Gls Purchased ({month})', fontsize=12, fontweight='bold', pad=10)
    plt.grid(True, linestyle='--', alpha=0.6)

    # Format axis with thousands separators
    plt.gca().xaxis.set_major_formatter(lambda x, _: f"{x:,.0f}")
    plt.gca().yaxis.set_major_formatter(lambda x, _: f"${x:,.0f}")

    # Add some padding to the y-axis limits to ensure text fits
    if y_values:
        plt.ylim(min(y_values) - vertical_offset * 3, max(y_values) + vertical_offset * 7)

    plt.tight_layout() # Adjust layout to prevent labels overlapping plot area
    plt.show()


def render_scatter_weighted_avg_price_by_gallons(weighted_prices_periods):
    """
    Render scatter plots of weighted average price vs. total gallons for each period.
    This function iterates through the nested dictionary of weighted prices 
    and calls the plotting function for each month's data.
    Args:
        weighted_prices_periods (dict): A nested dictionary containing buyer results and
                                        weighted prices for different periods and months.
                                        Example structure: 
                                        {df_index: {month: {'weighted_price': ..., 'buyer_results': ...}}}
    """
    for df_index, month_data in weighted_prices_periods.items():
        for month, period_data in month_data.items():
            buyer_results = period_data['buyer_results']
            weighted_price = period_data['weighted_price']
            render_scatter_plot_weighted_avg_price_by_total_gallons(buyer_results, month, weighted_price)


def create_summary_pivot_table(df):
    """
    Creates a pivot table summarizing fuel prices by service station and buyer.

    This function handles DataFrames both with and without geographical information
    (Department, Municipality). It calculates row-wise and weighted averages
    and formats the output for clear presentation. The output DataFrame will always
    contain 'Department' and 'ServiceStation' columns for structural consistency.

    Args:
        df (pd.DataFrame): The input DataFrame. Must contain 'PricePerGl',
                           'FlagServStLematized', 'Buyer', and 'Gls'.
                           It should also contain 'ServiceStation', and optionally 
                           'Municipality' and 'DepartmentProcessed'.

    Returns:
        pd.DataFrame: A formatted and sorted pivot table.
    """
    # 1. Create the base pivot table
    pivot_df = pd.pivot_table(
        df,
        values='PricePerGl',
        index='FlagServStLematized',
        columns='Buyer',
        aggfunc='mean'
    )

    # 2. Calculate and add the average station price (row-wise mean)
    pivot_df['AverageStationPrice'] = pivot_df.mean(axis=1)

    # 3. Calculate the weighted average price for each buyer
    buyers = list(df['Buyer'].unique())
    buyer_weighted_avg = {}
    for buyer in buyers:
        buyer_data = df[df['Buyer'] == buyer]
        total_gl_buyer = buyer_data['Gls'].sum()
        if total_gl_buyer > 0:
            weighted_price_buyer = (buyer_data['Gls'] * buyer_data['PricePerGl']).sum() / total_gl_buyer
            buyer_weighted_avg[buyer] = weighted_price_buyer
        else:
            buyer_weighted_avg[buyer] = np.nan

    # 4. Create and append the weighted average row to the pivot table
    weighted_avg_row_name = 'Weighted_Buyer_Average'
    weighted_avg_series = pd.Series(buyer_weighted_avg, name=weighted_avg_row_name)
    pivot_df = pd.concat([pivot_df, pd.DataFrame(weighted_avg_series).T])

    # Conditionally handle geographical data if it exists
    has_geo_data = all(col in df.columns for col in ['Municipality', 'DepartmentProcessed', 'ServiceStation'])

    if has_geo_data:
        # Get unique station and department info for labeling and sorting
        unique_stations_df = df.drop_duplicates(subset=['FlagServStLematized']).copy()

        # Clean up ServiceStation names for better labels
        words_to_remove = r'\b(EDS|ALIADA|AUTOMOTRIZ|SUPER|LA)\b'
        unique_stations_df['ServiceStation'] = (
            unique_stations_df['ServiceStation']
            .str.replace(words_to_remove, '', regex=True)
            .str.replace(r'\s+', ' ', regex=True)
            .str.strip()
        )
        
        # Create descriptive labels and mappings
        unique_stations_df['ServiceStationLabel'] = unique_stations_df['Municipality'].str[:15] + ' | ' + unique_stations_df['ServiceStation']
        station_mapping = unique_stations_df.set_index('FlagServStLematized')['ServiceStationLabel']
        department_mapping = unique_stations_df.set_index('FlagServStLematized')['DepartmentProcessed']

        pivot_df['ServiceStation'] = pivot_df.index.map(station_mapping)
        pivot_df['DepartmentProcessed'] = pivot_df.index.map(department_mapping)

        # Handle the summary row
        pivot_df.loc[weighted_avg_row_name, 'ServiceStation'] = ''
        pivot_df.fillna({'DepartmentProcessed': 'zzz_Weighted_Average'}, inplace=True)

        # Sort by department, then by price
        pivot_df = pivot_df.sort_values(by=['DepartmentProcessed', 'AverageStationPrice'], ascending=[True, False])
        
        # Finalize column names and order
        pivot_df.rename(columns={'DepartmentProcessed': 'Department'}, inplace=True)
        pivot_df.loc[pivot_df['Department'] == 'zzz_Weighted_Average', 'Department'] = ''
    else:
        # Fallback for DataFrames without geographical data
        pivot_df['Department'] = ''
        if 'ServiceStation' in df.columns:
            station_mapping = df.drop_duplicates('FlagServStLematized').set_index('FlagServStLematized')['ServiceStation']
            pivot_df['ServiceStation'] = pivot_df.index.map(station_mapping)
        else:
            pivot_df['ServiceStation'] = pivot_df.index

        pivot_df.loc[weighted_avg_row_name, 'ServiceStation'] = ''
        pivot_df.loc[weighted_avg_row_name, 'Department'] = ''
        
        # Sort by price, keeping the average row at the bottom
        avg_row = pivot_df.loc[[weighted_avg_row_name]]
        data_rows = pivot_df.drop(weighted_avg_row_name)
        data_rows = data_rows.sort_values(by='AverageStationPrice', ascending=False)
        pivot_df = pd.concat([data_rows, avg_row])
    
    # Define and ensure final column order, then return
    final_columns = ['Department', 'ServiceStation'] + buyers + ['AverageStationPrice']
    existing_cols = [col for col in final_columns if col in pivot_df.columns]
    
    pivot_df_fuel_month = pivot_df[existing_cols]
    display(pivot_df_fuel_month.head(3))
    return pivot_df_fuel_month


def create_summary_pivot_table_disclosed_fuel(dfs):
    """
    Creates pivot tables summarizing fuel prices for each month in each DataFrame.

    For each DataFrame in the input, this function identifies unique months in
    the 'PurchaseMonth' column and generates a separate pivot table for each month.
    Args:
        dfs (pd.DataFrame or list of pd.DataFrame): A single DataFrame or a list of
                                                     DataFrames with fuel purchase data.
    Returns:
        list of pd.DataFrame: A list containing all generated pivot tables.
    """
    if not isinstance(dfs, list):
        dfs = [dfs]

    pivot_tables_disclosed_fuel_list = []
    for i, df in enumerate(dfs):
        name = df.name
        print(f"\nProcessing DataFrame {i+1} of {len(dfs)} for pivot tables.")

        if 'PurchaseMonth' not in df.columns:
            raise ValueError(f"Warning: 'PurchaseMonth' column not in DataFrame {i+1}. Creating single pivot table.")

        months = df['PurchaseMonth'].unique()
        
        # Split the name and pass it to the pivot table
        ## Get the prefix of the name before the first digit (Search for the '_digit' structure)
        name_prefix = name[:name.find(next((f"_{d}" for d in name.split('_') if d.isdigit()), ''))]
        periods = [n for n in name.split('_') if n.isdigit()]  # Gets all numeric parts from split string
        for i, month in enumerate(months):
            print(f"-Creating pivot table for month: {month}")
            df_month = df[df['PurchaseMonth'] == month].copy()
            pivot_df = create_summary_pivot_table(df_month)
            pivot_df.name = name_prefix + '_' + periods[i]
            pivot_tables_disclosed_fuel_list.append(pivot_df)

    return pivot_tables_disclosed_fuel_list


def set_index_as_pivot_df_column(dfs):
    """
    Convert pivot table to dataframe and rename index column as FlagServStLematized.
    Args: pivot_df: Pivot table dataframe to process
    Returns: Processed dataframe with index as column
    """
    is_list = True
    if not isinstance(dfs, list):
        is_list = False
        dfs = [dfs]

    for i, pivot_df in enumerate(dfs):
        name = pivot_df.name
        dfs[i] = pivot_df.reset_index()
        dfs[i].rename(columns={'index': 'FlagServStLematized'}, inplace=True)
        dfs[i].name = name
        
        print(f"Setting index on the {i+1} dataframe out of {len(dfs)}")
        display(dfs[i].tail(4))

    if is_list:
        return dfs
    else:
        return dfs[0]


def fix_unnamed_headers(df):
    """
    Fixes DataFrames where most columns are Unnamed by using first row as header.
    Args:
        df: DataFrame to check and potentially fix
    Returns:
        Cleaned DataFrame with proper headers
    Raises:
        ValueError: If headers still mostly Unnamed after fix attempt
    """
    if sum(isinstance(col, str) and 'Unnamed' in col for col in df.columns) / len(df.columns) > 0.5:
        print("Most cols Unnamed... dropping first row and using as header")
        temp_df = pd.DataFrame(df.values[1:], columns=df.iloc[0]).copy()
        
        if sum(isinstance(col, str) and 'Unnamed' in col for col in temp_df.columns) / len(temp_df.columns) > 0.5:
            raise ValueError("Still most cols Unnamed after fixing. Check the file manually")
        else:
            print("Success! Most columns are named after fixing.")
            return temp_df
    return df


def preprocess_quotation_df(df, supplier, price_column):
    """
    Preprocesses the quotation dataframe by:
    1. Standardizing column names (uppercase, underscores, replace non-standard columns names)
    2. Filtering valid records (non-null, non-zero, not 'CERRADO')
    3. Cleaning text fields (uppercase, underscore formatting)
    4. Dropping empty columns
    5. Adding supplier identification
    6. Tracking price column name changes
    Args:
        df (pd.DataFrame): Input dataframe to process
        supplier (str): Name of supplier to add as 'BUYER' column
        price_column (str): Original name of price column to track through standardization
    Returns:
        tuple: 
            - pd.DataFrame: Processed dataframe with standardized format
            - str: New standardized name of price column (empty string if not found)
    """
    new_price_col = ""
    original_columns = df.columns.copy()

    # Standardize column names
    ## Uppercase the column names and replace spaces, en-dash (longer than a regular dash) and dash with underscores
    new_columns = df.columns.str.strip().str.upper().str.replace(r'[\s–-]+', '_', regex=True)
    ## Replace accents and ñ with their non-accented equivalents
    new_columns = new_columns.str.normalize('NFKD').str.encode('ascii', errors='ignore').str.decode('ascii')
    ## Replace the columns names
    df.columns = new_columns

    # Track the new name of the original price column
    if price_column in original_columns:
        price_col_index = list(original_columns).index(price_column)
        new_price_col = new_columns[price_col_index]

    # Replace columns name if they exist
    df.rename(columns={k: v for k, v in {'CD_DESTINO': 'COD_DESTINO', 
                                         'CD_ORIGEN': 'COD_ORIGEN'}.items() 
                      if k in df.columns}, inplace=True)

    # First validate the columns has EDS, CODIGO_RUTA, COD_ORIGEN, COD_DESTINO and TIPO
    # because some DFs do not have these columns (E.g. Otransa, ADP, Iceberg)
    if all(col in df.columns for col in ['EDS', 'CODIGO_RUTA', 'COD_ORIGEN', 'COD_DESTINO', 'TIPO']):
        # Exclude records that don't follow the criteria
        df = df[(df['EDS'].notna()) & 
                (df['EDS'] != '-') & 
                (df['CODIGO_RUTA'].notna()) &
                (df['CODIGO_RUTA'] != 0) &
                (df['CODIGO_RUTA'].str.upper() != 'CERRADO') &
                (df['COD_ORIGEN'].notna()) & 
                (df['COD_ORIGEN'] != 0) & 
                (df['COD_ORIGEN'].str.upper() != 'CERRADO') &
                (df['COD_DESTINO'].notna()) &
                (df['COD_DESTINO'] != 0) &
                (df['COD_DESTINO'].str.upper() != 'CERRADO')
                ].copy()
        # Make the 'TIPO' column uppercase, strip spaces and replace spaces with underscores
        df['TIPO'] = df['TIPO'].str.upper().str.strip().str.replace(r'\s+', '_', regex=True)
        df['ORIGEN'] = df['ORIGEN'].str.upper().str.strip().str.replace(r'\s+', '_', regex=True)
        df['DESTINO'] = df['DESTINO'].str.upper().str.strip().str.replace(r'\s+', '_', regex=True)

    # Exclude colums where all records are NaN
    df.dropna(axis=1, how='all', inplace=True)

    # Include a suppliers name column
    df['BUYER'] = supplier
    
    return df, new_price_col


def remove_duplicate_routes(df, df_name=""):
    """
    Check for duplicate routes in a dataframe based on origin/destination columns.
    Optionally drops duplicates and returns cleaned dataframe.
    Args:
        df: DataFrame to check for duplicates
    Returns:
        Cleaned DataFrame with duplicates removed
    """
    # Define minimum required columns
    min_cols = ['CODIGO_RUTA', 'ORIGEN','COD_ORIGEN', 'DESTINO', 'COD_DESTINO', 'TIPO']
    
    # Check if minimum columns exist, return original df if not
    if not all(col in df.columns for col in min_cols):
        print(f"Warning: Minimum columns not found in {df_name}. Returning original dataframe.")
        return df
        
    # Create subset of columns to check for duplicates
    subset = min_cols.copy()
    additional_cols = ['MUNICIPIO_ORIGEN', 'MUNICIPIO_DESTINO']
    subset.extend([col for col in additional_cols if col in df.columns])
    
    # Find and remove duplicates
    duplicates = df.duplicated(subset=subset, keep='first')
    print(f"Number of duplicates in {df_name}: {duplicates.sum()}")
    
    if duplicates.any():
        print("Sample of duplicates:")
        display(df[subset][duplicates].head())
        
        df = df[~duplicates].copy()
        print("Above duplicates dropped from dataframe\n")
        
    return df


def preprocess_quotations_dfs(dfs):
    """
    Preprocesses all quotations dataframes in a list.
    Args:
        dfs: List of quotations dataframes, each must have a 'supplier' attribute
    Returns:
        List of preprocessed quotations dataframes
    Raises:
        AttributeError: If any dataframe lacks the required 'supplier' attribute
    """
    is_list = True
    if not isinstance(dfs, list):
        is_list = False
        dfs = [dfs]
    
    for i, df in enumerate(dfs):
        if not hasattr(df, 'supplier'):
            raise AttributeError(f"DataFrame at index {i} is missing required 'supplier' attribute")
        if not hasattr(df, 'name'):
            raise AttributeError(f"DataFrame at index {i} is missing required 'name' attribute")

        # Store attributes before they are lost
        supplier = df.supplier
        name = df.name
        price_column = getattr(df, 'price_column', None)
        
        # Preprocess and re-assign attributes to the new dataframe
        processed_df, new_price_col = preprocess_quotation_df(df, supplier, price_column)
        
        # Remove duplicates using the new dataframe
        processed_df = remove_duplicate_routes(processed_df, name)
        
        # Re-assign attributes
        processed_df.supplier = supplier
        processed_df.name = name
        if new_price_col:
            processed_df.price_column = new_price_col

        # Re-assign the fully processed dataframe to the list
        dfs[i] = processed_df
    
    if is_list:
        return dfs
    else:
        return dfs[0]


def split_eds_string(eds_string):
    """
    Splits an EDS string by the first '/' that is not enclosed within any pair of parentheses. 
    Handles NaN values.
    Returns a list of strings.
    """
    if pd.isna(eds_string):
        return [''] # Return list with empty string for NaN

    eds_string = str(eds_string) # Ensure it's a string
    parenthesis_level = 0
    split_index = -1

    for i, char in enumerate(eds_string):
        if char == '(':
            parenthesis_level += 1
        elif char == ')':
            # Avoid negative level if closing parenthesis comes first
            parenthesis_level = max(0, parenthesis_level - 1)
        elif char == '/' and parenthesis_level == 0:
            split_index = i
            break # Found the first valid slash

    if split_index != -1:
        # Split at the found index
        part1 = eds_string[:split_index].strip()
        part2 = eds_string[split_index + 1:].strip()
        return [part1, part2]
    else:
        # No valid slash found, return the original string in a list
        return [eds_string.strip()]


def split_eds_string_and_explode(dfs, specific_suppliers_to_split=[]):
    """
    Splits and explodes EDS column in DataFrame(s) for specified suppliers.
    For each supplier in specific_suppliers_to_split (or all if empty), splits EDS strings at first 
    unenclosed '/', preserves originals in EDS_ORIGINAL, and maintains DataFrame attributes.
    Args:
        dfs: Single DataFrame or list of DataFrames with EDS column and supplier/name attributes
        specific_suppliers_to_split: List of supplier names to process (empty list processes all)
    Returns:
        New DataFrame(s) with split/exploded EDS column, preserving input structure
    Raises:
        AttributeError: Missing required 'supplier' or 'name' attributes
        ValueError: EDS column missing for specified suppliers
    Note: Non-destructive operation - always returns new DataFrame(s)
    """
    is_list = True
    if not isinstance(dfs, list):
        dfs = [dfs]
        is_list = False
    
    # Create a new list to store modified dataframes and keep the original attributes
    new_dfs = []
    for df in dfs:
        # Validate if the df has the supplier and name attributes
        if not hasattr(df, 'supplier'):
            raise AttributeError(f"DataFrame at index {i} is missing required 'supplier' attribute")
        if not hasattr(df, 'name'):
            raise AttributeError(f"DataFrame at index {i} is missing required 'name' attribute")
        # Get attributes before copying the dataframe
        supplier = df.supplier
        name = df.name
        price_column = getattr(df, 'price_column', None)
        # Copy the dataframe
        df_copy = df.copy()
        # Re-assign attributes that may have been lost
        df_copy.supplier = supplier
        df_copy.name = name
        if price_column:
            df_copy.price_column = price_column
        # Add the new dataframe to the list
        new_dfs.append(df_copy)
    
    # Default behavior: process all suppliers if no specific ones provided
    process_all_suppliers = not specific_suppliers_to_split

    to_pass_dfs = []
    for i, df in enumerate(new_dfs):
        supplier = df.supplier
        name = df.name
        price_column = getattr(df, 'price_column', None)

        # Process if: 
        # 1. Current supplier is in the specified list
        # 2. No specific suppliers were specified (process_all_suppliers=True)
        if supplier in specific_suppliers_to_split or process_all_suppliers:
            if 'EDS' in df.columns:
                print(f"\n{supplier} - {name}.shape: {df.shape} - Including EDS field to split and explode")
                df_exploded = df.copy()  # Create a copy to avoid modifying the original DataFrame
                df_exploded['EDS_ORIGINAL'] = df_exploded['EDS']
                df_exploded['EDS'] = df_exploded['EDS'].apply(split_eds_string)
                df_exploded = df_exploded.explode('EDS').reset_index(drop=True)
                
                print(f"df_exploded.shape: {df_exploded.shape} - EDS exploded")
                display(df_exploded[['EDS_ORIGINAL', 'EDS']].drop_duplicates(subset=['EDS_ORIGINAL']).head(3))
                
                # Re-assign attributes that may have been lost
                df_exploded.supplier = supplier
                df_exploded.name = name
                if price_column:
                    df_exploded.price_column = price_column

                to_pass_dfs.append(df_exploded)
            else:
                print(f"Warning: EDS column not found in {supplier} - {name}. Skipping...")
            print(f"-" * 50)

    if is_list:
        return to_pass_dfs
    else:
        return to_pass_dfs[0]


def add_flag_column_and_print_stats(df_list, unique_flags):
    """
    Adds FLAG column to each dataframe in list by extracting flag from EDS column,
    then prints flag distribution statistics and dataframe shapes.
    Processes only dataframes containing 'EDS' column. For each qualifying dataframe:
    1. Adds FLAG column by searching EDS strings for flags from unique_flags list
    2. Prints value counts of found flags (including NaN values)
    3. Prints dataframe name and shape
    Args:
        df_list (list): List of pandas DataFrames to process
        unique_flags (list): List of string flags to search for in EDS column
    Returns:
        None: Modifies dataframes in-place and prints statistics
    """
    for df in df_list:
        if 'EDS' in df.columns:
            df['FLAG'] = df['EDS'].apply(lambda x: find_flag_in_eds(x, unique_flags))
    
            # Print stats for the last CARGA and TRANSER dataframes as examples
            print(f"Flag distribution in {df.name}:")
            print(df['FLAG'].value_counts(dropna=False))

            print(f"{df.name}.shape: {df.shape}\n")
            print(f"-" * 50)


def lematize_eds_columns(dfs, unique_flags):
    """
    Lematize EDS columns for dataframes with EDS column.
    Creates EDS_LEMATIZED column and applies lematization.
    Args:
        dfs (list): List of dataframes
        unique_flags: Unique flags for lematization
    """
    # Process dataframes
    for df in dfs:
        if 'EDS' in df.columns:
            print(f"{df.name}.shape before lematization: {df.shape}")
            df['EDS_LEMATIZED'] = df['EDS']
            lematize_serv_st(df, "EDS_LEMATIZED", unique_flags)
        
            # Print shapes for debugging
            print(f"{df.name}.shape after lematization:  {df.shape}")
            print(f"-" * 50)


def create_flag_eds_lematized_column(dfs):
    """Create FLAG_EDS_LEMATIZED column by combining FLAG and EDS_LEMATIZED columns.
    Handles NaN values by filling with empty strings before concatenation.
    Args:
        df: DataFrame to modify
    Returns:
        Modified DataFrame with new column
    """
    for df in dfs:
        if 'EDS' in df.columns:
            print(f"{df.name}.shape before adding FLAG_EDS_LEMATIZED column: {df.shape}")
            
            # Create the FLAG_EDS_LEMATIZED column by combining FLAG and EDS_LEMATIZED columns.
            df['FLAG_EDS_LEMATIZED'] = df['FLAG'].fillna('') + df['EDS_LEMATIZED'].fillna('')
            
            print(f"{df.name}.shape after adding FLAG_EDS_LEMATIZED column:  {df.shape}")
            display(df[['FLAG', 'EDS_ORIGINAL', 'EDS_LEMATIZED', 'FLAG_EDS_LEMATIZED']].drop_duplicates(subset=['FLAG_EDS_LEMATIZED']).head(3))
            print(f"-" * 50)
    return dfs


def verify_nan_handling_and_unique_flags(dfs):
    """
    Verify NaN handling in FLAG column and display unique FLAG_EDS_LEMATIZED values
    
    Args:
        df1: First dataframe to check
        df2: Second dataframe to check 
        df1_name: Name for first dataframe in print statements
        df2_name: Name for second dataframe in print statements
    """
    cols = ['FLAG', 'EDS_ORIGINAL', 'EDS_LEMATIZED', 'FLAG_EDS_LEMATIZED']
    for df in dfs:
        if 'EDS' in df.columns:
            print(f"Verification of {df.supplier} - {df.name}:")
            print("NaN handling in the FLAG column:")
            display(df[cols][df['FLAG'].isna()].drop_duplicates(subset=['FLAG_EDS_LEMATIZED']).head(3))

            print(f"Unique values of FLAG_EDS_LEMATIZED:")
            display(df['FLAG_EDS_LEMATIZED'].unique())
            print(f"-" * 50)


def find_flag_in_eds(eds_text, flags):
    """
    Searches for the first flag from the list 'flags' within the 'eds_text'.
    Performs a case-insensitive search.
    Returns the matching flag or None if no flag is found or eds_text is NaN.
    """
    if pd.isna(eds_text):
        return None
    eds_upper = str(eds_text).upper()
    for flag in flags:
        if flag.upper() in eds_upper:
            return flag  # Return the first flag found  
    return None # Return None if no flag matches


def replace_eds_lematized(dfs, flag_eds_lematized_replacements, column_to_replace='FLAG_EDS_LEMATIZED'):
    """
    Make discretionary changes to the EDS_LEMATIZED column in the dataframe.
    Perform case-insensitive string replacements in a DataFrame column using a mapping dictionary.
    Efficiently applies multiple string replacements to a specified column across one or more DataFrames.
    The function handles both single DataFrame and list of DataFrames inputs.
    Args:
        dfs (pd.DataFrame or list[pd.DataFrame]): Input DataFrame(s) to modify
        flag_eds_lematized_replacements (dict): Replacement mapping {old_value: new_value} where:
            - keys should be uppercase strings to match against
            - values will be the replacement values
        column_to_replace (str, optional): Column name to modify. Defaults to 'FLAG_EDS_LEMATIZED'.
    Returns:
        list[pd.DataFrame]: List of modified DataFrames (even if single input was provided)
    Note:
        - Performs case-insensitive matching by converting column values to uppercase
        - Original values are preserved where no replacement mapping exists
        - Modifies DataFrames in-place but also returns them for chaining
    """
    if not isinstance(dfs, list):
        dfs = [dfs]

    for df in dfs:
        if column_to_replace in df.columns:
            # Map the replacements. `map` is generally faster than a loop.
            # It looks for keys in the uppercase version of the column.
            # Values that are not found in the map become NaN.
            new_values = df[column_to_replace].str.upper().map(flag_eds_lematized_replacements)

            # We use `fillna` to keep the original value where no replacement occurred.
            df[column_to_replace] = new_values.fillna(df[column_to_replace])
    return dfs


def find_match_nomatch_serv_st(dfs, unique_serv_st_lematized, columns_to_return, col_to_compare='FLAG_EDS_LEMATIZED'):
    """
    Compare service station values between input DataFrames and a reference list.
    Identifies and displays matching/non-matching service stations with statistics.
    Args:
        dfs (pd.DataFrame or list[pd.DataFrame]): Input DataFrame(s) to analyze
        unique_serv_st_lematized (list): Reference list of service station values to compare against
        columns_to_return (list): Columns to include in output samples
        col_to_compare (str, optional): Column containing values to compare. Defaults to 'FLAG_EDS_LEMATIZED'.
    Returns:
        None: Prints comparison results including:
            - Sample matching records (3 rows)
            - Count of unique matches
            - Sample non-matching records (if any)
            - Count of unique non-matches
    Raises:
        ValueError: If input types are invalid (columns_to_return not list,
                   col_to_compare not str, unique_serv_st_lematized not list)
    """
    if not isinstance(columns_to_return, list):
        return ValueError("columns_to_return must be a list")
    if not isinstance(col_to_compare, str):
        return ValueError("col_to_compare must be a string")
    if not isinstance(unique_serv_st_lematized, list):
        return ValueError("unique_serv_st_lematized must be a list")
    
    # Add col_to_compare (in case it is not in the columns_to_return list) while preserving original order 
    # and ensuring uniqueness
    if col_to_compare not in columns_to_return:
        columns_to_return = [col for col in columns_to_return + [col_to_compare] 
                           if col not in columns_to_return or col == col_to_compare]

    for df in dfs:
        if all(col in df.columns for col in columns_to_return):
            # if all(col in df.columns for col in columns_required):
            # Find matches
            df_matches = df[columns_to_return][df[col_to_compare].isin(unique_serv_st_lematized)].reset_index(drop=True)
            
            # Find no matches
            df_nomatches = df[columns_to_return][~df[col_to_compare].isin(unique_serv_st_lematized)].reset_index(drop=True)
            
            print(f"\n{df.name} - SERVICE STATIONS:")
            print(f"Sample matches (3 records):")
            display(df_matches.drop_duplicates(subset=[col_to_compare]).head(3))
            print(f"Matches: {df_matches.drop_duplicates(subset=[col_to_compare]).shape[0]}")

            nomatch_q = df_nomatches.drop_duplicates(subset=[col_to_compare])
            if nomatch_q.shape[0] > 0:
                print(f"\n No-Matches Sample:")
                display(nomatch_q)
            print(f"No-Matches: {nomatch_q.shape[0]}")
            print(f"--------------------------------")


def check_n_fix_dfs_first_empty_row(dfs):
    """
    Validate and fix DataFrame headers and first row issues.
    Checks for:
    1. Malformed headers (majority 'Unnamed' columns)
    2. Invalid first data row (mostly non-object types)
    Args:
        dfs: DataFrame or list of DataFrames to validate
    Returns:
        Validated DataFrame(s) with fixed headers
    Raises:
        ValueError: If empty DataFrame, bad headers, or invalid first row
    """
    is_list = True
    if not isinstance(dfs, list):
        is_list = False
        dfs = [dfs]

    for i, df in enumerate(dfs):
        name = getattr(df, 'name', None)  # Use current df's name attribute if exists
        supplier = getattr(df, 'supplier', None)  # Preserve supplier attribute if exists
        price_column = getattr(df, 'price_column', None)  # Preserve price_column attribute if exists

        if df.empty:
            raise ValueError(f"Cannot validate an empty DataFrame: {name if name else 'no-name'}")

        # 1. Validate headers for "Unnamed" columns
        print(f"Validating headers for {name if name else 'dataframe'}")
        fixed_df = fix_unnamed_headers(df)
        
        # Restore attributes that may have been lost
        if supplier:
            fixed_df.supplier = supplier
        if name:
            fixed_df.name = name
        if price_column:
            fixed_df.price_column = price_column

        dfs[i] = fixed_df

        # 2. Validate the first data row appears to be a valid header
        object_ratio = df.iloc[0].apply(lambda x: isinstance(x, object)).sum() / len(df.columns)
        if object_ratio < 0.7:
             raise ValueError(
                f"The first data row of '{name}' does not appear to contain valid headers. "
                f"Expected over 70% of values to be object type, but found only {object_ratio:.0%}. "
                "This may indicate that the file's header is missing or on a different row."
            )
        print(f"Validation passed: Headers are valid and the first data row has valid structure.\n")
    
    if is_list:
        return dfs
    else:
        return dfs[0]


def handle_accents_n_other_characters_col_content(df, force_to_string_types):
    """
    Normalize string columns (handle accents and ñ characters):
    1. Converting specified columns to string type (preserving NA values)
    2. Applying text standardization (strip, uppercase, space replacement)
    3. Converting Unicode characters to ASCII equivalents (NFKD normalization)
    Args:
        df: DataFrame to process
        force_to_string_types: List of columns to forcibly convert to string type
    Returns:
        DataFrame with processed string columns
    Raises:
        AttributeError: If non-string data is found in object-type columns
    """
    string_cols = df.select_dtypes(include='object').columns.tolist()
    # Append the force_to_string_types to the string_cols list
    string_cols.extend(force_to_string_types)

    print(f"Converting the following columns to string type: {force_to_string_types}")
    
    # Force the column to string if it is in the force_to_string_types list
    for col in string_cols:
        if col in df.columns:
            is_na = df[col].isna()
            df[col] = df[col].astype(str)
            df.loc[is_na, col] = pd.NA  # Restore NA values after conversion
            
            # Process the TRANSPORTE and ENTREGA columns (similar treatment)
            if col in {'TRANSPORTE', 
                       'ENTREGA', 
                       'MATERIAL'}:
                df[col] = df[col].astype(float).astype(int).astype(str)
            
            # Process the CREADO_EL column
            elif 'CREADO_EL' == col:
                # `errors='coerce'` will turn any unparseable dates into NaT (Not a Time).
                # `dayfirst=False` ensures 'mm/dd/yyyy' is parsed correctly.
                df[col] = pd.to_datetime(df[col], dayfirst=False, errors='coerce').astype(str)

            # Process string columns
            try:
                df[col] = (df[col].str.strip()
                                  .str.upper()
                                  .str.replace(r'\s+', '_', regex=True)
                                  .str.normalize('NFKD')
                                  .str.encode('ascii', errors='ignore')
                                  .str.decode('ascii'))
            except AttributeError:  # Only catch AttributeError specifically
                print(f"\n--- ERROR: Could not process column '{col}' with string operations. ---")
                print("This column has an 'object' dtype but likely contains non-string data.")
                print(f"Value types found in column '{col}':")
                print("!!!Consider adding this column to the 'force_to_string_types' list!!!.")
                print(df[col].apply(type).value_counts())
                print("---------------------------------------------------------------------\n")
                raise
    
    return df


def handle_strings_to_floats_col_content(df, force_to_float_types):
    """
    Handle strings to floats in the dataframe
    """
    print(f"Converting the following columns to float type: {force_to_float_types}")
    
    for col in force_to_float_types:
        if col in df.columns:
            if pd.api.types.is_numeric_dtype(df[col]):
                if df[col].dtype != 'float64':
                    print(f"Column '{col}' is numeric ({df[col].dtype}), converting to float64.")
                    df[col] = df[col].astype('float64')
                else:
                    print(f"Skipping column '{col}' as it is already float64.")
                continue  # Finished with this column, move to the next one
            
            # To perform string operations, we must handle non-string data.
            # Store original NaNs and then convert the column to string.
            is_na = df[col].isna()
            series = df[col].astype(str)
            series.loc[is_na] = pd.NA  # Restore NA values

            # Clean the string values before converting to numeric.
            ## This regex removes '$', ',' and all whitespace characters.
            ## e.g., '$ 1,234.56 ' -> '1234.56' and ' - 53 ' -> '-53'
            cleaned_series = series.str.replace(r'[$,]|\s+', '', regex=True)
            
            # Replace placeholder hyphens/en-dashes with NaN.
            ## The regex is anchored with `^` and `$` to ensure it only matches cells
            ## that contain *nothing but* a hyphen or en-dash (and optional whitespace).
            ## This is safe and will NOT affect negative numbers like '-53'.
            cleaned_series = cleaned_series.replace(r'^\s*(-|–+)\s*$', pd.NA, regex=True)

            # Convert to numeric, setting any values that fail to parse to NaN.
            df[col] = pd.to_numeric(cleaned_series, errors='coerce').astype('float64')
        else:
            print(f"Warning: Column '{col}' from force_to_float_types not found in DataFrame.")

    return df


def uppercase_df_columns_manage_content(df):
    """
    Standardize dataframe structure and content by:
    1. Cleaning column names (strip whitespace, remove unnamed/empty cols)
    2. Removing completely empty columns/rows
    3. Normalizing string format (uppercase, underscores for spaces)
    4. Applying consistent formatting to both headers and string content
    Args:
        df: DataFrame to process
    Returns:
        pd.DataFrame: Processed dataframe with standardized structure and formatting
    """
    # Strip the columns names of the dataframe
    df.columns = df.columns.str.strip()

    # Delete unnamed columns
    print(f"\nShape before potentially dropping columns: {df.shape}")
    cols_to_drop = []
    for col in df.columns:
        # Check if column is unnamed (matches 'Unnamed: X' pattern or is empty string)
        is_unnamed = bool(re.match(r'^Unnamed: \d+$', str(col))) or str(col) == ''
        # If condition are true, add to the drop list
        if is_unnamed:
            cols_to_drop.append(col)
    ## Drop these unnamed/empty columns if any were found
    if cols_to_drop:
        df = df.drop(columns=cols_to_drop)
        print(f"Dropped {len(cols_to_drop)} unnamed and empty columns: {cols_to_drop}")

    # Delete empty columns and rows
    df = df.dropna(axis=1, how='all')  # Drop columns with all NaN values
    df = df.dropna(axis=0, how='all')  # Drop rows with all NaN values
    print(f"Shape after dropping empty columns and rows: {df.shape}")

    # Uppercase the column names and replace spaces, en-dash (longer than a regular dash) and dashwith underscores
    df.columns = df.columns.str.strip().str.upper().str.replace(r'[\s–-]+', '_', regex=True)

    # Replace accents and ñ with their non-accented equivalents
    df.columns = df.columns.str.normalize('NFKD').str.encode('ascii', errors='ignore').str.decode('ascii')

    # Manage Content
    ## Handle accents and ñ in the row content
    ### Define a list of columns that sometimes can have types that are objects (datetime.time) 
    ### but should be forced to strings
    force_to_string_types = [  # This list can increase as needed
        "TRANSPORTE",
        "ENTREGA",
        "HORA_PLAN_INI", 
        "HORA_PLAN_FIN",
        "HORA_REAL_INI",
        "HORA_REAL_FIN",
        "CREADO_EL",
        "G",
        "MATERIAL",
    ]
    df = handle_accents_n_other_characters_col_content(df, force_to_string_types)
    
    ## Handle strings to numbers in the dataframe
    force_to_float_types = [
        "ADT",
        "CALCULO_WPS",
        "CANTIDAD_ENTREGA",
        "DIF_USD_RED_ID",
        "DISTANCIA",
        "LIQ_RT",
        "N_HL",
        "N_KM",
        "N_PALLET",
        "PALLET",
        "PESO_BRUTO",
        "PESO_NETO",
        "RT_BGT",
        "TARIFA_BASE_SIDER",
        "TARIFA_BGT",
        "TARIFA_DESVIO",
        "TARIFA_ESTACA",
        "TARIFA_FINAL_REAL",
        "V/_NETO",
        "VIAJES",
        "VIAJES_VIRTUALES",
        "VOLUMEN",
    ]
    df = handle_strings_to_floats_col_content(df, force_to_float_types)

    return df


def preprocess_adt_dataframe(dfs):
    """
    Standardizes and cleans one or more ADT (transport) dataframes for consistent analysis.
    Applies comprehensive preprocessing including:
    - Column name standardization (uppercase, underscores)
    - Removal of empty/unused columns and rows  
    - String content normalization (uppercase, underscores)
    - Special handling of TRANSPORTE column (int to str conversion)

    Handles both single dataframe and list of dataframes input (to ensure backward compatibility):
    - If input is a single dataframe: converts to list internally and returns single dataframe
    - If input is a list: processes all dataframes and returns list
    
    Preserves the original dataframe name attribute throughout processing to maintain
    traceability and identification of data sources in downstream operations.
    
    Args:
        dfs: Raw ADT dataframe or list of dataframes containing transport route data
    Returns:
        pd.DataFrame or list[pd.DataFrame]: Returns same type as input - single dataframe if input was single,
                                            list of dataframes if input was list
    """
    # Validate the dfs is a list
    is_list = True
    if not isinstance(dfs, list):
        dfs = [dfs]
        is_list = False

    # Process each df in the list by index to ensure modifications persist.
    # A 'for df in dfs:' loop would only modify a temporary reference, not the list itself.
    for i in range(len(dfs)):
        # Preserve the original name before processing
        original_name = getattr(dfs[i], 'name', None)
        
        # Reassign the processed dataframe back to the list at the correct index
        processed_df = uppercase_df_columns_manage_content(dfs[i])

        # Restore the original name to the new, processed dataframe
        if original_name:
            processed_df.name = original_name
        
        dfs[i] = processed_df

    # Return the processed dataframe or list of dataframes (to ensure backward compatibility)
    ## If the dfs is a list, return the list
    if is_list:
        return dfs
    ## If the dfs is not a list, return the single dataframe
    else:
        return dfs[0]


def validate_adts_column_types(dfs):
    """
    Validates that common columns across a list of DataFrames have the same data type.
    It prints each error found and raises a single ValueError at the end if any
    inconsistencies were found.
    Args:
        dfs: A list of pandas DataFrames to compare.
    Returns:
        Nothing. Raises a ValueError if a type inconsistency is found.
    """
    if not isinstance(dfs, list) or len(dfs) < 2:
        return  # No comparison needed for a single DataFrame or non-list input

    # Find the intersection of columns common to all DataFrames
    common_cols = set(dfs[0].columns)
    for df in dfs[1:]:
        common_cols.intersection_update(df.columns)

    print(f"\n--- DataFrame Validation Report ---")
    print(f"Found {len(common_cols)} common columns to validate.")

    mismatch_errors = []
    # Check for type inconsistencies in common columns
    for col in sorted(list(common_cols)):
        # Get the dtype from the first DataFrame to use as a reference
        first_dtype = dfs[0][col].dtype
        
        # Check if all other DataFrames have the same dtype for this column
        if not all(df[col].dtype == first_dtype for df in dfs[1:]):
            # Create an informative error message using the df.name attribute
            mismatched_types = {
                getattr(df, 'name', f"df_{i}"): str(df[col].dtype)
                for i, df in enumerate(dfs)
            }
            error_msg = f"Column '{col}' has mismatched data types: {mismatched_types}"
            print(f"ERROR: {error_msg}")
            mismatch_errors.append(error_msg)

    if mismatch_errors:
        all_errors = "\n".join(mismatch_errors)
        raise ValueError(f"Found {len(mismatch_errors)} type inconsistencies across DataFrames:\n{all_errors}")
    else:
        print("All common columns have consistent data types.")

    print(f"--- End of Report ---\n")


def validate_unique_transport_records(df_adts_list):
    """
    Validates that records are unique across dataframes, by comparing the TRANSPORTE column.
    
    Checks if 'TRANSPORTE' column values in each dataframe are distinct from other dataframes.
    Calculates intersection ratio between dataframes and raises error if overlap exceeds 1%.
    Args:
        df_adts_list: List of dataframes with 'TRANSPORTE' column and 'name' attribute
    Raises:
        ValueError: If shared transport records exceed 1% threshold
    """
    if len(df_adts_list) < 2:
        print("Skipping TRANSPORTE record validation because it has only one df. \
              It requires at least two dataframes to compare.")
        return

    eval_transport = {}
    for df in df_adts_list:
        eval_transport[df.name] = set(df['TRANSPORTE'])

    for df in df_adts_list:
        # Merge all the eval_transport dicts into a single set, except the current df
        other_dfs_transport = set()
        for other_df_name, transport_set in eval_transport.items():
            if other_df_name != df.name:
                other_dfs_transport.update(transport_set)
        
        # Find the intersection between the current df and all other dfs (other_dfs_transport)
        unique_transport = eval_transport[df.name]
        shared_transport = eval_transport[df.name].intersection(other_dfs_transport)
        shared_ratio = len(shared_transport) / len(unique_transport)
        
        # Print the results
        print(f"{df.name}. Shape: {df.shape}. Has {len(unique_transport)} unique TRANSPORTE records")
        if shared_ratio > 0:
            print(f"Shares {len(shared_transport)} TRANSPORTE records with other dataframes")
            if shared_transport:
                print(f"List of common records: {list(shared_transport)}")
            print(f"Shared ratio with other dataframes: {shared_ratio:.4f}")
        else:
            print(f"No shared transport records with other dataframes")
        print("-" * 70)

        # If the shared ratio is greater than 1%, raise an error
        if shared_ratio > 0.01:
            raise ValueError(f"The dataframe {df.name} has {shared_ratio:.4f} of shared TRANSPORTE records with other dataframes")


def find_extra_columns(df_list):
    """
    Finds all columns that are not common to every DataFrame in a list.

    This function compares the column sets of all provided DataFrames and
    identifies any columns that do not appear in every single one. This is
    useful for debugging before a merge or concatenation operation.
    Args:
        df_list: A list of pandas DataFrames.
    Returns:
        Prints the non-common columns found.
    """
    if not df_list or len(df_list) < 2:
        print("Requires at least two DataFrames to compare.")
        return set()
    # Generate sets of columns for all dataframes (one per dataframe)
    all_columns_sets = [set(df.columns) for df in df_list]
    # Find columns that are common to all dataframes
    common_columns = set.intersection(*all_columns_sets)
    # Find all columns that appear in any of the dataframes
    all_columns_union = set.union(*all_columns_sets)
    # The non-common columns are the ones in the union but not in the intersection
    non_common_columns = all_columns_union - common_columns
    if non_common_columns:
        print(f"Columns not shared across all DataFrames: {non_common_columns}")
    else:
        print("All DataFrames have the same columns.")


def include_adt_object_name_column(dfs):
    """Adds ADT object name as a column to one or more DataFrames.
    Args:
        dfs: A single pandas DataFrame or list of DataFrames. Each must have a 
             'name' attribute containing the object name to store.
    Returns:
        Same type as input: Returns processed DataFrame(s) with new col 'ADT_OBJECT_NAME'.
    """
    # Validate the dfs is a list
    is_list = True
    if not isinstance(dfs, list):
        dfs = [dfs]
        is_list = False
    
    for df in dfs:
        print(f"Shape before adding ADT_OBJECT_NAME for {df.name}: {df.shape}")
        df['ADT_OBJECT_NAME'] = df.name
        print(f"Shape after adding ADT_OBJECT_NAME for  {df.name}: {df.shape}\n")
    
    if is_list:
        return dfs
    else:
        return dfs[0]


def consolidate_adt_dataframe(df, additional_cols_to_sum=[]):
    """
    Consolidates ADT data by 'TRANSPORTE'
    """
    if 'TRANSPORTE' not in df.columns:
        raise ValueError("Input DataFrame must have a 'TRANSPORTE' column.")
    
    df_copy = df.copy()  # Make a copy of the dataframe to avoid modifying the original
    
    cols_to_sum = ['N_KM', 'VIAJES_VIRTUALES']
    
    # Validate the cols_to_sum is a list
    if not isinstance(additional_cols_to_sum, list):
        additional_cols_to_sum = [additional_cols_to_sum]

    # Find columns starting with 'TARIFA_'
    tarifa_cols = [col for col in df_copy.columns if isinstance(col, str) and col.startswith('TARIFA_')]
    
    # Combine user-provided columns with tarifa columns, ensuring no duplicates
    all_cols_to_sum = list(set(additional_cols_to_sum + cols_to_sum + tarifa_cols))

    # Build the aggregation dictionary
    agg_dict = {}
    for col in df_copy.columns:
        if col == 'TRANSPORTE':
            continue
        
        if col in all_cols_to_sum:
            agg_dict[col] = 'sum'
        else:
            agg_dict[col] = 'first'
            
    # Group by 'TRANSPORTE' and aggregate
    consolidated_df = df_copy.groupby('TRANSPORTE').agg(agg_dict).reset_index()

    # Drop rows where 'VIAJES_VIRTUALES' is less than 0.99
    consolidated_df = consolidated_df[consolidated_df['VIAJES_VIRTUALES'] >= 0.99]
    
    # Reorder columns to be similar to original, with TRANSPORTE at the beginning
    original_order_cols = [col for col in df_copy.columns if col != 'TRANSPORTE']
    final_cols = ['TRANSPORTE'] + [col for col in original_order_cols if col in consolidated_df.columns]
    consolidated_df = consolidated_df[final_cols]

    return consolidated_df


def export_llm_demo_data(dfs_list, period_demo, export_route_demo=None):
    """
    Export the demo data to a csv file
    
    Args:
        dfs_list: List of DataFrames to search through
        period_demo: Period identifier to match in DataFrame names
        export_route_demo: Path to export CSV file (None will skip export)
    """
    if not dfs_list:
        print("Warning: Empty DataFrame list provided")
        return
        
    ending_period_demo = "_" + str(period_demo)
    for df in dfs_list:
        if not hasattr(df, 'name'):
            print(f"Warning: DataFrame missing 'name' attribute, skipping")
            continue
            
        if df.name.endswith(ending_period_demo):
            if export_route_demo is None:
                print("Warning: No export path provided, skipping export")
                return
            to_export = consolidate_adt_dataframe(df)
            select_columns = ['TRANSPORTE', 'ET_FINAL', 'TARIFA_BASE_SIDER']
            columns_to_exclude = ['N_HL', 'N_PALLET']
            select_columns.extend([c for c in to_export if c.startswith('N_') if c not in columns_to_exclude])
            to_export = to_export[select_columns]
            to_export.to_csv(export_route_demo, index=False)


def export_top_frequent_routes_adt(df, dfs_size, export_route=None, additional_cols_to_sum=[]):
    """
    Aggregates transport routes data by 'TRANSPORTE' and calculates route statistics.
    Processes DataFrame to:
    - Group by 'TRANSPORTE' with sum/first aggregations
    - Calculate route modes and total kilometers
    - Filter and export results
    Args:
        df: Input DataFrame with transport data (must contain 'TRANSPORTE')
        export_route: Path to export results (None skips export)
        additional_cols_to_sum: Extra columns to sum during aggregation
    Returns:
        None (exports to Excel and displays results)
    """
    if export_route is None:
        print("NO EXPORT ROUTE has been selected (None).")
        print("Skipping top frequent routes ADT function. Please select an export route.")
        return
    
    # Get the consolidated dataframe
    consolidated_df = consolidate_adt_dataframe(df)
    
    # Create a matching key by N_CENTRO_ORIGEN and N_CENTRO_DESTINO
    consolidated_df['RUTA_KEY'] = consolidated_df['N_CENTRO_ORIGEN'] + '_' + consolidated_df['N_CENTRO_DESTINO']

    # Add a column with the count of matches by RUTA_KEY column
    consolidated_df['COUNT_RUTA'] = consolidated_df.groupby('RUTA_KEY')['RUTA_KEY'].transform('count')

    # Add a column with the mode of the N_KM given the RUTA column
    consolidated_df['N_KM_MODE'] = consolidated_df.groupby('RUTA_KEY')['N_KM'].transform(lambda x: x.mode()[0] if not x.mode().empty else np.nan)
    # Add a column with the mode of the RUTA given the RUTA_KEY column
    consolidated_df['RUTA_MODE'] = consolidated_df.groupby('RUTA_KEY')['RUTA'].transform(lambda x: x.mode()[0] if not x.mode().empty else np.nan)
    # Add a column with the mode of the INDENT_RUTA given the RUTA_KEY column
    consolidated_df['INDENT_RUTA_MODE'] = consolidated_df.groupby('RUTA_KEY')['INDENT_RUTA'].transform(lambda x: x.mode()[0] if not x.mode().empty else np.nan)

    # Create a total km per route of the COUNT_RUTA by the N_KM_MODE column
    consolidated_df['TOT_KM_RUTA'] = consolidated_df['COUNT_RUTA'] * consolidated_df['N_KM_MODE']

    # Create a total paid to the supplier per route key (RUTA_KEY). This is an annualized value
    consolidated_df['TOT_PAID_RUTA_KEY_ANNUALIZED'] = consolidated_df.groupby('RUTA_KEY')['TARIFA_BASE_SIDER'].transform('sum') / (dfs_size / 12)

    # Select the columns to keep
    consolidated_df = consolidated_df[['RUTA_MODE', 'RUTA_KEY', 'INDENT_RUTA_MODE', 
                                       'N_CENTRO_ORIGEN', 'N_CENTRO_DESTINO', 'N_KM_MODE', 
                                       'COUNT_RUTA', 'TOT_KM_RUTA', 'TOT_PAID_RUTA_KEY_ANNUALIZED']]

    # Organize in descending order by TOT_KM_RUTA
    consolidated_df = consolidated_df.sort_values(by='TOT_KM_RUTA', ascending=False)

    # Drop duplicates by RUTA_KEY and keep first match
    consolidated_df = consolidated_df.drop_duplicates(subset=['RUTA_KEY'], keep='first')
    
    # Drop the RUTA_KEY column and drop rows where TOT_KM_RUTA is 0
    consolidated_df = consolidated_df.drop(columns=['RUTA_KEY'])
    consolidated_df = consolidated_df[consolidated_df['TOT_KM_RUTA'] > 0]
    
    consolidated_df = consolidated_df.reset_index(drop=True)

    # Export the consolidated_df to excel
    consolidated_df.to_excel(export_route, index=False)

    print(f"Sum of TOT_PAID_RUTA_KEY_ANNUALIZED: {consolidated_df['TOT_PAID_RUTA_KEY_ANNUALIZED'].sum()}")
    display(consolidated_df)
    

def extract_origin_destination_centers(origin_destination_dfs, orig_dest_dict):
    """
    Consolidates origin/destination centers from multiple dataframes into standardized format.
    Processes each dataframe to extract and rename origin/destination columns, then combines all results.
    Normalizes column names to 'N_CENTRO' and 'CODIGO'.
    Args:
        origin_destination_dfs: DataFrame or list of DataFrames containing transport route data
        orig_dest_dict: Dict with keys:
            - origin_cols: [origin_name_col, origin_code_col]
            - destination_cols: [dest_name_col, dest_code_col]
    Returns:
        DataFrame with standardized columns ['N_CENTRO', 'CODIGO'] containing all unique centers
    """
    # Convert single dataframe to list if needed
    if not isinstance(origin_destination_dfs, list):
        origin_destination_dfs = [origin_destination_dfs]

    all_dfs = []
    for df in origin_destination_dfs:
        # Process origin columns
        df_origin = df[orig_dest_dict['origin_cols']].rename(
            columns={orig_dest_dict['origin_cols'][0]: 'N_CENTRO',
                     orig_dest_dict['origin_cols'][1]: 'CODIGO'})
        # Process destination columns 
        df_dest = df[orig_dest_dict['destination_cols']].rename(
            columns={orig_dest_dict['destination_cols'][0]: 'N_CENTRO',
                     orig_dest_dict['destination_cols'][1]: 'CODIGO'})
            
        all_dfs.extend([df_origin, df_dest])
    
    return pd.concat(all_dfs, ignore_index=True)


def get_center_mode(df, center_col='N_CENTRO', code_col='CODIGO'):
    """
    Creates a standardized mapping from center names to their most frequent codes (mode).
    Args:
        df: DataFrame containing center and code columns
        center_col: Name of column containing center names
        code_col: Name of column containing center codes
        
    Returns:
        DataFrame with unique center names and their most frequent codes,
        with all values in uppercase and duplicates removed.
    """
    # Create mapping from center (N_CENTRO) to most frequent code -mode- (CODIGO)
    center_code_mode = df.groupby(center_col)[code_col].apply(lambda x: x.mode()[0] if not x.mode().empty else None)
    
    # Add mode column and convert to uppercase
    df['CODIGO_MODE'] = df[center_col].map(center_code_mode)
    df[center_col] = df[center_col].str.upper()
    df['CODIGO_MODE'] = df['CODIGO_MODE'].str.upper()
    
    # Return unique centers with their mode codes
    return (df[['N_CENTRO', 'CODIGO_MODE']]
            .drop_duplicates(subset=['N_CENTRO'])
            .sort_values('N_CENTRO')
            .reset_index(drop=True))


def set_merge_key(df, name_msg, conversions_table, cargo_type_map, origin_col, destination_col, cargo_type_col, conversion_col, merge_col):
    """
    Creates a standardized merge key by mapping origin, destination and cargo type to their respective codes.
    
    Args:
        df: DataFrame to process (e.g. df_cga_qtn_nodetr_2505, df_trans_qtn_nodetr_2505, df_adt_2505)
        name_msg: Name identifier for logging (e.g. 'CARGA', 'TRANSER', 'ADT')
        conversions_table: Mapping table for location codes (e.g. conversions_table)
        cargo_type_map: Mapping table for cargo types (e.g. cargo_type_map)
        origin_col: Column name for origin locations (e.g. 'ORIGEN', 'N_CENTRO_ORIGEN')
        destination_col: Column name for destination locations (e.g. 'DESTINO', 'N_CENTRO_DESTINO')
        cargo_type_col: Column name for cargo types (e.g. 'TIPO', 'N_TIPO_DT')
        conversion_col: Column in conversions_table to use for lookup (e.g. 'N_CENTRO_SUPPLIERS', 'N_CENTRO_ADT')
        merge_col: Name of new column to create with merged key (e.g. 'KEY_TO_ADT', 'KEY_TO_SUPPLIER')
        
    Returns:
        DataFrame with new merge key column and validation warnings for missing mappings
    """
    # Make the columns content uppercase (This is redundant bc the columns were already converted to uppercase. Keep redundancy)
    df[origin_col] = df[origin_col].str.upper()
    df[destination_col] = df[destination_col].str.upper()
    df[cargo_type_col] = df[cargo_type_col].str.upper()

    location_map = pd.Series(conversions_table['CODIGO_MODE'].values, index=conversions_table[conversion_col]).to_dict()
    cargo_type_dict = pd.Series(cargo_type_map['CARGO_TYPE'].values, index=cargo_type_map[cargo_type_col]).to_dict()

    # Map 'ORIGEN', 'DESTINO' and 'TIPO' to their codes using the map
    df[merge_col] = df[origin_col].map(location_map).astype(str) + '_' +\
                    df[destination_col].map(location_map).astype(str) + '_' + \
                    df[cargo_type_col].map(cargo_type_dict).astype(str)

    # Display the first few rows with the new key to verify
    print(f"Sample of the new key ({name_msg}):")
    display(df[[origin_col, destination_col, cargo_type_col, merge_col]].head())

    # Check for missing mappings
    missing_origins = df[df[origin_col].map(location_map).isna()][origin_col].unique()
    missing_destinations = df[df[destination_col].map(location_map).isna()][destination_col].unique()
    missing_types = df[df[cargo_type_col].map(cargo_type_dict).isna()][cargo_type_col].unique()
    if len(missing_origins) > 0:
        print(f"Warning {name_msg}: Could not find CODIGO_MODE for these {origin_col} locations in conversions_table[{conversion_col}]: {list(missing_origins)}")
    if len(missing_destinations) > 0:
        print(f"Warning {name_msg}: Could not find CODIGO_MODE for these {destination_col} locations in conversions_table[{conversion_col}]: {list(missing_destinations)}")
    if len(missing_types) > 0:
        print(f"Warning {name_msg}: Could not find CARGO_TYPE for these {cargo_type_col} types in cargo_type_map['{cargo_type_col}']: {list(missing_types)}")
    
    return df


def set_merge_key_qtn_dfs(dfs, qtn_common_params):
    """Set the merge key for each dataframe in the quotations list (dfs).
    Process a list of dataframes to set standardized merge keys for cross-dataset matching.
    Args:
        dfs: A single dataframe or list of dataframes to process. Each dataframe should have 
             'name' and 'supplier' attributes.
        qtn_common_params: Dictionary of common parameters to pass to set_merge_key function.
            Expected keys:
            - conversions_table: Table mapping location names to codes
            - cargo_type_map: Table mapping cargo type names to codes  
            - origin_col: Column name for origin locations
            - destination_col: Column name for destination locations
            - cargo_type_col: Column name for cargo types
            - conversion_col: Column in conversions_table to use for mapping
            - merge_col: Name of new column to store merge key
    Returns:
        The processed dataframe(s) with merge keys added. Returns same type as input (single df or list).
    Note:
        Prints processing information for each dataframe including name, supplier and generated name_msg.
    """
    is_list = True
    if not isinstance(dfs, list):
        is_list = False
        dfs = [dfs]

    for df in dfs:
        # Generate identifier message for logging
        name = df.name
        supplier = df.supplier
        name_msg = supplier + '_' + name.split('_')[-1]
        print(f"name: {name} | supplier: {supplier} | name_msg: {name_msg}")
        
        # Apply merge key standardization
        df = set_merge_key(df=df, name_msg=name_msg, **qtn_common_params)
        
        print(f"-" * 50)

    if is_list:
        return dfs
    else:
        return dfs[0]
    

def remove_ambiguous_transporte_to_key_records(df):
    """
    Removes rows from the dataframe where a value in 'TRANSPORTE' column is associated with more 
    than one unique value in 'KEY_TO_SUPPLIER' column.
    Returns the filtered dataframe.
    """
    # Group df by 'TRANSPORTE' and count the number of unique 'KEY_TO_SUPPLIER' values for each group.
    transporte_key_counts = df.groupby('TRANSPORTE')['KEY_TO_SUPPLIER'].nunique()
    # Identify TRANSPORTE values that are associated with more than one unique KEY_TO_SUPPLIER (ambiguous mapping).
    transporte_multiple_keys = transporte_key_counts[transporte_key_counts > 1]
    if not transporte_multiple_keys.empty:
        print("\nDeleted TRANSPORTE values with more than one KEY_TO_SUPPLIER:")
        print(transporte_multiple_keys)
    # If there are any ambiguous TRANSPORTE values, remove all corresponding rows from df.
    print(f"Dataframe shape before removing ambiguous TRANSPORTE values: {df.shape}")
    filtered_df = df[~df['TRANSPORTE'].isin(transporte_multiple_keys.index)]
    print(f"Dataframe shape after removing ambiguous TRANSPORTE values:  {filtered_df.shape}")
    print(f"Total number of TRANSPORTE values removed: {len(transporte_multiple_keys)}")
    return filtered_df


def aggregate_adt_by_transporte(df_adt):
    """
    Aggregate ADT records by 'TRANSPORTE', summing key numeric columns and taking the first value of the others.
    Cleans numeric columns if needed. Returns both the full aggregation and a version excluding detours.
    Args: df_adt (pd.DataFrame): Raw ADT data.
    Returns: dict: {'adt_aggregated_all_cols_incl_detour': DataFrame with all routes,
                    'adt_aggregated_all_cols_no_detour': DataFrame with detours excluded,
                    'adt_columns_to_sum': List of summed columns}
    """
    ## The columns that need to be added (numeric) are:
    adt_columns_to_sum = [
        'VIAJES_VIRTUALES', 
        'N_KM',
        'TARIFA_BASE_SIDER',
        'TARIFA_DESVIO',
        ]
    ## Check if the columns exist in the dataframe
    adt_columns_to_sum = [col for col in adt_columns_to_sum if col in df_adt.columns]

    ## All other columns that need to take the first record are:
    all_other_columns = [col for col in df_adt.columns if col not in adt_columns_to_sum]

    ## Transform to float the columns to be added which are object
    for col in adt_columns_to_sum:
        # Check if the column exists and if its dtype is 'object'
        if col in df_adt.columns and df_adt[col].dtype == 'object':
            # 1. Ensure string type and remove thousands separators (commas)
            cleaned_col = df_adt[col].astype(str).str.replace(',', '', regex=False)
            # 2. Replace placeholders like '-' with pd.NA
            cleaned_col = cleaned_col.replace('-', pd.NA) 
            # 3. Convert to numeric, coercing errors
            df_adt.loc[:, col] = pd.to_numeric(cleaned_col, errors='coerce')

    print(f"ADT before aggregation - shape: {df_adt.shape} | TARIFA_BASE_SIDER sum: {df_adt['TARIFA_BASE_SIDER'].sum()}")

    ## Define the aggregation logic
    aggregation_rules = {}
    # Sum the numeric columns
    for col in adt_columns_to_sum:
        aggregation_rules[col] = 'sum'
    # Take the first value for the categorical columns
    for col in all_other_columns:
        # Exclude the grouping key 'TRANSPORTE' from the aggregation dict 
        # so it becomes the index (TRANSPORTE, as index is required to group by)
        if col != 'TRANSPORTE': 
            aggregation_rules[col] = 'first'

    ## Merge the df_adt on the TRANSPORTE column and then make the TRANSPORTE another column (reset_index)
    adt_aggregated_all_cols_incl_detour = df_adt.groupby('TRANSPORTE').agg(aggregation_rules).reset_index()
    print(f"ADT after aggregation  - shape: {adt_aggregated_all_cols_incl_detour.shape}  | TARIFA_BASE_SIDER sum: {adt_aggregated_all_cols_incl_detour['TARIFA_BASE_SIDER'].sum()}")

    # Select only the columns with NO detour
    adt_aggregated_all_cols_no_detour = adt_aggregated_all_cols_incl_detour[adt_aggregated_all_cols_incl_detour['DESVIO?'] == 'NO']
    print(f"ADT aggregated excluding detour (routes with NO detour) shape: {adt_aggregated_all_cols_no_detour.shape} | TARIFA_BASE_SIDER sum: {adt_aggregated_all_cols_no_detour['TARIFA_BASE_SIDER'].sum()}")

    aggregated_by_transporte_dict = {
        'adt_aggregated_all_cols_incl_detour': adt_aggregated_all_cols_incl_detour,
        'adt_aggregated_all_cols_no_detour': adt_aggregated_all_cols_no_detour,
        'adt_columns_to_sum': adt_columns_to_sum
    }
    return aggregated_by_transporte_dict


def aggregate_adt_data(df_adt):
    """
    Summarize ADT data by 'TRANSPORTE', returning two DataFrames:
    one with all routes, one excluding detours. Keeps only key columns.
    Args: df_adt (pd.DataFrame): Raw ADT data.
    Returns: tuple: (DataFrame with all routes, DataFrame with detours excluded)
    """
    # Aggregate the df_adt by TRANSPORTE and get the df with all columns
    aggregate_adt_by_transporte_dict = aggregate_adt_by_transporte(df_adt)

    ## Retrieve the dfs (with detour and without detour) and the columns to sum
    adt_aggregated_all_cols_incl_detour = aggregate_adt_by_transporte_dict['adt_aggregated_all_cols_incl_detour']
    adt_aggregated_all_cols_no_detour = aggregate_adt_by_transporte_dict ['adt_aggregated_all_cols_no_detour']
    cols_to_sum = aggregate_adt_by_transporte_dict['adt_columns_to_sum']
    
    # The columns that need to take the first record (categorical) are:
    adt_columns_to_take_first = [
        'TRANSPORTE', 
        'N_TIPO_DT',
        'N_CENTRO_ORIGEN', 
        'N_CENTRO_DESTINO', 
        'DESVIO?', 
        'ET_FINAL',
        'KEY_TO_SUPPLIER'
        ]
    ## Check if the columns exist in the dataframe
    adt_columns_to_take_first = [col for col in adt_columns_to_take_first if col in df_adt.columns]

    # Define the columns to keep (cols_to_sum + adt_columns_to_take_first)
    cols_to_keep = cols_to_sum + adt_columns_to_take_first
    
    ## For each of the dfs, select only the columns to keep
    adt_aggregated_incl_detour = adt_aggregated_all_cols_incl_detour[cols_to_keep]
    adt_aggregated_no_detour = adt_aggregated_all_cols_no_detour[cols_to_keep]

    # Display info to check data types
    print(adt_aggregated_no_detour.info())
    
    return adt_aggregated_incl_detour, adt_aggregated_no_detour


def report_adt_to_quotations(df_adt, dfs_qtn_all_split, adt_name, supplier_name=""):
    """Compare ADT routes against supplier quotations to analyze service coverage.
    For each supplier quotation dataframe provided, calculates and prints statistics showing:
    - Total ADT routes available
    - Routes covered by the supplier (count and percentage)
    - Routes not covered by the supplier (count and percentage)
    Args:
        df_adt (pd.DataFrame): DataFrame containing ADT route data. Must include:
            - 'KEY_TO_SUPPLIER' column: Unique identifiers for ADT routes
        dfs_qtn_all_split (pd.DataFrame or list): Either:
            - Single supplier quotation DataFrame, or 
            - List of multiple supplier DataFrames
            Each DataFrame must contain:
            - 'KEY_TO_ADT' column: Matching identifiers for supplier routes
            - 'supplier' attribute: Supplier name (required for multiple DataFrames)
            - 'name' attribute: Quotation name (required for multiple DataFrames)
        adt_name (str): Descriptive name for the ADT dataset in reports
        supplier_name (str, optional): Default supplier name if not provided in DataFrame
            attributes. Defaults to empty string.
    Returns:
        None: Prints formatted coverage statistics to stdout
    Raises:
        ValueError: If multiple DataFrames are provided without required 'supplier' and 'name' 
            attributes
    """
    # Ensure dfs_qtn_all_split is a list to handle single or multiple dataframes
    is_list = True
    if not isinstance(dfs_qtn_all_split, list):
        is_list = False
        dfs_qtn_all_split = [dfs_qtn_all_split]

    # Get the unique keys from the ADT dataframe
    adt_unique_keys = df_adt['KEY_TO_SUPPLIER'].unique()

    # For each supplier quotation dataframe, compare the ADT routes against the supplier routes
    for df_qtn_supplier in dfs_qtn_all_split:
        supp_name = getattr(df_qtn_supplier, 'supplier', '')
        qtn_name = getattr(df_qtn_supplier, 'name', '')
        
        supplier_period_name = supplier_name
        
        # Validate that the quotation name and supplier name are provided for multiple quotations
        if is_list:
            if not supp_name:
                raise ValueError("Supplier name is required")
            if not qtn_name:
                raise ValueError("Quotation name is required")
            
            # Create a period name for the supplier and quotation
            supplier_period_name = supp_name + '_' + qtn_name.split('_')[-1]
        
        # Context of number of ADT routes present in supplier
        supplier_unique_keys = df_qtn_supplier['KEY_TO_ADT'].unique()
        adt_keys_in_supplier = [k for k in adt_unique_keys if k in supplier_unique_keys]
        adt_keys_not_in_supplier = [k for k in adt_unique_keys if k not in supplier_unique_keys]
        
        print(f"QUOTATION (object name): {qtn_name}")
        print(f"CONTEXT - Total number of {adt_name} routes: {len(adt_unique_keys)}")
        print(f"Number of ADT routes serviced by {supplier_period_name}:       {len(adt_keys_in_supplier)} ({len(adt_keys_in_supplier)/len(adt_unique_keys)*100:.2f}%)")
        print(f"Number of ADT routes not serviced by {supplier_period_name}:   {len(adt_keys_not_in_supplier)} ({len(adt_keys_not_in_supplier)/len(adt_unique_keys)*100:.2f}%)")
        print(f"-" * 50)


def add_key_supplier_counts(df):
    """Adds route frequency count columns to ADT dataframe by grouping.
    Calculates and adds two new columns:
    1. KEY_TO_SUPPLIER_COUNT - Count of occurrences per KEY_TO_SUPPLIER (route identifier)
    2. KEY_TO_SUPPLIER_COUNT_PER_ET_FINAL - Count of occurrences per ET_FINAL (final stage)
    Args:
        df (pd.DataFrame): Input DataFrame containing ADT routes without detours. Must contain:
            - KEY_TO_SUPPLIER: Route identifier column
            - ET_FINAL: Final stage column 
            - TARIFA_BASE_SIDER: Base fare column (used for validation)
    Returns:
        pd.DataFrame: Modified DataFrame with added count columns, preserving all original data.
    """
    print(f"ADT aggregated shape before KEY_TO_SUPPLIER_COUNT: {df.shape} | TARIFA_BASE_SIDER sum: {df['TARIFA_BASE_SIDER'].sum()}")
    
    df = df.assign(
        KEY_TO_SUPPLIER_COUNT=df.groupby('KEY_TO_SUPPLIER')['KEY_TO_SUPPLIER'].transform('count'),
        KEY_TO_SUPPL_COUNT_PER_ET_FINAL_PER_ADT_NAME=df.groupby(['KEY_TO_SUPPLIER', 'ET_FINAL', 'ADT_OBJECT_NAME'])['KEY_TO_SUPPLIER'].transform('count')
    )
    
    print(f"ADT aggregated shape after KEY_TO_SUPPLIER_COUNT: {df.shape} | TARIFA_BASE_SIDER sum: {df['TARIFA_BASE_SIDER'].sum()}")
    display(df.head(3))
    
    return df


def add_frequency_mapping(qtn_dfs, adt_aggregated):
    """
    Maps ADT route frequencies to supplier quotation DataFrames and calculates weighted kilometers.
    For each supplier route (KEY_TO_ADT), this function:
    1. Calculates the percentage frequency of each ADT route (KEY_TO_SUPPLIER)
       based on its occurrence in the aggregated ADT data
    2. Maps these percentage frequencies to matching supplier routes
    3. Calculates weighted kilometers by multiplying route frequency with route distance
    Args:
        qtn_dfs (DataFrame or list): Supplier quotation DataFrame(s) to process. 
            If single DataFrame, will be converted to list internally.
        adt_aggregated (DataFrame): Pre-aggregated ADT data containing:
            - KEY_TO_SUPPLIER: Route identifiers
            - KEY_TO_SUPPLIER_COUNT: Pre-calculated route counts
            - KEY_TO_SUPPLIER_COUNT_PER_ET_FINAL: Pre-calculated route counts per each ET_FINAL
            - As it is the complete ADT dataset, it is used as denominator for percentage calculations.
    Returns:
        Same type as input (DataFrame or list): Returns input structure with added columns:
            - KEY_MTCH_ADT_%_FREQ: Percentage frequency of each route (0-100)
            - TOT_KM_PER_ADT_FREQ: Weighted kilometers (KM_POR_TRAYECTO × KEY_MTCH_ADT_%_FREQ -frequency-)
            - KEY_TO_SUPPLIER_COUNT_PER_ET_FINAL: Count per ET_FINAL stage
    """
    is_list = True
    if not isinstance(qtn_dfs, list):
        is_list = False
        qtn_dfs = [qtn_dfs]
    
    # Create a mapping dictionary from ADT data. It has a KEY_TO_SUPPLIER: COUNT (Percentage) mapping
    key_count_mapping = (adt_aggregated.set_index('KEY_TO_SUPPLIER')['KEY_TO_SUPPLIER_COUNT'] / len(adt_aggregated) * 100).to_dict()

    # Create a mapping for KEY_TO_SUPPLIER_COUNT_PER_ET_FINAL. Using a multi-index Series
    count_per_et_final_mapping = adt_aggregated.drop_duplicates(
                    subset=['KEY_TO_SUPPLIER', 'ET_FINAL']
                    ).set_index(['KEY_TO_SUPPLIER', 'ET_FINAL'])['KEY_TO_SUPPLIER_COUNT_PER_ET_FINAL']

    for qtn_df in qtn_dfs:
        # Include Frequency to supplier quotations
        qtn_df['KEY_MTCH_ADT_%_FREQ'] = qtn_df['KEY_TO_ADT'].map(key_count_mapping)
        # Include KM per route x frequency in ADT
        qtn_df['TOT_KM_PER_ADT_FREQ'] = qtn_df['KEY_MTCH_ADT_%_FREQ'] * qtn_df['KM_POR_TRAYECTO']

        # Create tuples of (KEY_TO_ADT, supplier) for mapping
        map_keys = list(zip(qtn_df['KEY_TO_ADT'], [qtn_df.supplier] * len(qtn_df)))
        # Map the count per ET_FINAL using the new mapping
        qtn_df['KEY_TO_SUPPLIER_COUNT_PER_ET_FINAL'] = pd.Series(map_keys, index=qtn_df.index).map(count_per_et_final_mapping)
        # Include GLS per route per ET_FINAL
        qtn_df['GLS_PER_ROUTE_PER_ET_FINAL'] = qtn_df['KEY_TO_SUPPLIER_COUNT_PER_ET_FINAL'] * qtn_df['GALONES_/_RUTA']

    if is_list:
        return qtn_dfs
    else:
        return qtn_dfs[0]


def merge_quotations_with_fuel_prices(quotations_df, fuel_prices_df):
    """
    Merge quotations dataframe with disclosed fuel dataframe.
    Args:
        quotations_df: Dataframe containing quotations data
        fuel_prices_df: Dataframe containing fuel prices data
        company_name: Name of company for logging purposes
    Returns:
        Merged dataframe
    """
    merged_dfs = []
    for quot_df in quotations_df:
        # Extract the objects name (e.g. 'df_cga_qtn_nodetr_2502')
        quot_name = quot_df.name
        # Get the period as a list
        quot_period = [n for n in quot_name.split('_') if n.isdigit()]
        if len(quot_period) > 1:  # Evaluate only one period is being passed
            raise ValueError(f"More than one period found in {quot_name}")
        # Get the supplier name
        quot_supplier = quot_df.supplier
        # Get the price column if exists
        quot_price_col = getattr(quot_df, 'price_column', None)
        
        for fuel_df in fuel_prices_df:
            # Extract the objects name (e.g. 'df_fuel_2502')
            fuel_name = fuel_df.name
            # Get the period as a list
            fuel_period = [n for n in fuel_name.split('_') if n.isdigit()]
            if len(fuel_period) > 1:  # Evaluate only one period is being passed
                raise ValueError(f"More than one period found in {fuel_name}")
            
            # Perform the merge for each matching period
            if quot_period == fuel_period:
                print("\nPerforming merge...")
                print(f"Left (quotations) dataframe: {quot_name} - shape: {quot_df.shape}")
                print(f"Right (disclosed fuel) df:   {fuel_name} - shape: {fuel_df.shape}")
                to_merge_df = pd.merge(
                                    quot_df,
                                    fuel_df,
                                    how='left',
                                    left_on='FLAG_EDS_LEMATIZED',
                                    right_on='FlagServStLematized'
                                    )
                # Re-assign attributes to the merged dataframe 
                # (use .__dict__ to force-assign attributes to the dataframe)
                to_merge_df.__dict__['supplier'] = quot_supplier 
                to_merge_df.__dict__['period'] = quot_period 
                if quot_price_col:
                    to_merge_df.__dict__['price_column'] = quot_price_col
                
                merged_dfs.append(to_merge_df)
                print(f"Merged quotations dataframe shape: {to_merge_df.shape}")

    return merged_dfs


def calculate_min_max_avg_difference(df, required_columns, base_price_column):
    """
    Calculate price differentials (min, max, avg) between suppliers and identify potential savings.
    Args:
        df: DataFrame containing supplier price data
        required_columns: Dict with 'suppliers_columns' key listing price columns
        base_price_column: Column name to compare against (reference price)
    Returns:
        DataFrame with added columns for min/max prices, absolute/percentage savings,
        and flags indicating potential savings opportunities.
    """
    df_copy = df.copy()
    suppliers_columns = required_columns['suppliers_columns']
    # Validate suppliers_columns are in the df_copy columns
    suppliers_columns = [col for col in suppliers_columns if col in df_copy.columns]

    df_copy['MinPrice'] = df_copy[suppliers_columns].min(axis=1)
    df_copy['MaxPrice'] = df_copy[suppliers_columns].max(axis=1)
    # Average already calculated in the dataframe (AverageStationPrice)

    df_copy['Potential_Svngs_Avg'] = df_copy['AverageStationPrice'] - df_copy[base_price_column]
    df_copy['Potential_Svngs_Min'] = df_copy['MinPrice'] - df_copy[base_price_column]
    df_copy['Potential_Svngs_Max'] = df_copy['MaxPrice'] - df_copy[base_price_column]
    df_copy['Potential_Svngs_%_Avg'] = (df_copy['AverageStationPrice'] / df_copy[base_price_column] - 1) * 100
    df_copy['Potential_Svngs_%_Min'] = (df_copy['MinPrice'] / df_copy[base_price_column] - 1) * 100
    df_copy['Potential_Svngs_%_Max'] = (df_copy['MaxPrice'] / df_copy[base_price_column] - 1) * 100

    # Add findings based on price difference
    df_copy['Findings_Avg'] = np.where(
                                    df_copy['Potential_Svngs_Avg'] < 0,
                                    'Potential Savings Detected',
                                    ''
                                )
    df_copy['Findings_Min'] = np.where(
                                    df_copy['Potential_Svngs_Min'] < 0,
                                    'Potential Savings Detected',
                                    ''
                                )
    df_copy['Findings_Max'] = np.where(
                                    df_copy['Potential_Svngs_Max'] < 0,
                                    'Potential Savings Detected',
                                    ''
                                )
    
    return df_copy


def arrange_columns(df, required_columns, sheet_name, base_price_column):
    """
    Filters and orders dataframe columns according to configuration.
    Combines base columns, supplier prices, analysis columns and base price column,
    then returns only columns that exist in the input dataframe.
    Args:
        df: Input DataFrame with price data
        required_columns: Dict containing:
            - base_columns: Core columns to always include
            - adt_match_info_columns: Columns with ADT matching information
            - suppliers_columns: Price columns from different suppliers 
            - analysis_columns: Additional columns per sheet_name
        sheet_name: Key to select specific analysis_columns
        base_price_column: Reference price column for comparisons
    Returns:
        DataFrame containing only requested columns that exist in input, in specified order
    """
    arranged_df = df.copy()

    # Select only desired columns and ensure proper order
    req_cols = required_columns['base_columns'].copy()
    req_cols.extend(required_columns['adt_metrics_columns'])
    req_cols.extend(required_columns['suppliers_columns'])
    req_cols.append(base_price_column)  # Append bc it is a single string, not a list
    req_cols.extend(required_columns['analysis_columns'][sheet_name])
    
    # Make sure all columns are in the dataframe
    selected_cols = [col for col in req_cols if col in arranged_df.columns]
    
    # Select only the columns that are in the dataframe
    arranged_df = arranged_df[selected_cols]
    print(f"Selected columns {sheet_name}: {selected_cols}")

    return arranged_df


def quotations_stats_dict(df, required_columns, base_price_column):
    """
    Creates a dictionary of DataFrames organized by sheet name for Excel export.
    Each DataFrame contains selected columns.
    Args:
        df: Source DataFrame containing price and analysis data
        required_columns: Configuration dict with 'analysis_columns' mapping sheet names to column lists
        base_price_column: Key column used as price reference
    Returns:
        dict: {sheet_name: DataFrame} pairs with formatted data for each analysis view
    """
    dfs_by_sheet = {}
    for sheet_name in required_columns['analysis_columns'].keys():
        arranged_df = arrange_columns(df, required_columns, sheet_name, base_price_column)
        dfs_by_sheet[sheet_name] = arranged_df
        
    return dfs_by_sheet


def qtn_min_max_avg_diff(df_merged_quot_fuel_prices, required_columns):
    """
    Processes multiple merged fuel price DataFrames to calculate and organize statistics.
    
    For each input DataFrame, calculates min/max/avg differences and organizes results
    into a dictionary of DataFrames for Excel export, while preserving metadata.
    Args:
        df_merged_quot_fuel_prices (list): List of DataFrames containing merged fuel price data.
            Each DataFrame should have attributes:
            - price_column: Column name containing price data
            - supplier: Supplier name
            - period: Time period
        required_columns (dict): Configuration dictionary specifying columns to include,
            matching the structure expected by quotations_stats_dict()
    Returns:
        list: List of dictionaries where each contains:
            - DataFrames organized by sheet name (from quotations_stats_dict)
            - Metadata (supplier, period, price_column)
    Raises:
        ValueError: If any input DataFrame is missing the required price_column attribute
    """
    stats_qtns_dict_list = []
    for df_merged in df_merged_quot_fuel_prices:
        price_column = getattr(df_merged, 'price_column', None)
        supplier = df_merged.supplier
        period = df_merged.period
        
        supplier_period = supplier.lower() + '_' + period[0]

        if price_column:
            print(f"\n{price_column} - {supplier_period}")
            df_q_mrg = calculate_min_max_avg_difference(df_merged, required_columns, price_column)
            # Create a dictionary of DataFrames organized by sheet name for Excel export.
            stats_qtn = quotations_stats_dict(df_q_mrg, required_columns, price_column)
            
            # Re-assign attributes to the stats DICTIONARY (It is a dictionary, not a dataframe)
            stats_qtn['supplier'] = supplier
            stats_qtn['period'] = period
            stats_qtn['price_column'] = price_column
            
            stats_qtns_dict_list.append(stats_qtn)
        else:
            raise ValueError(f"Missing 'price_column' in {supplier_period}. Verify 'price_column' declaration at file load")
    return stats_qtns_dict_list


def export_qtn_to_excel(dfs_by_sheet, dir_path, file_name, required_columns, base_price_column):
    """Export DataFrames to Excel with formatted columns.
    Args:
        dfs_by_sheet: {sheet_name: df} mapping
        dir_path: Output directory 
        file_name: Base filename (appended to 'fuel_optimal_prices_')
        required_columns: Dict with columns to style
        base_price_column: Reference price column to highlight
    """
    # Construct file path for output Excel file
    file_path = dir_path + 'fuel_optimal_prices_' + file_name.lower() + '.xlsx'

    # Create output directory if it doesn't exist
    os.makedirs(dir_path, exist_ok=True)

    # Determine write mode ('a'ppend or 'w'rite) based on file existence
    mode = 'a' if os.path.exists(file_path) else 'w'
    engine_kwargs = {'if_sheet_exists': 'replace'} if mode == 'a' else {}

    with pd.ExcelWriter(file_path, mode=mode, engine='openpyxl', **engine_kwargs) as writer:
        for sheet_name, df_sheet in dfs_by_sheet.items():
            
            # Only process items that are DataFrames (do not process metadata - attributes)
            if not isinstance(df_sheet, pd.DataFrame):
                continue
            
            # Identify all columns to style
            cols_to_style = required_columns['analysis_columns'][sheet_name].copy()
            
            # Round numeric columns before exporting (exclude Findings columns)
            cols_to_round = [c for c in cols_to_style if 'Findings' not in c]
            cols_to_round.extend(['KEY_MTCH_ADT_%_FREQ', 'TOT_KM_PER_ADT_FREQ'])
            for col in cols_to_round:
                if col in df_sheet.columns:
                    # Coerce to numeric, turning any errors into NaN, then round
                    df_sheet[col] = pd.to_numeric(df_sheet[col], errors='coerce').round(2)
            
            # Write dataframe to sheet
            df_sheet.to_excel(writer, sheet_name=sheet_name, index=False)

            # Get worksheet object for formatting
            worksheet = writer.sheets[sheet_name]
            
            # Define light blue fill for highlighting columns
            light_blue_fill = PatternFill(start_color='c4e3ed', end_color='c4e3ed', fill_type='solid')

            # Get columns to style from config, including price column if specified
            if base_price_column and base_price_column not in cols_to_style:
                cols_to_style.append(base_price_column)

            header = list(df_sheet.columns)
            for col_name in cols_to_style:
                if col_name in header:
                    col_idx = header.index(col_name) + 1
                    col_letter = get_column_letter(col_idx)
                    # Apply fill to entire column at once
                    for cell in worksheet[col_letter]:
                        cell.fill = light_blue_fill
    
    print(f"File '{file_path.split('/')[-1]}' exported successfully with sheets: {list(dfs_by_sheet.keys())}\n")


def qtns_stats_export_excel(stats_qtns_dict_list, project_root, required_columns):
    """
    Exports supplier quotation statistics to Excel files with proper validation.
    Processes each supplier's quotation statistics, validates required fields, and exports
    to Excel using the export_qtn_to_excel function. Handles both single and multiple
    supplier cases.
    Args:
        stats_qtns_dict_list (list[dict]): List of dictionaries containing supplier quotation statistics.
            Each dictionary must contain:
            - 'supplier': Supplier name (str)
            - 'period': Period identifier (list[str])
            Optional:
            - 'price_column': Column name containing price data (str)
        project_root (str): Root directory path for project files
        required_columns (dict): Dictionary of required columns for Excel export
    Raises:
        ValueError: If any supplier quotation is missing the required 'price_column'
    Returns:
        None: Outputs Excel files to specified directory
    """
    for stats_qtn_dict in stats_qtns_dict_list:
        # These attributes are values on a dictionary, not attributes of the dataframe
        supplier = stats_qtn_dict['supplier']
        period = stats_qtn_dict['period']
        price_column = stats_qtn_dict.get('price_column', None)

        supplier_period = supplier.lower() + '_' + period[0]
        
        if price_column:
            print(f"Exported {supplier_period} to excel")
            export_qtn_to_excel(stats_qtn_dict,
                                project_root + '/data_exports/excel_exports/',
                                supplier_period,
                                required_columns,
                                price_column
                                )
        else:
            raise ValueError(f"Missing 'price_column' in {supplier_period}. Verify 'price_column' declaration at file load")


def get_total_savings_analysis(stats_qtn_data_sources, required_columns):
    """
    Consolidate savings analysis across multiple data sources and price categories.
    Args:
        stats_qtn_data_sources (list): List of (stats_qtn_dict, period_name) tuples, 
            where each dict contains DataFrames for different price categories.
        required_columns (dict): Contains 'base_columns' (columns to merge on) and 
            'analysis_columns' (dict mapping price categories to relevant column names).
    Returns:
        dict: For each price category (e.g., 'Min_Price'), a DataFrame with merged savings 
            data per period, total average savings, and a computed Savings_Index.
    """
    # Define key columns for merging, excluding metrics that may vary across sources
    base_columns = required_columns['base_columns']  # Base columns to merge on
    base_columns_adt_metrics = base_columns + required_columns['adt_metrics_columns']  # Base columns + ADT metrics columns. Use '+', do not use .extend()
    savings_analysis_dict = {}

    for price_cat in required_columns['analysis_columns'].keys():  # (e.g. 'Min_Price', 'Max_Price', 'Average_Price')
        # Get all dataframes (e.g. stats_qtn_dict_carga_2505) for current price category (e.g. 'Min_Price')
        ## (e.g. stats_qtn_dict_carga_2502 with 'Min_Price', stats_qtn_dict_carga_2505 with 'Min_Price')
        dfs_by_category = [source_dict[price_cat] for source_dict, period in stats_qtn_data_sources]
        
        # Create a base_supplier_svngs_df, dropping duplicates based on key columns only
        # The values for the metric columns will be taken from the first occurrence of the route
        base_supplier_svngs_df = pd.concat([df[base_columns_adt_metrics] for df in dfs_by_category], 
                                           ignore_index=True
                                           ).drop_duplicates(subset=base_columns, keep='first')
        
        # Get the name of the savings column for the current category (e.g. 'Potential_Svngs_%_Min', 'Potential_Svngs_%_Avg')
        savings_col_name = required_columns['analysis_columns'][price_cat][2]  # Third element of the list (e.g. 'Potential_Svngs_%_Max')

        # Merge savings data from each source using the base_columns
        for source_dict, period in stats_qtn_data_sources:
            df = source_dict[price_cat]
            # Prepare a temporary dataframe with key columns and the renamed savings column
            df_savings = df[base_columns + [savings_col_name]].rename(
                columns={savings_col_name: f'{savings_col_name}_{period}'})
            # Merge the savings data into the main dataframe for this category
            base_supplier_svngs_df = pd.merge(base_supplier_svngs_df, df_savings, on=base_columns, how='left')
        
        # Add Total and Savings_Index columns
        ## Add the total average of the savings columns
        base_supplier_svngs_df[f'{savings_col_name}_Total'] = base_supplier_svngs_df.filter(regex=f'{savings_col_name}_.*').mean(axis=1)
        ## Add the Savings_Index column
        base_supplier_svngs_df['Savings_Index'] = base_supplier_svngs_df[f'{savings_col_name}_Total'] * base_supplier_svngs_df['TOT_KM_PER_ADT_FREQ']
        
        # Add the consolidated dataframe to the dictionary
        savings_analysis_dict[price_cat] = base_supplier_svngs_df
        print(f"Shape for {price_cat}: {base_supplier_svngs_df.shape}")
    print("-------------------------------")
    
    return savings_analysis_dict


def analyze_suppliers_savings(stats_qtns_dict_list, required_columns):
    """
    Analyze savings for all suppliers by:
    1. Getting unique suppliers from stats_qtns_dict_list
    2. For each supplier, collect their period-specific data
    3. Calculate savings analysis using get_total_savings_analysis()
    4. Return consolidated analysis for all suppliers
    
    Args:
        stats_qtns_dict_list: List of dictionaries containing supplier quotation data
        required_columns: Dictionary defining required columns for analysis
        
    Returns:
        List of dictionaries containing savings analysis per supplier
    """
    # Get unique suppliers
    unique_suppliers_list = list(set(d['supplier'] for d in stats_qtns_dict_list))
    
    dict_all_suppliers_savings_analysis = []
    
    for supplier in unique_suppliers_list:
        # Collect all period data for current supplier
        supplier_stats_qtn_data_sources = []
        
        for stats_qtn_dict in stats_qtns_dict_list:
            if stats_qtn_dict['supplier'] == supplier:
                period = stats_qtn_dict['period'][0]  # Extract period from list
                supplier_stats_qtn_data_sources.append((stats_qtn_dict, period))
        
        print(f"Processing supplier: {supplier}")
        
        # Calculate savings analysis for current supplier
        dict_supplier_savings_analysis = get_total_savings_analysis(
                                                                supplier_stats_qtn_data_sources, 
                                                                required_columns
                                                                )
        dict_supplier_savings_analysis['supplier'] = supplier
        dict_all_suppliers_savings_analysis.append(dict_supplier_savings_analysis)
    
    # Display the first 3 rows of the Average_Price dataframe for each supplier
    for s in dict_all_suppliers_savings_analysis:
        print(f"\nSample of {s['supplier']} Average Price:")
        display(s['Average_Price'].head(3))
    
    return dict_all_suppliers_savings_analysis


def export_dataframes_to_excel_sheets(data_dict, file_path, required_columns):
    """
    Exports a dictionary of DataFrames to an Excel file, applying styling.
    Each key-value pair in the dictionary becomes a separate worksheet.
    It rounds numeric columns and applies a blue background to all columns
    except for the base columns.
    Args:
        data_dict (dict): {sheet_name: DataFrame} mapping to export.
        file_path (str): Full path including filename for Excel output.
        required_columns (dict): Configuration dict containing 'base_columns'.
    """
    # Ensure the output directory exists
    os.makedirs(os.path.dirname(file_path), exist_ok=True)
    
    with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
        base_cols = required_columns['base_columns']
        light_blue_fill = PatternFill(start_color='c4e3ed', end_color='c4e3ed', fill_type='solid')

        for sheet_name, df in data_dict.items():
            if not isinstance(df, pd.DataFrame):  # Skip non-DataFrame values (e.g. metadata attributes)
                continue
            df_copy = df.copy()

            # Identify columns to style (all columns except base_columns)
            cols_to_style = [col for col in df_copy.columns if col not in base_cols]
            
            # Round all styled columns to 2 decimal places
            for col in cols_to_style:
                if col in df_copy.columns:
                    df_copy[col] = pd.to_numeric(df_copy[col], errors='coerce').round(2)

            # Write dataframe to sheet
            df_copy.to_excel(writer, sheet_name=sheet_name, index=False)

            # Get worksheet object for formatting
            worksheet = writer.sheets[sheet_name]
            
            # Apply fill to styled columns
            col_prefix_not_to_color = required_columns['analysis_columns'][sheet_name][2]
            cols_to_style = [col for col in cols_to_style
                             if not (  # If not starts with the prefix and ends with 4 numbers
                                 col.startswith(col_prefix_not_to_color) and
                                 re.search(r'_\d{4}$', col))  # To avoid filtering the columns ending in '_Total'
                            ]
            header = list(df_copy.columns)
            for col_name in cols_to_style:
                if col_name in header:
                    col_idx = header.index(col_name) + 1
                    col_letter = get_column_letter(col_idx)
                    for cell in worksheet[col_letter]:
                        cell.fill = light_blue_fill


def export_savings_analysis_to_excel(dict_all_suppliers_svngs_analysis, project_root, required_columns):
    """
    Exports savings analysis data for multiple suppliers to Excel files.
    For each supplier in the input dictionary, this function exports their savings analysis data
    to a separate Excel file with multiple sheets.
    Args:
        dict_all_suppliers_svngs_analysis (list): List of dictionaries containing savings analysis data
            for each supplier. Each dictionary should contain a 'supplier' key.
        project_root (str): Root directory path for the project where files will be saved.
        required_columns (list): List of columns to include in the exported Excel files.
    Returns:
        None: Files are saved to disk but nothing is returned.
    """
    for dict_supplier_savings_analysis in dict_all_suppliers_svngs_analysis:
        # Get the supplier name. Convert to lowercase
        supplier = dict_supplier_savings_analysis['supplier'].lower()
        print(f"Exporting {supplier.capitalize()} to Excel")
        
        export_dataframes_to_excel_sheets(
                    dict_supplier_savings_analysis,
                    project_root + f'/data_exports/excel_exports/fuel_optimal_by_adt_freq_{supplier}.xlsx',
                    required_columns
                )


def plot_savings_vs_km(data_dict, supplier_name):
    """
    Plot potential savings vs total KM per route from a savings analysis dictionary.
    
    Args:
        data_dict (dict): Dictionary containing savings analysis dataframes
        supplier_name (str): Name of the supplier to be used in the plot title
    """
    for key, df in data_dict.items():
        if not isinstance(df, pd.DataFrame):  # Skip non-DataFrame values (e.g. metadata attributes)
            continue

        dataframe = df.copy()
        x_col = [col for col in dataframe.columns if col.startswith('Potential_Svngs_%_') and col.endswith('_Total')][0]
        
        # Clean column names by replacing spaces with underscores
        dataframe.columns = dataframe.columns.str.replace(' ', '_')

        # Invert savings column and filter positive values
        dataframe['Inverted_Savings'] = dataframe[x_col] * -1
        positive_inverted_savings_df = dataframe[dataframe['Inverted_Savings'] > 0].copy()

        # Create plot
        plt.figure(figsize=(12, 6))
        
        scatterplot = sns.scatterplot(
            data=positive_inverted_savings_df,
            x='Inverted_Savings',
            y='TOT_KM_PER_ADT_FREQ',
            # s=100,  # Increased marker size from default to 100
            color='blue',
        )

        # Configure plot
        plt.title(f'{supplier_name} Potential Savings {key} vs. Total KM per Route', fontsize=16)
        plt.xlabel('Potential Savings % (Inverted from Negative Values)', fontsize=12)
        plt.ylabel('Total KM per ADT Frequency (Log Scale)', fontsize=12)
        plt.yscale('log')

        # Add route code labels
        for i, point in positive_inverted_savings_df.iterrows():
            scatterplot.text(point['Inverted_Savings'] + 0.01, 
                            point['TOT_KM_PER_ADT_FREQ'], 
                            str(point['CODIGO_RUTA']), 
                            fontdict={'size': 8})

        plt.tight_layout()
        plt.show()


def plot_savings_vs_km_all_suppliers(dict_all_suppliers_svngs_analysis):
    """
    Plots savings vs kilometers for all suppliers in the analysis.
    Iterates through each supplier's savings analysis dictionary and generates
    a savings vs kilometers plot for each one. The supplier name is extracted
    and capitalized for display purposes.
    Parameters:
    dict_all_suppliers_svngs_analysis : list of dict
        List containing savings analysis dictionaries for all suppliers.
        Each dictionary should contain a 'supplier' key with the supplier name.
    Returns:
    None
        Generates plots but doesn't return any value.
    """
    for dict_supplier_savings_analysis in dict_all_suppliers_svngs_analysis:
        # Get the supplier name. Capitalize it.
        supplier = dict_supplier_savings_analysis['supplier'].capitalize()
        print(f"\nPlotting {supplier} Savings vs. KM")
        
        plot_savings_vs_km(dict_supplier_savings_analysis, supplier)


def preprocess_terpel(df_terpel_prices):
    """Preprocess Terpel dataframe.
    Args:
        df_terpel_prices: Terpel dataframe
    Returns:
        df_terpel_prices: Preprocessed Terpel dataframe
    """
    # Transform column names to lowercase, strip spaces and replace spaces with underscores
    df_terpel_prices.columns = df_terpel_prices.columns.str.lower().str.strip().str.replace(r'\s+', '_', regex=True)

    # Rename columns
    df_terpel_prices.rename(columns={'año': 'year', 'centro_de_distribución': 'distribution_center'}, inplace=True)
    print(f"Terpel's columns: {df_terpel_prices.columns},\nshape: {df_terpel_prices.shape}")

    print(f"Shape of Terpel dataframe before cleaning: {df_terpel_prices.shape}")
    # Convert 'fecha' to datetime, turning unparseable dates into NaT. It has the format mm/dd/yyyy
    # dayfirst=True helps correctly parse formats like 'dd/mm/yyyy'
    df_terpel_prices['fecha'] = pd.to_datetime(df_terpel_prices['fecha'], errors='coerce', dayfirst=True)
    ## Remove rows with invalid dates (NaT)
    df_terpel_prices.dropna(subset=['fecha'], inplace=True)
    print(f"Shape of Terpel dataframe after removing rows with invalid dates: {df_terpel_prices.shape}")

    # For string_columns, clean the data by stripping whitespace and handling missing values
    ## Detect non-numeric (str) columns
    string_cols = df_terpel_prices.select_dtypes(include=['object']).columns.tolist()
    ## Make values uppercase, strip whitespace and replace multiple spaces with a single space
    for col in string_cols:
        df_terpel_prices[col] = df_terpel_prices[col].str.upper().str.strip().str.replace(r'\s+', ' ', regex=True)
        # Replace empty strings, 'NA', 'N/A' with NaN (It does not delete the rows, just replaces the values)
        df_terpel_prices[col] = df_terpel_prices[col].replace(['', 'NA', 'N/A'], np.nan)
    print(f"Shape of Terpel dataframe after cleaning string columns: {df_terpel_prices.shape}")
    return df_terpel_prices


def validate_terpel_date_consistency(df_terpel_prices):
    # Validate date consistency between columns. 
    ## Create a conversion dict from month to a number
    month_to_number = {
        'ENERO': 1,
        'FEBRERO': 2,
        'MARZO': 3,
        'ABRIL': 4,
        'MAYO': 5,
        'JUNIO': 6,
        'JULIO': 7,
        'AGOSTO': 8,
        'SEPTIEMBRE': 9,
        'OCTUBRE': 10,
        'NOVIEMBRE': 11,
        'DICIEMBRE': 12
    }

    ## A record is considered valid if its 'Mes' and 'año' either match the 'fecha' month and year, or the 
    ## month immediately following. This accounts for transactions made at the end of one month and recorded in the next.
    ## Condition 1: 'mes' and 'year' match the month and year of 'fecha'.
    same_month_mask = (df_terpel_prices['fecha'].dt.year == df_terpel_prices['year']) & \
                    (df_terpel_prices['fecha'].dt.month == df_terpel_prices['mes'].map(month_to_number))
    ## Condition 2: 'mes' and 'año' match the month and year following 'fecha'.
    ## pd.DateOffset(months=1) correctly handles year changes (e.g., December to January).
    next_month_date = df_terpel_prices['fecha'] + pd.DateOffset(months=1)
    next_month_mask = (next_month_date.dt.year == df_terpel_prices['year']) & \
                    (next_month_date.dt.month == df_terpel_prices['mes'].map(month_to_number))
    ## The final mask includes rows that satisfy either condition.
    mask = same_month_mask | next_month_mask
    ## Filter the dataframe
    df_terpel_prices = df_terpel_prices[mask]
    # df_non_matching_dates = df_terpel_prices[~mask]
    print(f"Shape of Terpel dataframe after filtering by date consistency: {df_terpel_prices.shape}")
    
    return df_terpel_prices


def delete_terpel_inconsistent_data(df_terpel_prices):
    # Delete inconsistent data
    ## Drop rows where producto is 'Urea' and unidad_de_medida is 'Galones' (Urea is sold in liters)
    urea_mask = (df_terpel_prices['producto'] == 'Urea') & (df_terpel_prices['unidad_de_medida'] == 'Galones')
    df_terpel_prices = df_terpel_prices[~urea_mask]
    print(f"Shape of Terpel dataframe after removing rows with urea and galones: {df_terpel_prices.shape}")

    ## Drop rows where nombre_eds doesn't have at least 3 alphanumeric characters (letters or numbers)
    eds_mask = ~df_terpel_prices['nombre_eds'].str.contains(r'[a-zA-Z0-9]{3}', regex=True)
    df_terpel_prices = df_terpel_prices[~eds_mask]
    print(f"Shape of Terpel dataframe after removing rows with less than 3 alphanumeric characters: {df_terpel_prices.shape}")

    display(df_terpel_prices.head(3))
    print(f"Shape of Terpel filtered dataframe: {df_terpel_prices.shape}")

    return df_terpel_prices


def mrge_suplrs_qtns_to_terpel(terpel_df, supplier_df, supplier_price_col, supl_price_rename_column, left_key='FlagTerpelStLematized', right_key='FLAG_EDS_LEMATIZED'):
    """
    Merge supplier fuel quotations data with Terpel's monthly price dataframe by lematized EDS names.
    Args:
        terpel_df (pd.DataFrame): Target DataFrame containing Terpel's monthly fuel prices
        supplier_df (pd.DataFrame): Source DataFrame containing supplier fuel quotations
        supplier_price_col (str): Column name in supplier_df containing fuel prices
        supl_price_rename_column (str): New name for the supplier price column in output
        left_key (str, optional): Key column in terpel_df. Defaults to 'FlagTerpelStLematized'
        right_key (str, optional): Key column in supplier_df. Defaults to 'FLAG_EDS_LEMATIZED'
    Returns:
        Terpel merged dataframe with supplier quotations added
    """
    # Extract and process supplier quotations
    supplier_quot = supplier_df[[right_key, supplier_price_col]]
    supplier_quot = supplier_quot.groupby(right_key, as_index=False)[supplier_price_col].mean()
    supplier_quot = supplier_quot.rename(columns={supplier_price_col: supl_price_rename_column})
    
    # Merge with Terpel dataframe
    merged_df = pd.merge(terpel_df, 
                        supplier_quot, 
                        left_on=left_key, 
                        right_on=right_key, 
                        how='left')
    
    return merged_df.drop(columns=[right_key])


def merge_all_suppliers_qtn_to_terpel(df_qtn_all_split, terpel_diesel_monthly_cols):
    """
    Merge all supplier fuel quotations into Terpel's monthly diesel price dataframe.
    Processes each supplier's quotations and merges them with the Terpel dataframe,
    creating new columns for each supplier's price quotations by period.
    Note: Only supplier quotations with EDS in their files are merged (EDS required).
    Some suppliers like Otransa, Iceberg and ADP may be excluded as they don't have EDS in their quotations.
    Args:
        df_qtn_all_split (list): List of supplier DataFrames containing fuel quotations.
                                Each DataFrame should have these attributes:
                                - supplier (str): Supplier name (e.g. 'Carga', 'Transer')
                                - name (str): Period extracted from object namein format 'prefix_period' (e.g. 'object_name_2501')
                                - price_column (str): Column name containing price data
        terpel_diesel_monthly_cols (pd.DataFrame): Terpel's monthly diesel price dataframe containing:
                                - FlagTerpelStLematized: Key column for merging with supplier EDS names
                                - Other monthly price columns
    Returns:
        pd.DataFrame: Terpel dataframe with merged supplier quotations as new columns.
                     New column names follow format: '{supplier}_quot_{period}'
                     (e.g. 'carga_quot_2501')
    """
    suppliers = list({s.supplier for s in df_qtn_all_split})
    print(f"Unique suppliers: {suppliers}")
    
    shape_before_merging = terpel_diesel_monthly_cols.shape

    for supplier in suppliers:
        print(f"\nSupplier: {supplier}")

        for df_supplier_qtn in df_qtn_all_split:
            suppl_name = df_supplier_qtn.supplier
            if suppl_name == supplier:
                print("-" * 5)
                # Get period and price column
                period = df_supplier_qtn.name.split('_')[-1]
                supplier_price_col = df_supplier_qtn.price_column
                # Generate the price rename column name for each supplier
                supl_price_rename_column = suppl_name.lower() + '_quot_' + period

                terpel_diesel_monthly_cols = mrge_suplrs_qtns_to_terpel(
                                                terpel_diesel_monthly_cols,
                                                df_supplier_qtn,
                                                supplier_price_col,
                                                supl_price_rename_column
                                                )
                print(supplier_price_col)
                print(supl_price_rename_column)
        print("-"*50)
    
    print(f"Shape of terpel_diesel_monthly_cols before merging suppliers quotations: {shape_before_merging}")
    print(f"Shape of terpel_diesel_monthly_cols after merging suppliers quotations:  {terpel_diesel_monthly_cols.shape}\n")
    display(terpel_diesel_monthly_cols.head(3))

    return terpel_diesel_monthly_cols


def consolidate_adt_frequency(terpel_main_df, df_qtn_all_split, terpel_eds_key='FlagTerpelStLematized', supplier_eds_key='FLAG_EDS_LEMATIZED'):
    """
    Consolidates ADT frequency from multiple supplier dataframes into the main dataframe.
    Args:
        terpel_main_df: The main dataframe to merge frequencies into
        df_qtn_all_split: List of supplier dataframes containing frequency data
        terpel_eds_key: Key column name in Terpel dataframe
        supplier_eds_key: Key column name in supplier dataframes
    Returns:
        DataFrame with merged frequency column
    """
    # Extract and combine frequency data
    frequency_data = pd.concat([
        df[[supplier_eds_key, 'KEY_MTCH_ADT_%_FREQ']] for df in df_qtn_all_split if 'KEY_MTCH_ADT_%_FREQ' in df.columns
    ]).dropna()

    # Calculate the total frequency of adt routes by EDS
    frequency_by_eds = frequency_data.groupby(supplier_eds_key)['KEY_MTCH_ADT_%_FREQ'].sum()
    total_frequency = frequency_by_eds.sum()

    # Calculate each eds's frequency as a percentage of the total frequency
    if total_frequency > 0:
        adt_frequency_map = (frequency_by_eds / total_frequency * 100).reset_index()
    else:
        adt_frequency_map = frequency_by_eds.reset_index()
        adt_frequency_map['KEY_MTCH_ADT_%_FREQ'] = 0
    
    adt_frequency_map = adt_frequency_map.rename(columns={'KEY_MTCH_ADT_%_FREQ': 'adt_frequency(%)'})

    # Merge the ADT frequency into the main dataframe
    merged_df = pd.merge(
        terpel_main_df,
        adt_frequency_map,
        left_on=terpel_eds_key,
        right_on=supplier_eds_key,
        how='left'
    )

    # Clean up by dropping the redundant key column from the merge
    if supplier_eds_key in merged_df.columns:
        merged_df = merged_df.drop(columns=[supplier_eds_key])

    print("Shape of terpel_diesel_monthly_cols after adding adt_frequency(%):", merged_df.shape)
    display(merged_df.head(3))

    return merged_df


def mrg_disclosed_fuel_prices_to_terpel(terpel_df, disclosed_fuel_df, discl_suffix, unique_buyers_list, left_key='FlagTerpelStLematized', right_key='FlagServStLematized'):
    """
    Processes and merges the disclosed supplier fuel price with Terpel dataframe.
    Args:
        terpel_df (pd.DataFrame): Terpel's main dataframe
        disclosed_fuel_df (pd.DataFrame): Disclosed fuel prices dataframe
        discl_suffix (str): Suffix to add to the column names
        unique_buyers_list (list): List of buyer names to rename columns
        merge_key (str): Column name to merge on (default 'FlagTerpelStLematized')
    Returns:
        pd.DataFrame: Merged dataframe with supplier fuel prices
    """
    # Process supplier fuel data
    ## Delete the last row of the disclosed_fuel_df (it is the average of each supplier)
    fuel_by_eds = disclosed_fuel_df.iloc[:-1]  # Remove last (average) row
    ## Drop some non-relevant columns
    fuel_by_eds = fuel_by_eds.drop(columns=['AverageStationPrice', 'Department', 'ServiceStation'])
    
    ## Add a suffix to the columns that are in the unique_buyers_list. Make the column names lowercase
    fuel_by_eds = fuel_by_eds.rename(
        columns={buyer: f'{buyer.lower()}{discl_suffix}' for buyer in unique_buyers_list}
    )
    
    # Merge the disclosed fuel df with Terpel dataframe
    merged_df = pd.merge(
        terpel_df,
        fuel_by_eds,
        left_on=left_key,
        right_on=right_key,
        how='left'
    )
    
    return merged_df.drop(columns=[right_key])


def all_merge_disclosed_fuel_prices_to_terpel(terpel_diesel_monthly_cols, pivot_tables_disclosed_fuel_list, unique_buyers_all_periods):
    """
    Merges multiple disclosed fuel price DataFrames with the main Terpel DataFrame.
    Processes each disclosed fuel price DataFrame in the input list, extracts the period from its name,
    and merges it with the main Terpel DataFrame using a period-specific suffix for column names.
    Args:
        terpel_diesel_monthly_cols (pd.DataFrame): Main Terpel DataFrame containing fuel data
        pivot_tables_disclosed_fuel_list (list): List of DataFrames with disclosed fuel prices, 
            each DataFrame should have a .name attribute containing the period identifier
        unique_buyers_all_periods (list): List of buyer names to be used for column renaming
    Returns:
        pd.DataFrame: The merged Terpel DataFrame with all disclosed fuel prices added as new columns
    """
    shape_before = terpel_diesel_monthly_cols.shape
    for df_price_discl_fuel_period in pivot_tables_disclosed_fuel_list:
        # Get the name of the period
        name = df_price_discl_fuel_period.name
        print(f"Merging Terpel and disclosed fuel prices for {name}")
        # Extract the period from the name (e.g. 'object_name_2501')
        period = name.split('_')[-1]
        # Create a suffix for the disclosed fuel prices (e.g. '_discl_2501')
        discl_suffix = '_discl_' + period
        terpel_diesel_monthly_cols = mrg_disclosed_fuel_prices_to_terpel(
                                                                    terpel_diesel_monthly_cols,
                                                                    df_price_discl_fuel_period,
                                                                    discl_suffix,
                                                                    unique_buyers_all_periods
                                                                    )
    display(terpel_diesel_monthly_cols.head(3))
    print(f"Shape of terpel_diesel_monthly_cols before merging disclosed fuel prices: {shape_before}")
    print(f"Shape of terpel_diesel_monthly_cols after merging disclosed fuel prices: {terpel_diesel_monthly_cols.shape}")

    return terpel_diesel_monthly_cols


def calculate_total_average_delta(terpel_diesel_monthly_cols):
    # Calculate the total average of the delta in $ and % (row-wise)
    ## Add a column with the average of the quotations delta in $ and %
    terpel_diesel_monthly_cols['quotation_delta($)'] = terpel_diesel_monthly_cols.filter(regex=r'.*_quot_.*_delta\(\$\).*').mean(axis=1)
    terpel_diesel_monthly_cols['quotation_delta(%)'] = terpel_diesel_monthly_cols.filter(regex=r'.*_quot_.*_delta\(\%\).*').mean(axis=1)

    ## Add a column with the average of the disclosed delta in $ and %
    terpel_diesel_monthly_cols['disclosed_delta($)'] = terpel_diesel_monthly_cols.filter(regex=r'.*_discl_.*_delta\(\$\).*').mean(axis=1)
    terpel_diesel_monthly_cols['disclosed_delta(%)'] = terpel_diesel_monthly_cols.filter(regex=r'.*_discl_.*_delta\(\%\).*').mean(axis=1)

    display(terpel_diesel_monthly_cols.head(3))
    print(f"Shape of terpel_diesel_monthly_cols after general delta calculation: {terpel_diesel_monthly_cols.shape}")

    print(f"\nquotation_delta($): {terpel_diesel_monthly_cols['quotation_delta($)'].unique()}")
    print(f"\nquotation_delta(%): {terpel_diesel_monthly_cols['quotation_delta(%)'].unique()}")
    print(f"\ndisclosed_delta($): {terpel_diesel_monthly_cols['disclosed_delta($)'].unique()}")
    print(f"\ndisclosed_delta(%): {terpel_diesel_monthly_cols['disclosed_delta(%)'].unique()}")

    return terpel_diesel_monthly_cols


def calculate_terpel_delta_columns(terpel_diesel_monthly_cols):
    # Identify columns with all NaN values
    nan_terpel_match_cols = terpel_diesel_monthly_cols.columns[terpel_diesel_monthly_cols.isna().all()].tolist()
    print(f"Columns with all NaN values: {nan_terpel_match_cols}\nThis means that the Terpel's EDS did not match with any of the quotations/disclosed fuel prices for that month")

    # Create a list of the columns that start with 'terpel_'
    terpel_cols = [col for col in terpel_diesel_monthly_cols.columns if col.startswith('terpel_')]
    print(f"terpel_cols: {terpel_cols}\n")

    # Create a list of the columns that are not in the terpel_cols list and are not in the nan_terpel_match_cols list
    non_terpel_cols = [col for col in terpel_diesel_monthly_cols.columns 
                    if not col.startswith('terpel_') 
                    and col not in ['FlagTerpelStLematized', 'nombre_eds']  # Exclude columns that are not Terpel's prices, nor quotations/disclosed prices
                    and col not in nan_terpel_match_cols]
    print(f"non_terpel_cols: {non_terpel_cols}\n")

    # Calculate the delta in $ and % for each column in the non_terpel_cols list
    for ter_yr_col in terpel_cols:  # Iterate over the Terpel's columns
        for non_ter_col in non_terpel_cols:  # Iterate over the non-Terpel's columns
            if ter_yr_col.split('_')[-1] == non_ter_col.split('_')[-1]:  # If the year is the same
                terpel_diesel_monthly_cols[f'{non_ter_col}_delta($)'] = terpel_diesel_monthly_cols[ter_yr_col] - terpel_diesel_monthly_cols[non_ter_col]
                terpel_diesel_monthly_cols[f'{non_ter_col}_delta(%)'] = terpel_diesel_monthly_cols[ter_yr_col] / terpel_diesel_monthly_cols[non_ter_col] - 1

    print(f"Shape of terpel_diesel_monthly_cols after columns delta calculation: {terpel_diesel_monthly_cols.shape}")

    # Calculate the average of the delta in $ and % (row-wise)
    terpel_diesel_monthly_cols = calculate_total_average_delta(terpel_diesel_monthly_cols)

    return terpel_diesel_monthly_cols


def move_column_to_end(df, column_name):
    """
    Moves specified column to the end of the dataframe columns list
    Args:
        df: pandas DataFrame
        column_name: str, name of column to move
        
    Returns:
        DataFrame with reordered columns
    """
    cols = df.columns.tolist()
    cols.remove(column_name)
    cols.append(column_name)
    return df[cols]


def define_terpel_column_types(terpel_diesel_monthly_cols):
    """
    Selects key columns for Terpel analysis, converts them to numeric types, and
    returns a new DataFrame ready for analysis.

    This function is designed to create a summary DataFrame for analysis by
    selecting specific columns by name, making it robust against changes in
    column order. It identifies columns related to Terpel's prices, calculated
    deltas, and ADT frequency.

    Args:
        terpel_diesel_monthly_cols (pd.DataFrame): The input DataFrame which
            contains monthly Terpel prices, supplier quotations, and calculated
            deltas.

    Returns:
        pd.DataFrame: A new DataFrame containing only the selected columns, with
            numeric columns converted to float types for further analysis.
    """
    # Define the columns to be selected for the analysis.
    # Identifier columns.
    # id_cols = ['FlagTerpelStLematized', 'nombre_eds']
    id_cols = ['nombre_eds']
    # Terpel's monthly price columns.
    terpel_cols = [col for col in terpel_diesel_monthly_cols.columns if col.startswith('terpel_')]
    # Analysis columns (deltas and frequency).
    analysis_cols = [
                    'quotation_delta($)', 
                    'quotation_delta(%)',
                    'disclosed_delta($)', 
                    'disclosed_delta(%)',
                    'adt_frequency(%)'
                    ]

    # Combine all column lists, ensuring they exist in the DataFrame.
    selected_cols = id_cols + terpel_cols + analysis_cols
    existing_cols = [col for col in selected_cols if col in terpel_diesel_monthly_cols.columns]

    # Create the analysis DataFrame with the selected columns.
    terpel_all_analysis = terpel_diesel_monthly_cols[existing_cols].copy()

    # Convert all columns except the identifiers to numeric types.
    for col in terpel_all_analysis.columns:
        if col not in id_cols:
            if terpel_all_analysis[col].dtype == 'object':
                # Clean and convert object columns to numeric.
                cleaned_col = terpel_all_analysis[col].astype(str).str.replace(r'[$,%]', '', regex=True)
                terpel_all_analysis[col] = pd.to_numeric(cleaned_col, errors='coerce')
            else:
                # Ensure other numeric types are converted, handling any potential errors.
                terpel_all_analysis[col] = pd.to_numeric(terpel_all_analysis[col], errors='coerce')

    return terpel_all_analysis


def create_terpel_analysis_columns(terpel_all_analysis):
    ## Make the _delta(%) columns percentage (x 100)
    terpel_all_analysis['disclosed_delta(%)'] = (terpel_all_analysis['disclosed_delta(%)'] * 100)
    terpel_all_analysis['quotation_delta(%)'] = (terpel_all_analysis['quotation_delta(%)'] * 100)

    # Insert savings_index columns
    ## Savings_Index_Quotations: 
    terpel_all_analysis['savings_index_quotations'] = terpel_all_analysis['quotation_delta(%)'] * terpel_all_analysis['adt_frequency(%)']
    ## Savings_Index_Disclosed: 
    terpel_all_analysis['savings_index_disclosed'] = terpel_all_analysis['disclosed_delta(%)'] * terpel_all_analysis['adt_frequency(%)']
    
    # Make all columns 2 decimal places
    terpel_all_analysis = terpel_all_analysis.round(2)
    print(f"Shape of terpel_all_analysis: {terpel_all_analysis.shape}")

    return terpel_all_analysis


def split_terpel_df_into_printable_dfs(terpel_all_analysis, terpel_diesel_monthly_cols, verbose=False):
    """
    Splits the terpel_all_analysis dataframe into printable dataframes
    Args:
        terpel_all_analysis: pandas DataFrame
        verbose: bool, whether to display the dataframes
    Returns:
        list of tuples with the dataframes and their names
    """
    # Display the rows with any value in the total delta columns
    terpel_with_match = terpel_all_analysis.dropna(how='all', subset=terpel_diesel_monthly_cols.columns[-4:])
    if verbose:
        print("\nTerpel with match:")
        display(terpel_with_match)
        print(f"Shape of terpel_with_match: {terpel_with_match.shape}")

        print("\nTerpel is cheaper than the disclosed/quotations prices (negative values):")
        print("\nSavings detected in disclosed prices:")
    terp_below_disclosed = terpel_with_match[(terpel_with_match['disclosed_delta($)'] < 0) | (terpel_with_match['disclosed_delta(%)'] < 0)]
    if verbose:
        display(terp_below_disclosed)
        print(f"Shape of terp_below_disclosed: {terp_below_disclosed.shape}")

        print("\nSavings detected in quotations:")
    terp_below_quotations = terpel_with_match[(terpel_with_match['quotation_delta($)'] < 0) | (terpel_with_match['quotation_delta(%)'] < 0)]
    if verbose:
        display(terp_below_quotations)
        print(f"Shape of terp_below_quotations: {terp_below_quotations.shape}")

        print("\nTerpel is more expensive than the disclosed/quotations prices (positive values):")
        print("\nOvercost detected in disclosed prices:")
    terp_above_disclosed = terpel_with_match[(terpel_with_match['disclosed_delta($)'] > 0) | (terpel_with_match['disclosed_delta(%)'] > 0)]
    if verbose:
        display(terp_above_disclosed)
        print(f"Shape of terp_above_disclosed: {terp_above_disclosed.shape}")

        print("\nOvercost detected in quotations:")
    terp_above_quotations = terpel_with_match[(terpel_with_match['quotation_delta($)'] > 0) | (terpel_with_match['quotation_delta(%)'] > 0)]
    if verbose:
        display(terp_above_quotations)
        print(f"Shape of terp_above_quotations: {terp_above_quotations.shape}")

    # Export analysis results to Excel with sheetname
    terpel_dataframes = [
        (terp_below_disclosed, 'terp_below_disclosed', 'disclosed_delta(%)'),
        (terp_below_quotations, 'terp_below_quotations', 'quotation_delta(%)'),
        (terp_above_disclosed, 'terp_above_disclosed', 'disclosed_delta(%)'),
        (terp_above_quotations, 'terp_above_quotations', 'quotation_delta(%)'),
        (terpel_with_match, 'all_matches', ''),
        (terpel_all_analysis, 'all_analysis', ''),
    ]

    return terpel_dataframes


def export_terpel_to_excel(terpel_dataframes, file_path):
    """Export list of dataframes to Excel with each dataframe as a separate sheet.
    Args:
        terpel_dataframes: List of tuples containing (dataframe, sheet_name)
        file_path: Path to save Excel file
    """
    directory = os.path.dirname(file_path)
    if directory:  # If filepath has a directory
        os.makedirs(directory, exist_ok=True)  # Create directory if it doesn't exist

    # Determine the mode for ExcelWriter based on file existence
    mode = 'a' if os.path.exists(file_path) else 'w'
    engine_kwargs = {'if_sheet_exists': 'replace'} if mode == 'a' else {}
    
    with pd.ExcelWriter(file_path, mode=mode, engine='openpyxl', **engine_kwargs) as writer:
        light_blue_fill = PatternFill(start_color='c4e3ed', end_color='c4e3ed', fill_type='solid')
        for df, sheet_name, _ in terpel_dataframes:
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            worksheet = writer.sheets[sheet_name]
            
            header = list(df.columns)
            cols_to_style = ['quotation_delta(%)', 'disclosed_delta(%)']
            cols_to_style.extend(df.columns[-3:].tolist())

            for col_name in cols_to_style:
                col_idx = header.index(col_name) + 1
                col_letter = get_column_letter(col_idx)
                for cell in worksheet[col_letter]:
                    cell.fill = light_blue_fill


def plot_terpel_deltas_vs_adt_frequency(terpel_dataframes, x_capped=True):
    for df in terpel_dataframes[:4]:
        df_eds_analysis = df[0].copy()
        df_name = df[1]
        delt_type = df[2]

        # Skip if dataframe is empty or required columns don't exist
        if df_eds_analysis.empty or delt_type not in df_eds_analysis.columns or 'adt_frequency(%)' not in df_eds_analysis.columns:
            print(f"Skipping {df_name} - no matching data to plot")
            continue

        # Invert the 'quotation_delta(%)' column for the x-axis
        df_eds_analysis['inverted_delta'] = df_eds_analysis[delt_type] * -1
        if x_capped:
            positive_inverted_delta_df = df_eds_analysis[df_eds_analysis['inverted_delta'] > 0].copy()
        else:
            positive_inverted_delta_df = df_eds_analysis

        # Skip if no data points after filtering
        if positive_inverted_delta_df.empty:
            print(f"Skipping {df_name} - no data points after filtering")
            continue

        plt.figure(figsize=(12, 6))

        # Create the scatterplot with the inverted x-axis data
        scatterplot = sns.scatterplot(
            data=positive_inverted_delta_df,
            x='inverted_delta',
            y='adt_frequency(%)',
            s=100,
            color='blue'
        )

        # Add labels to each point using the 'nombre_eds' column
        # A small offset is added to prevent text from overlapping the points
        for i, point in positive_inverted_delta_df.iterrows():
            scatterplot.text(
                point['inverted_delta'] + 0.005,  # Small horizontal offset
                point['adt_frequency(%)'], 
                str(point['nombre_eds']), 
                fontdict={'size': 9}
            )

        # Set plot titles and axis labels for clarity
        plt.title(f'TERPEL - Quotation Delta vs. ADT Frequency by EDS - {df_name}', fontsize=16)
        plt.xlabel(f'{delt_type} (Inverted)', fontsize=12)
        plt.ylabel('ADT Frequency (%)', fontsize=12)

        # Adjust layout and display the plot
        plt.tight_layout()
        plt.show()


def merge_adt_quotation(qtn_cols_prefix, df_adt, df_supplier_quotation, qtn_supplier, qtn_price_column, qtn_name):
    """Merges ADT route data with supplier quotations, handling multiple EDS entries for a single route.

    This function performs a left merge from the ADT routes dataframe (`df_adt`) to the
    supplier's quotation data (`df_supplier_quotation`). It prefixes quotation columns to
    avoid naming conflicts.

    A primary merge (`first_eds`) is always created, taking the first available EDS entry for
    each route key. A second merge (`last_eds`) is created only if the function detects
    that a single route key (`KEY_TO_ADT`) corresponds to multiple, different service
    stations (`EDS_LEMATIZED`). This handles scenarios where suppliers offer alternative
    service stations for the same route.

    Args:
        qtn_cols_prefix (str): A prefix to add to all columns from the quotation dataframe
            to prevent naming collisions after the merge.
        df_adt (pd.DataFrame): DataFrame with ADT route data. Must contain 'KEY_TO_SUPPLIER'.
        df_supplier_quotation (pd.DataFrame): DataFrame with a supplier's quotation data.
            Must contain 'KEY_TO_ADT' and 'EDS_LEMATIZED'.
        qtn_supplier (str): The name of the supplier for the quotation.
        qtn_price_column (str): The name of the price column in the quotation.
        qtn_name (str): A descriptive name for the quotation dataset.

    Returns:
        dict: A dictionary containing the merged DataFrames.
            - 'first_eds': (pd.DataFrame) Merged data using the first EDS entry for each route.
            - 'last_eds': (pd.DataFrame, optional) Merged data using the last EDS entry.
                          Only present if duplicates with different EDS were found.
    """
    adt_shape_before_merge = df_adt.shape
    # Prefix quotation columns to avoid conflicts, preserving the join key.
    df_supplier_quotation.columns = [qtn_cols_prefix + str(c) if c != 'KEY_TO_ADT' else c for c in df_supplier_quotation.columns]
    
    merged_dict = {}

    # Always perform a merge using the 'first' entry for any duplicates in the quotation's route key.
    df_adt_merged_first = pd.merge(
        left=df_adt,
        right=df_supplier_quotation.drop_duplicates(subset=['KEY_TO_ADT'], keep='first'),
        left_on='KEY_TO_SUPPLIER',
        right_on='KEY_TO_ADT',
        how='left',
    )
    # Assign metadata attributes to the merged dataframe
    df_adt_merged_first.qtn_supplier = qtn_supplier
    df_adt_merged_first.qtn_price_column = qtn_price_column
    df_adt_merged_first.qtn_name = qtn_name
    merged_dict['first_eds'] = df_adt_merged_first
    adt_shape_after_merge_first = df_adt_merged_first.shape
    
    # --- Duplicate Handling Logic ---
    # A true duplicate is a route key ('KEY_TO_ADT') that has multiple different service stations ('EDS_LEMATIZED').
    renamed_eds_col = qtn_cols_prefix + 'EDS_LEMATIZED'
    
    # Find all rows that have a duplicated route key.
    duplicated_keys_df = df_supplier_quotation[df_supplier_quotation['KEY_TO_ADT'].duplicated(keep=False)]
    
    key_duplicates_count = 0
    duplicate_key_n_eds = []

    if not duplicated_keys_df.empty:
        # Count unique keys that have duplicates.
        key_duplicates_count = duplicated_keys_df['KEY_TO_ADT'].nunique()
        print(f"Found {key_duplicates_count} route keys with multiple entries. Checking for different service stations...")
        
        # From the duplicated keys, filter down to those that have more than one unique service station.
        duplicate_key_n_eds = duplicated_keys_df.groupby('KEY_TO_ADT').filter(
            lambda x: x[renamed_eds_col].nunique() > 1
        )['KEY_TO_ADT'].unique()

    # If true duplicates (different EDS for the same key) were found, create a second dataframe using the 'last' entry.
    if len(duplicate_key_n_eds) > 0:
        print(f"Found {len(duplicate_key_n_eds)} route keys with different service stations. Creating a second merge using the 'last' entry.")
        df_adt_merged_last = pd.merge(
            left=df_adt,
            right=df_supplier_quotation.drop_duplicates(subset=['KEY_TO_ADT'], keep='last'),
            left_on='KEY_TO_SUPPLIER',
            right_on='KEY_TO_ADT',
            how='left',
        )
        # Assign metadata attributes to the merged dataframe
        df_adt_merged_last.qtn_supplier = qtn_supplier
        df_adt_merged_last.qtn_price_column = qtn_price_column
        df_adt_merged_last.qtn_name = qtn_name
        merged_dict['last_eds'] = df_adt_merged_last
        adt_shape_after_merge_last = df_adt_merged_last.shape
    elif key_duplicates_count > 0:
        # This case handles when keys are duplicated but the service station is the same.
        print("All duplicated route keys point to the same service station. No second merge needed.")

    # --- Final Logging ---
    print(f"Shape of ADT data before merge: {adt_shape_before_merge}")
    print(f"Shape of ADT data after merge (first EDS): {adt_shape_after_merge_first}")
    if 'last_eds' in merged_dict:
        print(f"Shape of ADT data after merge (last EDS): {adt_shape_after_merge_last}")
        print("The following keys had different service stations:")
        print(duplicate_key_n_eds)

    return merged_dict


def merge_adt_all_quotations(df_adt, df_quotations_all_suppliers):
    """Iterates through all supplier quotations and merges them with the ADT route data.

    This function serves as a wrapper to process multiple quotation files. It organizes the
    quotations by supplier and the period they cover, then calls `merge_adt_quotation`
    for each one to perform the actual merge logic. The results are structured in a
    nested dictionary for easy access and analysis.

    Args:
        df_adt (pd.DataFrame): The main ADT route data, containing a 'KEY_TO_SUPPLIER' column
            for joining.
        df_quotations_all_suppliers (list of pd.DataFrame): A list where each DataFrame represents
            a quotation from a supplier. Each DataFrame in the list must have `.supplier`,
            `.price_column`, and `.name` attributes, where the name typically includes the period.

    Returns:
        dict: A nested dictionary structured as {supplier: {period: {merge_results}}}.
              'merge_results' is the dictionary returned by `merge_adt_quotation`,
              containing the 'first_eds' and an optional 'last_eds' DataFrame.
    """
    adt_merged_quot_dict = {}
    # Identify all unique suppliers from the list of quotation dataframes to loop through them.
    unique_suppliers = list({quot.supplier for quot in df_quotations_all_suppliers if hasattr(quot, 'supplier')})

    # Process quotations for each supplier separately.
    for supplier in unique_suppliers:
        print(f"\nProcessing supplier: {supplier}")
        adt_merged_quot_dict[supplier] = {}
        # Iterate through all quotation files to find the ones for the current supplier.
        for df_supplier_quot in df_quotations_all_suppliers:
            if df_supplier_quot.supplier == supplier:
                # --- Prepare for Merge ---
                # Extract metadata from the dataframe's attributes.
                qtn_supplier = df_supplier_quot.supplier
                qtn_price_column = df_supplier_quot.price_column
                qtn_name = df_supplier_quot.name
                # The period is assumed to be the last part of the quotation's name string.
                period = qtn_name.split('_')[-1]
                print(f"Processing object: {qtn_name}")
                
                # Create a copy to ensure the original dataframe in the list remains unmodified.
                df_supplier_quotation = df_supplier_quot.copy()
                
                # Create a unique prefix for this quotation's columns to avoid merge conflicts.
                qtn_cols_prefix = f"{supplier.lower()}_qtn_{period}_"
                
                # --- Perform the Merge ---
                # Call the main merge function with the prepared data.
                merged_dict = merge_adt_quotation(
                    qtn_cols_prefix, 
                    df_adt, 
                    df_supplier_quotation, 
                    qtn_supplier,
                    qtn_price_column,
                    qtn_name
                )

                # Store the results (the dictionary from merge_adt_quotation) in the nested dictionary.
                adt_merged_quot_dict[supplier][period] = merged_dict
                print("-----")
    return adt_merged_quot_dict
