"""
A script to convert the Excel data request from 
EERIE into a series of JSON tables
"""

import pandas as pd
import os
import json

homedir = os.getenv('HOME')
# location of input excel file
xlsdir= f'{homedir}/code/dreq_tools/xls'
xl_file = pd.ExcelFile(f"{xlsdir}/EERIE_superset_top60_prior_v1.6.xlsx")

dfs = {sheet_name: xl_file.parse(sheet_name)
          for sheet_name in xl_file.sheet_names}


# convert the excel to csv format and store in TMPDIR (do we need this?)
tmpdir = os.getenv('TMPDIR')
for sheet_name in xl_file.sheet_names:
    dfs[sheet_name].to_csv(f"{tmpdir}/EERIE_{sheet_name}.csv")

del dfs["Notes"]

varkeys_temp={"frequency": "6hrPt", 
            "modeling_realm": "aerosol", 
            "standard_name": "volume_scattering_function_of_radiative_flux_in_air_due_to_ambient_aerosol_particles", 
            "units": "m-1 sr-1", 
            "cell_methods": "area: mean time: point", 
            "cell_measures": "area: areacella", 
            "long_name": "Aerosol Backscatter Coefficient", 
            "comment": "Aerosol  Backscatter at 550nm and 180 degrees, computed from extinction and lidar ratio", 
            "dimensions": "longitude latitude alevel time1 lambda550nm", 
            "out_name": "bs550aer", 
            "type": "real", 
            "positive": ""}
tobeset={
            "valid_min": "", 
            "valid_max": "", 
            "ok_min_mean_abs": "", 
            "ok_max_mean_abs": ""}
varkeys_temp=varkeys_temp.keys()

for sheet_name in dfs.keys():
    if "Long name" in dfs[sheet_name].columns.values:
        dfs[sheet_name]["long_name"]=dfs[sheet_name]["Long name"]
    if "CF Standard Name" in dfs[sheet_name].columns.values:
        dfs[sheet_name]["standard_name"]=dfs[sheet_name]["CF Standard Name"]
    if "CMOR Name" in dfs[sheet_name].columns.values:
        dfs[sheet_name]["out_name"]=dfs[sheet_name]["CMOR Name"]
    for varkey in varkeys_temp:
        if varkey not in list(dfs[sheet_name].columns.values):
            print(sheet_name, varkey)

for sheet_name in dfs.keys():
    counts=dfs[sheet_name]["CMOR Name"].value_counts()
    for var in counts.index :
        if counts[var] != 1:
            print(f"{var}, {sheet_name}")

header={
        "data_specs_version": "", 
        "cmor_version": "", 
        "table_id": "Table", 
        "realm": "", 
        "table_date": "08 May 2023", 
        "missing_value": "1e20", 
        "int_missing_value": "-999", 
        "product": "model-output", 
        "approx_interval": "", 
        "generic_levels": "", 
        "mip_era": "", 
        "Conventions": "EERIE-DMP UGRID CF-1.7 CMIP-6.2"
    }

new_sheets={}
for sheet_name in dfs.keys():
    new_sheets[sheet_name]={}
    new_sheets[sheet_name]["Header"]=header.copy()
    new_sheets[sheet_name]["variable_entry"]={}
    recent_table=dfs[sheet_name]
    for var in recent_table["CMOR Name"].unique():
        new_sheets[sheet_name]["variable_entry"][var]=dict()
        for k in varkeys_temp:
            new_sheets[sheet_name]["variable_entry"][var][k]=recent_table[recent_table["CMOR Name"]==var][k].values[0]
            if pd.isna(new_sheets[sheet_name]["variable_entry"][var][k]):
                new_sheets[sheet_name]["variable_entry"][var][k] = ""
        new_sheets[sheet_name]["variable_entry"][var].update(tobeset)

jsondir= f'{homedir}/code/dreq_tools/Tables'

# make the output directory if it doesn't exist
os.makedirs(jsondir,exist_ok=True)

# now save as json
for sheet_name in new_sheets.keys():
    new_sheets[sheet_name]["Header"]["table_id"]+=f" {sheet_name}"
    try:
        with open(f"{jsondir}/EERIE_{sheet_name}.json","w") as f:
            json.dump(
                new_sheets[sheet_name],
                f,
                indent=4
            )
    except Exception as e:
        print(sheet_name)
        print(e)