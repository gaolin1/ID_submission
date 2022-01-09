import numpy as np
from numpy import dtype, exp
from numpy.lib.function_base import append
import msoffcrypto
import io
import pandas as pd
import xlsxwriter
import streamlit as st
import os
import streamlit.bootstrap
from streamlit import config as _config
import sys
from streamlit import cli as stcli

def main():
    df = import_combined()
    df_complete = pivot(df)
    df_removed = remove_first_specimen(df, df_complete)
    export_df(df, df_removed)
    launch_input = input("Launch dashboard app? (Y/N): ")
    while launch_input not in ("Y", "N"):
        launch_input = input("Please enter either Y or N: ")
    if launch_input == "Y":
        sys.argv = ['streamlit', 'run', 'sub.py']
        sys.exit(stcli.main())
    else:
        quit()

def import_df():
    df = st.file_uploader("Upload Epic export", type=["xlsx"])
    try:
        df = pd.read_excel(df)
        df["Order"] = 1
        return df
    except ValueError:
        pass

def remove_first_specimen(df, df_complete):
    df_staphyl = df_complete[df_complete["Organism"].str.contains("Staphylococcus aureus", na=False)]
    df_staphyl_sorted = df_staphyl.sort_values(by="Received")
    df_staphyl_removed = df_staphyl_sorted.drop_duplicates(subset=["Patient Name"])
    df_entero = df_complete[df_complete["Organism"].str.contains("Enterococcus", na=False)]
    df_entero = df_entero[~df_entero["Organism"].str.contains("gallinarum", na=False)]
    df_entero_removed = df_entero[~df_entero["Organism"].str.contains("casseliflavus", na=False)]
    df_entero_sorted = df_entero_removed.sort_values(by="Received")
    df_entero_ready = df_entero_sorted.drop_duplicates(subset="Patient Name")
    df_removed = df_staphyl_removed.append(df_entero_ready)
    #use the specimen ID of filtered df and back finds the specimen ID and organism
    gp_list = df_removed["Specimen ID"].tolist()
    organism_list = df_removed["Organism"].tolist()
    df = df.set_index("Specimen ID")
    df = df.loc[gp_list]
    df = df.reset_index()
    df = df.set_index("Organism")
    df = df.loc[organism_list]
    return df

def pivot(df):
    df = df.pivot_table(index=("Patient Name", "Organism", "Specimen ID", "Specimen Type", "MRN", "Received"), 
                        columns="Antibiotic", 
                        values="Antibiotic Interpretation", 
                        aggfunc="first"
                        )
    df = df.reset_index()
    df = df.replace('dummy',np.nan)
    df = df.sort_values("Specimen ID")
    #df.drop_duplicates()
    #print(df)
    return df

def find_last_specimen(df):
    df["duplicate"] = df.duplicates(subset=["MRN", "Organism"])
    print(df)
    return df

def sort_by_timestamp(df):
    df = df.sort_values("Received")
    #print(df)
    return df

def import_combined():
    i = 1
    number_of_file = int(input("Enter number of Excel files to be processed: "))
    df = import_df()
    while i < number_of_file:
        df_new = import_df()
        df = df.append(df_new)
        i += 1
        #df.drop_duplicates()
    df = df.set_index("Received")
    df = df.reset_index()
    #to ensure if ' ' will not cause loss of data
    df["Organism"] = df["Organism"].fillna("dummy")
    df["Antibiotic Interpretation"] = df["Antibiotic Interpretation"].fillna("dummy")
    df["Resulted"] = df["Resulted"].fillna("dummy")
    df["dummy"] = np.nan
    df.sort_values(by=["Received"], inplace=True)
    #print(df)
    return df

def import_df():
    decrypted = io.BytesIO()
    read_path = input("Enter Epic Report Output Path: ")
    key = input("Enter File Password: ")
    with open(read_path, "rb") as f:
        file = msoffcrypto.OfficeFile(f)
        file.load_key(password=key)
        file.decrypt(decrypted)
    df = pd.read_excel(decrypted)
    #print(df)
    return df

def export_df(df1, df2):
    write_path = input("Enter location for output xlsx file: ")
    write = pd.ExcelWriter(write_path, engine='xlsxwriter')
    df1.to_excel(write, "Complete")
    df2.to_excel(write, "First Specimen (Gram Positive)")
    write.save()


if __name__ == '__main__':
    main()