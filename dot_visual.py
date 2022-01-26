import pandas as pd
import io
import msoffcrypto
import sys
from streamlit import cli as stcli

def main():
    main_type = input("Please indicate type of report (DDD/DOT): ")
    while main_type not in ("DDD", "DOT"):
        main_type = input("Please enter either DDD or DOT: ")
    if main_type == "DDD":
        df_ddd_department = import_df("DDD per Department")
        df_ddd_department = import_processed(df_ddd_department, "DDD Department")
        df_ddd_location = import_df("DDD per Location")
        df_ddd_location = import_processed(df_ddd_location, "DDD Location")
        export_df(df_ddd_department, "DDD Department", df_ddd_location, "DDD Location")
    else:
        pass
    launch_input = input("Launch dashboard app? (Y/N): ")
    while launch_input not in ("Y", "N"):
        launch_input = input("Please enter either Y or N: ")
    if launch_input == "Y":
        sys.argv = ['streamlit', 'run', 'dot_visual_st.py']
        sys.exit(stcli.main())
    else:
        quit()

def import_processed(df, type):
    df[type + " Per 1000 Patients"] = df["Per 1000 Patients"]
    df = df.drop(columns=["Per 1000 Patients"])
    df["GROUPER_NAME"] = df["GROUPER_NAME"].str.replace("ERX CONCEPT", "")
    return df

def import_df(type):
    decrypted = io.BytesIO()
    read_path = input("Enter Epic " + type + " Report Output Path: ")
    key = input("Enter File Password: ")
    with open(read_path, "rb") as f:
        file = msoffcrypto.OfficeFile(f)
        file.load_key(password=key)
        file.decrypt(decrypted)
    df = pd.read_excel(decrypted)
    return df

def export_df(df1, type1, df2, type2):
    write_path = input("Enter location for output xlsx file: ")
    write = pd.ExcelWriter(write_path, engine='xlsxwriter')
    df1.to_excel(write, type1, index=False)
    df2.to_excel(write, type2, index=False)
    write.save()

if __name__ == '__main__':
    main()