import io
import msoffcrypto
import sys
from streamlit.web import cli as stcli
from tkinter import Tk
from tkinter.filedialog import askopenfilename
from tkinter.filedialog import asksaveasfilename

def main():
    report_input = input("Process Epic Report? (Y/N): ")
    while report_input not in ("Y", "N"):
        report_input = input("Please enter either Y or N: ")
    if report_input == "Y":
        main_type = input("\n1. DDD\n2. DOT\nPlease indicate type of report (1/2) to be processed: ")
        while main_type not in ("1", "2"):
            main_type = input("Please enter either 1 or 2: ")
        if main_type == "1":
            main_type = "DDD"
        else:
            main_type = "DOT"
        sub_type = int(input("\n1. Location\n2. Department\n3. Both\nPlease choose one of the following (1/2/3): "))
        while sub_type not in (1,2,3):
            sub_type = int(input("Please enter 1, 2 or 3: "))
        if sub_type == 1:
            sub_type = "Location"
            import_type_location = main_type + " per " + sub_type
            df_type_location = import_df(import_type_location)
            df_type_location = import_processed(df_type_location, main_type, sub_type)
            export_one_df(df_type_location, import_type_location)
        elif sub_type == 2:
            sub_type = "Department"
            import_type_department = main_type + " per Department"
            df_type_department = import_df(import_type_department)
            df_type_department = import_processed(df_type_department, main_type, sub_type)
            export_one_df(df_type_department, import_type_department)  
        else:
            sub_type = "Location"
            import_type_location = main_type + " per Location"
            df_type_location = import_df(import_type_location)
            df_type_location = import_processed(df_type_location, main_type, sub_type)
            sub_type = "Department"
            import_type_department = main_type + " per Department"
            df_type_department = import_df(import_type_department)
            df_type_department = import_processed(df_type_department, main_type, sub_type)
            export_two_df(df_type_location, import_type_location, df_type_department, import_type_department)
    else:
        pass
    launch_input = input("Launch dashboard app? (Y/N): ")
    while launch_input not in ("Y", "N"):
        launch_input = input("Please enter either Y or N: ")
    if launch_input == "Y":
        sys.argv = ['streamlit', 'run', 'asp_visual_st.py']
        sys.exit(stcli.main())
    else:
        quit()

def import_processed(df, main_type, sub_type):
    df[main_type + " " + sub_type + " Per 1000 Patients"] = df["Per 1000 Patients"]
    df = df.drop(columns=["Per 1000 Patients"])
    df["GROUPER_NAME"] = df["GROUPER_NAME"].str.replace("ERX CONCEPT", "")
    return df

def import_df(type):
    decrypted = io.BytesIO()
    read_path = get_file()
    key = input("Enter File Password: ")
    with open(read_path, "rb") as f:
        file = msoffcrypto.OfficeFile(f)
        file.load_key(password=key)
        file.decrypt(decrypted)
    df = pd.read_excel(decrypted)
    return df

def get_file():
    Tk().withdraw()
    filename = askopenfilename(initialdir = "./", title = "Select file",filetypes = [("Excel Files","*.xlsx")])
    return filename

def export_file():
    Tk().withdraw()
    filename = asksaveasfilename(initialdir = "./", title = "Save file",filetypes = [("Excel Files","*.xlsx")])
    return filename

def export_one_df(df1, type1):
    write_path = export_file()
    write = pd.ExcelWriter(write_path, engine='xlsxwriter')
    df1.to_excel(write, type1, index=False)
    write.save()

def export_two_df(df1, type1, df2, type2):
    write_path = export_file()
    write = pd.ExcelWriter(write_path, engine='xlsxwriter')
    df1.to_excel(write, type1, index=False)
    df2.to_excel(write, type2, index=False)
    write.save()

if __name__ == '__main__':
    main()