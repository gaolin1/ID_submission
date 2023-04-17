import pandas as pd
import io
import msoffcrypto
import sys
import PySimpleGUI as sg
import xlsxwriter
from streamlit import config as _config
import sys
from streamlit.web import cli as stcli

def main():
    ##general utility functions
    #returns the selected variable per string value
    def get_selected_var(window, list, initial_value, rel_pos):
        var_value = window[initial_value].TKIntVar.get()
        var_selected = list[int(var_value/1000%100)-rel_pos] if var_value else None    
        #return var_value #(for troubleshooting)    
        return var_selected

    #main function: takes read path of the file and password and returns the dataframe
    def import_df(read_path, key):
        decrypted = io.BytesIO()
        #read_path = get_file()
        #key = input("\nEnter File Password: ")
        with open(read_path, "rb") as f:
            file = msoffcrypto.OfficeFile(f)
            file.load_key(password=key)
            file.decrypt(decrypted)
        df = pd.read_excel(decrypted)
        #print(df)
        return df

    def import_processed(df, main_type, sub_type):
        df[main_type + " " + sub_type + " Per 1000 Patients"] = df["Per 1000 Patients"]
        df = df.drop(columns=["Per 1000 Patients"])
        df["GROUPER_NAME"] = df["GROUPER_NAME"].str.replace("CONCEPT", "")
        df["GROUPER_NAME"] = df["GROUPER_NAME"].str.replace("HHS", "")
        df["GROUPER_NAME"] = df["GROUPER_NAME"].str.replace("GENERAL", "")
        df["GROUPER_NAME"] = df["GROUPER_NAME"].str.replace("ERX", "")
        if main_type == "DOT":
            value = "DOT_VALUE"
        else:
            value = "DDD_VALUE"
        df_loc = df.groupby(["MONTH_BEGIN_DT", "GROUPER_NAME", "SERVICE_AREA", "LOC_NAME"], as_index=False)[value, "Patient Days"].sum()
        df_loc[main_type + " Location Per 1000 Patients"] = df_loc[value]/df_loc["Patient Days"]*1000
        return df, df_loc

    def export_two_df(df1, type1, df2, type2, write_path):
        write = pd.ExcelWriter(write_path, engine='xlsxwriter')
        df1.to_excel(write, type1, index=False)
        df2.to_excel(write, type2, index=False)
        write.save()

    #global lists
    report_choices = [
        "DDD",
        "DOT"
    ]

    sg.SetOptions(font="any 16", background_color="black", text_color="yellow")

    def main_window():
        layout = [
            [sg.Text("Choose Report Type", background_color="black")],
            [[sg.Radio(text, "TYPE", enable_events=True, key=f"TYPE {i}", background_color="black")] for i, text in enumerate(report_choices)],
            [sg.Text("Choose one of the following to process Epic output file", background_color="black")],
            [sg.Button("Process Epic Export File")],
            [sg.Text("   " ,background_color="black")],
            [sg.Button("Launch Analytics Dashboard")],
            [sg.Text("   " ,background_color="black")],
            [sg.Exit()]
        ]
        window = sg.Window("HHS ASP DDD/DOT Analytics", layout=layout, finalize=True)
        return window       


    def single_process_window(report_type,level_type):
        layout = [
            [sg.Text("Choose the Epic output file", background_color="black")],
            [sg.InputText(key="-SINGLE_FILE_PATH-"),
            sg.FileBrowse(initial_folder="./", file_types=(["Excel files", "*.xlsx"],))],
            [sg.Text("Enter file password below", background_color="black")],
            [sg.InputText(key="-SINGLE_PWD-", password_char="*")],
            [sg.Button("Import and process")],
            [sg.Text("   " ,background_color="black")],
            [sg.Text("Save Processed File", background_color="black")],
            [sg.InputText(key="-SAVE_PATH-", enable_events=True),
            sg.FileSaveAs(initial_folder="./", file_types=(("Excel files", "*.xlsx"),), default_extension="*.xlsx")],
            [sg.Button("Export"), sg.Exit()]
        ]
        window = sg.Window("Process " + report_type + " " + level_type + " Epic Output File", layout=layout, finalize=True)
        return window     

    window1, window2 = main_window(), None

    while True:
        window, event, values = sg.read_all_windows()
        if event in (sg.WIN_CLOSED, 'Exit'):
            window.close()
            if window == window2:       # if closing win 2, mark as closed
                window2 = None
            elif window == window1:     # if closing win 1, exit program
                break
        if event == "Process Epic Export File" and not window2:
            level_type = event
            report_type = get_selected_var(window, report_choices, "TYPE 0", 3)
            window2 = single_process_window(report_type, level_type)
        if event == "Import and process":
            df = import_df(values["-SINGLE_FILE_PATH-"], values["-SINGLE_PWD-"])
            df_dep, df_loc = import_processed(df, report_type, "Department")
            sg.popup("Import and processing successful", background_color="black")
        if event == "Export":
            combined_loc_type = report_type + " per Location"
            combined_dep_type = report_type + " per Department"
            write_path = values["-SAVE_PATH-"]
            export_two_df(df_loc, combined_loc_type, df_dep, combined_dep_type, write_path)
            sg.popup("File save sucessful, saved under " + write_path, background_color="black")
            window2.close()
        if event == "Launch Analytics Dashboard":
            print("Press Control + C to terminate the program")
            sys.argv = ['streamlit', 'run', 'asp_visual_st.py']
            sys.exit(stcli.main())



    window.close()

if __name__ == '__main__':
    main()