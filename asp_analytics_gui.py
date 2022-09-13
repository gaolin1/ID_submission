import pandas as pd
import io
import msoffcrypto
import sys
import PySimpleGUI as sg
import xlsxwriter
import streamlit.bootstrap
from streamlit import config as _config
import sys
from streamlit import cli as stcli

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
        df["GROUPER_NAME"] = df["GROUPER_NAME"].str.replace("ERX CONCEPT", "")
        return df

    def export_one_df(df, type, write_path):
        write = pd.ExcelWriter(write_path, engine='xlsxwriter')
        df.to_excel(write, type, index=False)
        write.save()

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
            [sg.Button("Location"), sg.Button("Department"), sg.Button("Both")],
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
            [sg.Button("Export"), sg.Button("Cancel")]
        ]
        window = sg.Window("Process " + report_type + " " + level_type + " Epic Output File", layout=layout, finalize=True)
        return window

    def double_process_window(report_type):
        layout = [
            [sg.Text("Choose Location output file", background_color="black")],
            [sg.InputText(key="-LOC_FILE_PATH-"),
            sg.FileBrowse(initial_folder="./", file_types=(["Excel files", "*.xlsx"],))],
            [sg.Text("Enter file password below", background_color="black")],
            [sg.InputText(key="-LOC_PWD-", password_char="*")],
            [sg.Text("Choose Department output file", background_color="black")],
            [sg.InputText(key="-DEP_FILE_PATH-"),
            sg.FileBrowse(initial_folder="./", file_types=(["Excel files", "*.xlsx"],))],
            [sg.Text("Enter file password below", background_color="black")],
            [sg.InputText(key="-DEP_PWD-", password_char="*")],
            [sg.Button("Import and process (both)")],
            [sg.Text("   " ,background_color="black")],
            [sg.Text("Save Processed File", background_color="black")],
            [sg.InputText(key="-SAVE_PATH-", enable_events=True),
            sg.FileSaveAs(initial_folder="./", file_types=(("Excel files", "*.xlsx"),), default_extension="*.xlsx")],
            [sg.Button("Export (both)"), sg.Button("Cancel")]
        ]
        window = sg.Window("Process " + report_type + " location and department Epic Output File", layout=layout, finalize=True)
        return window        

    window1, window2 = main_window(), None

    while True:
        window, event, values = sg.read_all_windows()
        if event in (sg.WIN_CLOSED, 'Exit', 'Cancel'):
            window.close()
            if window == window2:       # if closing win 2, mark as closed
                window2 = None
            elif window == window1:     # if closing win 1, exit program
                break
        if event in ("Location", "Department") and not window2:
            level_type = event
            report_type = get_selected_var(window, report_choices, "TYPE 0", 3)
            window2 = single_process_window(report_type, level_type)
        if event == "Import and process":
            df = import_df(values["-SINGLE_FILE_PATH-"], values["-SINGLE_PWD-"])
            df = import_processed(df, report_type, level_type)
            sg.popup("Import and processing successful", background_color="black")
        if event == "Export":
            combined_type = report_type + " per " + level_type
            write_path = values["-SAVE_PATH-"]
            export_one_df(df, combined_type, write_path)
            sg.popup("File save sucessful, saved under " + write_path, background_color="black")
            window.close()
        if event == "Both" and not window2:
            report_type = get_selected_var(window, report_choices, "TYPE 0", 3)
            window2 = double_process_window(report_type)
        if event == "Import and process (both)":
            df_loc = import_df(values["-LOC_FILE_PATH-"], values["-LOC_PWD-"])
            df_dep = import_df(values["-DEP_FILE_PATH-"], values["-DEP_PWD-"])
            df_loc = import_processed(df_loc, report_type, "Location")
            df_dep = import_processed(df_dep, report_type, "Department")
            sg.popup("Import and processing successful", background_color="black")
        if event == "Export (both)":
            combined_loc_type = report_type + " per Location"
            combined_dep_type = report_type + " Department"
            write_path = values["-SAVE_PATH-"]
            export_two_df(df_loc, combined_loc_type, df_dep, combined_dep_type, write_path)
            sg.popup("File save sucessful, saved under " + write_path, background_color="black")
            window.close()
        if event == "Launch Analytics Dashboard":
            sys.argv = ['streamlit', 'run', 'asp_visual_st.py']
            sys.exit(stcli.main())



    window.close()

if __name__ == '__main__':
    main()