import numpy as np
from numpy import dtype, exp
from numpy.lib.function_base import append
import msoffcrypto
import io
import pandas as pd
import xlsxwriter
import streamlit as st

def main():
    try:
        df_complete, df_gp = import_df_web()
        type = st.radio("Select a data source",["complete data", "gram-positive specimens"])
        msg = "for gram-positive specimens, data has been filtered for first specimens only \nand Enterococcus gallinarum and casseliflavus has also been excluded."
        st.text(msg)
        if type == "complete data":
            complete_data(df_complete)
        else:
            complete_data(df_gp)
    except TypeError or AttributeError:
        pass

def complete_data(df):
    try:
        df = df.set_index("Organism")
        all_organisms = df.index.unique().tolist()
        container = st.container()
        all = st.checkbox("Select all organisms")
        if all:
            organism = container.multiselect("Choose organism(s)", options=all_organisms, default=all_organisms, key="4")
            if not organism:
                container.error("Please select at least one organism or use select all")
        else:
            organism = container.multiselect("Choose organism(s)", options=all_organisms, key="1")
            if not organism:
                container.error("Please select at least one organism")
        df = checkbox(df, organism, "Antibiotic")
        df = prepare_df(df)
        df = checkbox(df, organism, "Antibiotic Interpretation")
        df = prepare_df(df)
        df = checkbox(df, organism, "Specimen Type")
        show_df(df)
        count = df["Specimen ID"].nunique()
        st.metric(label="Count of unique specimen: ", value=count)
    except AttributeError:
        pass

def prepare_df(df):
    df = df.reset_index()
    df = df.set_index("Organism")  
    return df


def checkbox(df,organism, type):
    try:
        df_filtered = df.loc[organism]
        var_list = df_filtered[type].unique().tolist()
        container = st.container()
        all = st.checkbox("Select all " + type)
        if all:
            var = container.multiselect("Choose " + type, options=var_list, default=var_list, key="4")
            if not var:
                container.error("Please select at least one " + type + " or use select all")
        else:
            var = container.multiselect("Choose " + type, options=var_list, key="1")
            if not var:
                container.error("Please select at least one " + type + " or use select all")
        df_filtered = df_filtered.reset_index()
        df_filtered = df_filtered.set_index(type)
        df_filtered = df_filtered.loc[var]
        return df_filtered
    except KeyError:
        pass

def show_df(df):
    df = df.astype(str)
    st.dataframe(df)

def import_df_web():
    df = st.file_uploader("Upload Epic export", type=["xlsx"])
    try:
        df_complete = pd.read_excel(df, "Complete")
        df_complete = df_complete.drop(columns=["Unnamed: 0", "dummy"])
        df_ge = pd.read_excel(df, "First Specimen (Gram Positive)")
        return df_complete, df_ge
    except ValueError or TypeError:
        pass

if __name__ == '__main__':
    main()