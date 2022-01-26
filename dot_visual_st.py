import pandas as pd
from pandas.io import excel
import streamlit as st
import altair as alt

def main():
    try:
        excel = st.file_uploader("Upload Epic export", type=["xlsx"])
        type = st.radio("Select either DDD (Defined Daily Dose) or DOT (Days Of Therapy)",["DDD", "DOT"])
        type_dep = type + " Department"
        type_loc = type + " Location"
        df_dep = pd.read_excel(excel, type_dep)
        df_loc = pd.read_excel(excel, type_loc)
        grab_drug(df_dep, df_loc, type)
    except ValueError or TypeError or AttributeError:
        pass

def prepare_loc(df, drug_list, loc_list):
    df = df.set_index("GROUPER_NAME")
    df_drug = df.loc[drug_list]
    df_drug = df_drug.reset_index()
    df_drug_loc = df_drug.set_index("LOC_NAME")
    df_drug_loc = df_drug_loc.loc[loc_list]
    return df_drug_loc

def make_chart(data, dep_or_loc, ddd_or_dot, type):
    data = data.reset_index()
    value_type = ddd_or_dot + " " + dep_or_loc + " Per 1000 Patients"
    chart = (
        alt.Chart(data)
        .mark_line(opacity=0.4)
        .encode(
            x="MONTH_BEGIN_DT:T",
            y=alt.Y(value_type + ":Q",stack=None, aggregate="sum"),
            color=type+":N"
        )
        .interactive()
        )
    return chart

def prepare_df(df):
    df = df.reset_index()
    df = df.set_index("GROUPER_NAME")  
    return df

def checkbox(df, drug, type):
    try:
        df_filtered = df.loc[drug]
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
        return df_filtered, var
    except KeyError:
        pass

def grab_drug(df, df2, type):
    df = df.set_index("GROUPER_NAME")
    drug_list = df.index.unique().tolist()
    container = st.container()
    all = st.checkbox("Select all medication groupers")
    if all:
        drug_selected = container.multiselect("Choose grouper", options=drug_list, default=drug_list, key="4")
        if not drug_selected:
            container.error("Please select at least one grouper or use select all")
    else:
        drug_selected = container.multiselect("Choose grouper", options=drug_list, key="1")
        if not drug_selected:
            container.error("Please select at least one grouper")
    #st.text(drug_selected)
    df_loc, loc_list = checkbox(df, drug_selected, "LOC_NAME")
    df_loc = prepare_df(df_loc)
    df_loc_dep, dep_list = checkbox(df_loc, drug_selected, "DEPARTMENT_NAME")
    df_loc = prepare_loc(df2, drug_selected, loc_list)
    dep_chart = make_chart(df_loc_dep, "Department", type, "DEPARTMENT_NAME")
    loc_chart = make_chart(df_loc, "Location", type,"LOC_NAME")
    show_df(df_loc)
    show_df(df_loc_dep)
    st.altair_chart(loc_chart, use_container_width=True)
    st.altair_chart(dep_chart, use_container_width=True)

def show_df(df):
    df = df.astype(str)
    st.dataframe(df)

if __name__ == '__main__':
    main()