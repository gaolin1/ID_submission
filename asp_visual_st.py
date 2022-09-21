from logging import PlaceHolder
from typing import Container
from altair.vegalite.v4.schema.channels import Column
import pandas as pd
from pandas.core import frame
from pandas.io import excel
import streamlit as st
import altair as alt
import numpy as np
from io import BytesIO

def main():
    try:
        excel = st.file_uploader("Upload Epic export", type=["xlsx"])
        type = st.radio("Select either DDD (Defined Daily Dose) or DOT (Days Of Therapy)",["DDD", "DOT"])
        report_type = st.radio("Select type of uploaded report", ["Location", "Department", "Both"])
        if report_type == "Both":
            type_dep = type + " per Department"
            type_loc = type + " per Location"
            df_dep = pd.read_excel(excel, type_dep)
            df_loc = pd.read_excel(excel, type_loc)
            grab_drug_both(df_dep, df_loc, type)
        elif report_type == "Location":
            type_loc = type + " per Location"
            df_loc = pd.read_excel(excel, type_loc)
            grab_drug_one(df_loc, type, type_loc)
        else:
            type_dep = type + " per Department"
            df_loc = pd.read_excel(excel, type_dep)
            grab_drug_one(df_loc, type, type_dep)         
    except ValueError or TypeError or AttributeError:
        pass

def prepare_loc(df, drug_list, loc_list):
    df = df.set_index("GROUPER_NAME")
    df_drug = df.loc[drug_list]
    df_drug = df_drug.reset_index()
    df_drug_loc = df_drug.set_index("LOC_NAME")
    df_drug_loc = df_drug_loc.loc[loc_list]
    return df_drug_loc

def make_chart(data, dep_or_loc, ddd_or_dot, type, df_total):
    data = data.reset_index()
    #to correct for "one-off" error in time axis
    data["MONTH"] = data["MONTH_BEGIN_DT"].dt.strftime('%b %d %Y')
    df_total["MONTH"] = df_total["MONTH_BEGIN_DT"].dt.strftime('%b %d %Y')
    value_type = ddd_or_dot + " " + dep_or_loc + " Per 1000 Patients"
    base = alt.Chart(data).properties(height=450,width=800)
    highlight = alt.selection(type='single', on='mouseover', fields=['symbol'], nearest=True)
    chart = (
        base.mark_line(opacity=0.4)
        .encode(
            x=alt.X("month(MONTH):T", title = "month"),
            y=alt.Y(value_type + ":Q",stack=None, aggregate="sum", title = ddd_or_dot + " per 1000 patient days "),
            color=type+":N",
        )
        .interactive()
        )
    total = (
        alt.Chart(df_total).properties(height=450,width=800).mark_line(color='yellow', opacity=1, strokeDash=[0.2,4])
        .encode(
            x=alt.X("month(MONTH):T", title = "month"),
            y=alt.Y(value_type + ":Q"),
        )
    )

    #take out for now
    #rule = base.mark_rule(color='red', opacity=0.8, strokeDash=[1,1]).transform_window(
    #    cumulative= 'sum(' + value_type + ')',
    #).encode(
    #    y='cumulative:Q',
    #    x=alt.X("month(MONTH_BEGIN_DT):T", title = "month"),
    #    )
    chart = chart + total
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

def grab_drug_both(df, df2, type):
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
    df_loc_for_dep, loc_list = checkbox(df, drug_selected, "LOC_NAME")
    df_loc_for_dep = prepare_df(df_loc_for_dep)
    df_loc = prepare_loc(df2, drug_selected, loc_list)
    df_loc_total = get_total(df_loc, "Location", type)
    loc_chart = make_chart(df_loc, "Location", type, "LOC_NAME", df_loc_total)
    show_df(df_loc)
    download(df_loc,type, " per location export")
    st.altair_chart(loc_chart, use_container_width=True)
    df_loc_dep, dep_list = checkbox(df_loc_for_dep, drug_selected, "DEPARTMENT_NAME")
    df_dep_total = get_total(df_loc_dep, "Department", type)
    dep_chart = make_chart(df_loc_dep, "Department", type, "DEPARTMENT_NAME", df_dep_total)
    show_df(df_loc_dep)
    download(df_loc_dep, type, " per department export")
    #show_df(df_loc_total)
    #show_df(df_dep_total)
    st.altair_chart(dep_chart, use_container_width=True)

def download(df,type, loc_or_dep):
    df_excel = to_excel(df)
    filename = type + loc_or_dep + ".xlsx"
    st.download_button(
        label="Download Above Dataframe To Excel",
        data=df_excel,
        file_name=filename
    )

def to_excel(df):
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='Sheet1')
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']
    format1 = workbook.add_format({'num_format': '0.00'}) 
    worksheet.set_column('A:A', None, format1)  
    writer.save()
    processed_data = output.getvalue()
    return processed_data

def get_total(df, dep_or_loc, ddd_or_dot):
    value = ddd_or_dot + " " + dep_or_loc + " Per 1000 Patients"
    ddd_or_dot_value = ddd_or_dot + "_VALUE"
    df_total_days = df.groupby("MONTH_BEGIN_DT", as_index=False)["Patient Days"].sum()
    df = df.merge(df_total_days, left_on="MONTH_BEGIN_DT", right_on="MONTH_BEGIN_DT").drop("Patient Days_x", axis="columns")
    df[value] = df[ddd_or_dot_value]/df["Patient Days_y"]*1000
    df[value] = df[value].round(2)
    df = df.groupby("MONTH_BEGIN_DT", as_index=False)[value].sum()
    #show_df(df)
    return df

def grab_drug_one(df, ddd_or_dot, type):
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
    if "Location" in type:
        df_loc, loc_list = checkbox(df, drug_selected, "LOC_NAME")
        df_loc = prepare_df(df_loc)
        df_loc_total = get_total(df_loc, "Location", ddd_or_dot)
        loc_chart = make_chart(df_loc, "Location", ddd_or_dot, "LOC_NAME", df_loc_total)
        show_df(df_loc)
        download(df_loc, type, " export")
        st.altair_chart(loc_chart, use_container_width=True)
    else:
        df_dep, dep_list = checkbox(df, drug_selected, "DEPARTMENT_NAME")
        df_dep = prepare_df(df_dep)
        df_dep_total = get_total(df_dep, "Department", ddd_or_dot)
        dep_chart = make_chart(df_dep, "Department", ddd_or_dot, "DEPARTMENT_NAME", df_dep_total)
        show_df(df_dep)
        download(df_dep, type, " export")
        st.altair_chart(dep_chart, use_container_width=True)

def show_df(df):
    df = df.astype(str)
    st.dataframe(df)

if __name__ == '__main__':
    main()