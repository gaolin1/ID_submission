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
    st.set_page_config(layout="wide")
    count = 0
    try:
        excel = st.sidebar.file_uploader("Upload Epic export", type=["xlsx"])
        type = st.sidebar.radio("Select either DDD (Defined Daily Dose) or DOT (Days Of Therapy)",["DDD", "DOT"])
        report_type = st.sidebar.radio("Select type of uploaded report", ["Location", "Department", "Both"])
        combine = st.sidebar.checkbox("Combine Data For All Imported Hospitals/Departments?")
        if report_type == "Both":
            type_dep = type + " per Department"
            type_loc = type + " per Location"
            df_dep = pd.read_excel(excel, type_dep)
            df_loc = pd.read_excel(excel, type_loc)
            count = grab_drug_both(df_dep, df_loc, type, count, combine)
        elif report_type == "Location":
            type_loc = type + " per Location"
            df_loc = pd.read_excel(excel, type_loc)
            count = grab_drug_one(df_loc, type, type_loc, count, combine)
        else:
            type_dep = type + " per Department"
            df_dep = pd.read_excel(excel, type_dep)
            count = grab_drug_one(df_dep, type, type_dep, count, combine)         
    except ValueError or TypeError or AttributeError:
        pass

def prepare_loc(df, drug_list, loc_list, loc_or_dep="loc"):
    df_drug = df.reset_index()
    df_drug = df_drug.set_index("GROUPER_NAME")
    df_drug = df_drug.loc[drug_list]
    df_drug = df_drug.reset_index()
    if loc_or_dep == "loc":
        df_drug_loc = df_drug.set_index("LOC_NAME")
    else:
        df_drug_loc = df_drug.set_index("DEPARTMENT_NAME")
    df_drug_loc = df_drug_loc.loc[loc_list]
    return df_drug_loc

def make_chart(data, dep_or_loc, ddd_or_dot, type, df_total, count, drug_list, combine=False):
    data = data.reset_index()
    #to correct for "one-off" error in time axis
    data["MONTH"] = data["MONTH_BEGIN_DT"].dt.strftime('%b %d %Y')
    df_total["MONTH"] = df_total["MONTH_BEGIN_DT"].dt.strftime('%b %d %Y')
    value_type = ddd_or_dot + " " + dep_or_loc + " Per 1000 Patients"
    if combine == True:
        data["Legend"] = data["GROUPER_NAME"] + " - HHS"
        data["Combined"] = "Combined"
        data = data.groupby(["MONTH", "Legend","GROUPER_NAME","Combined"], as_index=False)[value_type].sum()
    else:
        data["Legend"] = data["GROUPER_NAME"] + " - " + data[type]
    data = data.set_index("Legend")
    data = data.reset_index()
    base = alt.Chart(data).properties(height=800)
    brush = alt.selection(type='interval', encodings=['x'])
    #show_df(data)
    if dep_or_loc == "Department":
        field_name = "DEPARTMENT_NAME"
    elif combine == True:
        field_name = "Combined"
    else:
        field_name = "LOC_NAME"
    df_total["Weighted Average"]="Weighted Average"
    total = (
        alt.Chart(df_total).mark_line(color='red', opacity=1, strokeDash=[5,5],strokeWidth=5)
        .encode(
            x=alt.X("yearmonth(MONTH):T"),
            y=alt.Y("average:Q"),
            color=alt.Color("Weighted Average:N",legend=alt.Legend(orient="top",title="Legend"),
                            scale=alt.Scale(range=["red"]))
        )
    )
    
    chart_circle = (
        base.mark_circle(size=70)
        .encode(
            x=alt.X("yearmonth(MONTH):T", title = "month",axis=alt.Axis(tickCount="month")),
            y=alt.Y(value_type + ":Q",stack=None, aggregate="sum", title = ddd_or_dot + " per 1000 patient days ",axis=alt.Axis(format="~s")),
            color=alt.Color("Legend:N",legend=alt.Legend(titleLimit=600, labelLimit=400, direction="vertical", orient="top-right",
                                                                                labelOpacity=0.7,
                                                                                titleOpacity=1,
                                                                                symbolOpacity=0.7,
                                                                                title="Legend 'Antimicrobial - Unit'",
                                                                                titleFontSize=13,
                                                                                labelFontSize=11,
                                                                                symbolSize=8),
                                                                                scale=alt.Scale(scheme="tableau20")),
            tooltip=[alt.Tooltip(field='GROUPER_NAME'),alt.Tooltip(field=field_name),alt.Tooltip(field=value_type),alt.Tooltip(field="MONTH",type="temporal")]
        )
        .add_selection(
            brush
        )
        )

    chart_line = (
        base.mark_line(opacity=0.8, strokeWidth=4)
        .encode(
            x=alt.X("yearmonth(MONTH):T", title = "month"),
            y=alt.Y(value_type + ":Q",stack=None, aggregate="sum"),
            color=alt.Color("Legend:N", legend=None))
        )
    chart = chart_circle + chart_line
    if len(drug_list) > 1 or count > 1:
        chart = alt.layer(chart, total, data=data).resolve_scale(color="independent")
    else:
        pass
    chart.configure_axis(title="title")
    st.altair_chart(chart, use_container_width=True)


def count_selected(data, type=""):
    data = data.reset_index()
    if "DEPARTMENT" in type:
        count = data["DEPARTMENT_NAME"].drop_duplicates().count()
    else:
        count = data["LOC_NAME"].drop_duplicates().count()
    return count

def prepare_df(df):
    df = df.reset_index()
    df = df.set_index("GROUPER_NAME")  
    return df

def checkbox(df, drug, type, count):
    try:
        df_filtered = df.loc[drug]
        var_list = df_filtered[type].unique().tolist()
        container = st.container()
        all = st.checkbox("Select all " + type)
        if all:
            var = container.multiselect("Choose " + type, options=var_list, default=var_list, key=count)
            count += 1
            if not var:
                container.error("Please select at least one " + type + " or use select all")
        else:
            var = container.multiselect("Choose " + type, options=var_list, key=count)
            count +=1
            if not var:
                container.error("Please select at least one " + type + " or use select all")
        df_filtered = df_filtered.reset_index()
        df_filtered = df_filtered.set_index(type)
        df_filtered = df_filtered.loc[var]
        return df_filtered, var, count
    except KeyError:
        pass

def grab_drug_both(df, df2, type, count, combine):
    df = df.set_index("GROUPER_NAME")
    drug_list = df.index.unique().tolist()
    container = st.container()
    all = st.checkbox("Select all medication groupers")
    if all:
        drug_selected = container.multiselect("Choose grouper", options=drug_list, default=drug_list, key=count)
        count += 1
        if not drug_selected:
            container.error("Please select at least one grouper or use select all")
    else:
        drug_selected = container.multiselect("Choose grouper", options=drug_list, key=count)
        count += 1
        if not drug_selected:
            container.error("Please select at least one grouper")
    #st.text(drug_selected)
    if combine == True:
        drug_combined = prep_combine_total(df2, type, "Location", drug_selected)
        df_combined_total = get_total(drug_combined, "Location", type, combine)
        make_chart(drug_combined, "Location", type, "LOC_NAME", df_combined_total, 0, drug_selected, combine)
    else:
        pass
    df_loc_for_dep, loc_list, count = checkbox(df, drug_selected, "LOC_NAME", count)
    df_loc_for_dep = prepare_df(df_loc_for_dep)
    df_loc = prepare_loc(df2, drug_selected, loc_list)
    df_loc_total = get_total(df_loc, "Location", type)
    show_df(df_loc)
    show_df(df_loc_total, "total", "loc")
    download(df_loc,type, " per location export")
    sel_count = count_selected(df_loc, "LOC")
    make_chart(df_loc, "Location", type, "LOC_NAME", df_loc_total, sel_count, drug_list)
    df_loc_dep, dep_list, count = checkbox(df_loc_for_dep, drug_selected, "DEPARTMENT_NAME", count)
    df_dep_total = get_total(df_loc_dep, "Department", type)
    show_df(df_loc_dep)
    show_df(df_dep_total, "total", "dep")
    download(df_loc_dep, type, " per department export")
    #show_df(df_loc_total)
    #show_df(df_dep_total)
    sel_dep_count = count_selected(df_loc_dep, "DEPARTMENT")
    make_chart(df_loc_dep, "Department", type, "DEPARTMENT_NAME", df_dep_total, sel_dep_count, drug_list)
    return count

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

def get_total(df, dep_or_loc, ddd_or_dot, combined=False):
    value = ddd_or_dot + " " + dep_or_loc + " Per 1000 Patients"
    ddd_or_dot_value = ddd_or_dot + "_VALUE"
    df = df.reset_index()
    if dep_or_loc == "Location":
        finder = "LOC_NAME"
    else:
        finder = "DEPARTMENT_NAME"
    if combined == True:
        df = df.groupby(["MONTH_BEGIN_DT","GROUPER_NAME"], as_index=False)["Patient Days",value,ddd_or_dot_value].sum()
    else:
        df = df.groupby(["MONTH_BEGIN_DT","GROUPER_NAME",finder], as_index=False)["Patient Days",value,ddd_or_dot_value].sum()
    df_total_days = df.groupby(["MONTH_BEGIN_DT"], as_index=False)["Patient Days"].sum()
    df = df.merge(df_total_days, left_on="MONTH_BEGIN_DT", right_on="MONTH_BEGIN_DT")
    df["weight"] = df["Patient Days_x"]/df["Patient Days_y"]
    df["average"] = df[value]*df["weight"]
    df = df.groupby("MONTH_BEGIN_DT", as_index=False)["average"].sum().round(2)
    #show_df(df)
    return df

def get_combined(df, dep_or_loc, ddd_or_dot, drug_list):
    value = ddd_or_dot + " " + dep_or_loc + " Per 1000 Patients"
    ddd_or_dot_value = ddd_or_dot + "_VALUE"
    df = prepare_df(df)
    df_combined = df.loc[drug_list]
    df_combined = df_combined.groupby(["MONTH_BEGIN_DT","GROUPER_NAME","Patient Days",ddd_or_dot_value], as_index=True)[value].sum().round(2)
    df_combined = prepare_df(df_combined)
    return df_combined

def grab_drug_one(df, ddd_or_dot, type, count, combine=False):
    df = df.set_index("GROUPER_NAME")
    drug_list = df.index.unique().tolist()
    container = st.container()
    all = st.checkbox("Select all medication groupers")
    if all:
        drug_selected = container.multiselect("Choose grouper", options=drug_list, default=drug_list, key=count)
        count += 1
        if not drug_selected:
            container.error("Please select at least one grouper or use select all")
    else:
        drug_selected = container.multiselect("Choose grouper", options=drug_list, key=count)
        count += 1
        if not drug_selected:
            container.error("Please select at least one grouper")
    if "Location" in type:
        if combine == True:
            drug_combined = prep_combine_total(df, ddd_or_dot, "Location", drug_selected)
            df_combined_total = get_total(drug_combined, "Location", ddd_or_dot, combine)
            #show_df(drug_combined)
            make_chart(drug_combined, "Location", ddd_or_dot, "LOC_NAME", df_combined_total, 0, drug_selected, combine)
        else:
            pass
        df_loc, loc_list, count = checkbox(df, drug_selected, "LOC_NAME", count)
        df_loc = prepare_df(df_loc)
        df_loc_total = get_total(df_loc, "Location", ddd_or_dot)
        show_df(df_loc)
        show_df(df_loc_total, "total")
        download(df_loc, type, " export")
        sel_count = count_selected(df_loc, "LOC")
        make_chart(df_loc, "Location", ddd_or_dot, "LOC_NAME", df_loc_total, sel_count, drug_selected)
    else:
        if combine == True:
            df_combined = prep_combine_total(df, ddd_or_dot, "Department", drug_selected)
            df_combined_total = get_total(df_combined, "Department", ddd_or_dot, combine)
            #show_df(df_combined)
            #show_df(df_combined_total)
            make_chart(df_combined, "Department", ddd_or_dot, "DEPARTMENT_NAME", df_combined_total, 0, drug_selected, combine)
        else:
            pass
        df_dep, dep_list, count= checkbox(df, drug_selected, "DEPARTMENT_NAME", count)
        df_dep = prepare_df(df_dep)
        df_dep_total = get_total(df_dep, "Department", ddd_or_dot)
        show_df(df_dep)
        show_df(df_dep_total, "total")
        download(df_dep, type, " export")
        sel_count = count_selected(df_dep, "DEPARTMENT")
        make_chart(df_dep, "Department", ddd_or_dot, "DEPARTMENT_NAME", df_dep_total, sel_count, drug_selected)
    return count

def prep_combine_total(df, ddd_or_dot, loc_or_dep, drug_selected):
    if ddd_or_dot == "DOT":
        if loc_or_dep == "Department":
            name = "DEPARTMENT_NAME"
            prep_argu = "dep"
            locator = "DOT Department Per 1000 Patients"
        else:
            name = "LOC_NAME"
            prep_argu = "loc"
            locator = "DOT Location Per 1000 Patients"
        locator_value = "DOT_VALUE"
    else:
        if loc_or_dep == "Department":
            name = "DEPARTMENT_NAME"
            prep_argu = "dep"
            locator = "DDD Department Per 1000 Patients"
        else:
            name = "LOC_NAME"
            prep_argu = "loc"
            locator = "DDD Location Per 1000 Patients"
        locator_value = "DDD_VALUE"
    df = df.reset_index()
    df = df.set_index("GROUPER_NAME")
    df = df.loc[drug_selected]
    all_list = df[name].unique().tolist()
    df = prepare_loc(df, drug_selected, all_list, prep_argu)
    df = df.reset_index()
    #show_df(df)
    df = df.groupby(["MONTH_BEGIN_DT","GROUPER_NAME"], as_index=False)["Patient Days",locator_value,locator].sum()
    df[locator] = df[locator_value]/df["Patient Days"]*1000
    return df

def show_df(df, type="", loc_or_dep=""):
    df = df.astype(str)
    if type=="total":
        if loc_or_dep == "loc":
            st.text("Weighted Total By Month (Hospital)")
        elif loc_or_dep =="dep":
            st.text("Weighted Total By Month (Department)")
        else:
            st.text("Weighted Total By Month")
    else:
        st.text("Selected Data")
    st.dataframe(df)

if __name__ == '__main__':
    main()