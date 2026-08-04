import streamlit as st
import pandas as pd


# Page settings
st.set_page_config(
    page_title="Steinweg Landside Calculator",
    page_icon="🚢"
)


# Load Excel files

msc_rates = pd.read_excel("MSC_Rates.xlsx")
landside_rates = pd.read_excel("Landside_Rates.xlsx")


# Title

st.title("STEINWEG")
st.subheader("Master Landside Freight Calculator")


st.divider()


# User inputs

shipping_line = st.selectbox(
    "Shipping Line",
    [
        "MSC",
        "Maersk",
        "CMA CGM",
        "Hapag Lloyd"
    ]
)


container = st.selectbox(
    "Container Size",
    [
        "20GP",
        "40GP",
        "40HC"
    ]
)


port = st.selectbox(
    "Port",
    [
        "Durban",
        "Cape Town",
        "Port Elizabeth"
    ]
)


st.divider()


transport = st.number_input(
    "Transport Cost (ZAR)",
    min_value=0
)


clearing = st.number_input(
    "Clearing Cost (ZAR)",
    min_value=0
)


handling = st.number_input(
    "Handling Cost (ZAR)",
    min_value=0
)



if st.button("Calculate"):


    # Find freight rate

    freight = 0


    if shipping_line == "MSC":

        result = msc_rates[
            (msc_rates["Container"] == container)
        ]


        if not result.empty:
            freight = result.iloc[0]["Freight"]



    total = (
        freight +
        transport +
        clearing +
        handling
    )


    st.success("Calculation Complete")


    st.write("### Cost Breakdown")

    st.write(
        f"Ocean Freight: R {freight:,.2f}"
    )

    st.write(
        f"Transport: R {transport:,.2f}"
    )

    st.write(
        f"Clearing: R {clearing:,.2f}"
    )

    st.write(
        f"Handling: R {handling:,.2f}"
    )


    st.divider()


    st.header(
        f"TOTAL: R {total:,.2f}"
    )
