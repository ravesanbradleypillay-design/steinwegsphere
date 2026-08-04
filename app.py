import streamlit as st
import pandas as pd
import os


st.set_page_config(
    page_title="Steinweg Freight Calculator",
    page_icon="🚢",
    layout="wide"
)


st.title("STEINWEG")
st.subheader("Master Freight & Transport Calculator")


# EXACT FILE NAMES FROM GITHUB

freight_file = "freight rates.xlsx"
transport_file = "transport rates.xlsx"



# SHOW FILES STREAMLIT CAN SEE

st.write("Files detected by Streamlit:")

st.write(
    os.listdir()
)



# CHECK FILES

if freight_file not in os.listdir():

    st.error(
        "Cannot find freight rates.xlsx"
    )

    st.stop()



if transport_file not in os.listdir():

    st.error(
        "Cannot find transport rates.xlsx"
    )

    st.stop()



# LOAD FILES

@st.cache_data
def load_files():

    freight = pd.read_excel(
        freight_file
    )

    transport = pd.read_excel(
        transport_file
    )

    return freight, transport



freight, transport = load_files()



# CLEAN HEADERS

freight.columns = (
    freight.columns
    .astype(str)
    .str.strip()
)


transport.columns = (
    transport.columns
    .astype(str)
    .str.strip()
)



# PREVIEW

with st.expander("View Excel Data"):

    st.write("Freight Rates")

    st.dataframe(
        freight.head()
    )


    st.write("Transport Rates")

    st.dataframe(
        transport.head()
    )



st.divider()


st.header("Shipment Details")


pod = st.selectbox(
    "POD",
    sorted(
        freight["POD"]
        .dropna()
        .unique()
    )
)



equipment = st.selectbox(
    "Equipment Type",
    sorted(
        freight["EQUIP TYPE"]
        .dropna()
        .unique()
    )
)



payment = st.selectbox(
    "Prepaid / Collect",
    sorted(
        freight["PREPAID / COLLECT"]
        .dropna()
        .unique()
    )
)



qty = st.number_input(
    "Number of Containers",
    min_value=1,
    value=1
)



st.divider()


zone = st.selectbox(
    "Transport Zone",
    sorted(
        transport["ZONE"]
        .dropna()
        .unique()
    )
)



if st.button("Calculate"):


    freight_match = freight[

        (freight["POD"] == pod)

        &

        (freight["EQUIP TYPE"] == equipment)

        &

        (freight["PREPAID / COLLECT"] == payment)

    ]



    if freight_match.empty:

        freight_cost = 0

        st.warning(
            "No freight rate found"
        )

    else:

        freight_cost = float(
            freight_match.iloc[0]["ALL IN RATE"]
        )



    transport_match = transport[

        transport["ZONE"] == zone

    ]



    if transport_match.empty:

        transport_cost = 0

        st.warning(
            "No transport rate found"
        )

    else:

        transport_cost = float(
            transport_match.iloc[0]["TOTAL CHARGE"]
        )



    total = (
        (freight_cost + transport_cost)
        *
        qty
    )



    st.success(
        "Calculation Complete"
    )


    st.write(
        f"Ocean Freight: R {freight_cost:,.2f}"
    )


    st.write(
        f"Transport: R {transport_cost:,.2f}"
    )


    st.header(
        f"TOTAL: R {total:,.2f}"
    )
