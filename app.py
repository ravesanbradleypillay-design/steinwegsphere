import streamlit as st
import pandas as pd
import os


# =========================
# PAGE SETTINGS
# =========================

st.set_page_config(
    page_title="Steinweg Freight Calculator",
    page_icon="🚢",
    layout="wide"
)


st.title("STEINWEG")
st.subheader("Master Freight & Transport Calculator")


# =========================
# FILE NAMES
# =========================

freight_file = "freight rates.xlsx"
transport_file = "transport rates.xls"



# =========================
# CHECK FILES
# =========================

files = os.listdir()

st.write("Files detected:")
st.write(files)


if freight_file not in files:

    st.error("Missing freight rates.xlsx")

    st.stop()


if transport_file not in files:

    st.error("Missing transport rates.xls")

    st.stop()



# =========================
# LOAD DATA
# =========================

@st.cache_data
def load_data():

    freight = pd.read_excel(
        freight_file,
        engine="openpyxl"
    )


    transport = pd.read_excel(
        transport_file,
        engine="xlrd"
    )


    return freight, transport



freight, transport = load_data()



# Clean column names

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



# =========================
# SHOW DATA
# =========================

with st.expander("View Freight Database"):

    st.dataframe(
        freight.head(10)
    )


with st.expander("View Transport Database"):

    st.dataframe(
        transport.head(10)
    )



st.divider()



# =========================
# SHIPMENT INPUTS
# =========================

st.header("Shipment Details")


col1, col2, col3 = st.columns(3)


with col1:

    pod = st.selectbox(
        "POD",
        sorted(
            freight["POD"]
            .dropna()
            .unique()
        )
    )


with col2:

    equipment = st.selectbox(
        "Equipment Type",
        sorted(
            freight["EQUIP TYPE"]
            .dropna()
            .unique()
        )
    )


with col3:

    payment = st.selectbox(
        "Prepaid / Collect",
        sorted(
            freight["PREPAID / COLLECT"]
            .dropna()
            .unique()
        )
    )



quantity = st.number_input(
    "Number of Containers",
    min_value=1,
    value=1
)



st.divider()



# =========================
# TRANSPORT
# =========================

st.header("Transport")


zone = st.selectbox(
    "Transport Zone",
    sorted(
        transport["ZONE"]
        .dropna()
        .unique()
    )
)



# =========================
# CALCULATE
# =========================

if st.button("Calculate Quote"):


    # Freight lookup

    freight_match = freight[

        (freight["POD"] == pod)

        &

        (freight["EQUIP TYPE"] == equipment)

        &

        (freight["PREPAID / COLLECT"] == payment)

    ]


    if freight_match.empty:

        ocean_rate = 0

        st.warning(
            "No freight rate found"
        )

    else:

        ocean_rate = float(
            freight_match.iloc[0]["ALL IN RATE"]
        )



    # Transport lookup

    transport_match = transport[

        transport["ZONE"] == zone

    ]


    if transport_match.empty:

        transport_rate = 0

        st.warning(
            "No transport rate found"
        )

    else:

        transport_rate = float(
            transport_match.iloc[0]["TOTAL CHARGE"]
        )



    # Totals

    ocean_total = ocean_rate * quantity

    transport_total = transport_rate * quantity


    total = (
        ocean_total +
        transport_total
    )



    # Results

    st.success(
        "Quote Generated"
    )


    st.subheader(
        "Cost Breakdown"
    )


    st.write(
        f"Ocean Freight: R {ocean_total:,.2f}"
    )


    st.write(
        f"Transport ({zone}): R {transport_total:,.2f}"
    )


    st.divider()


    st.header(
        f"TOTAL LOGISTICS COST: R {total:,.2f}"
    )
