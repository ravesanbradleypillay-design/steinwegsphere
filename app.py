import streamlit as st
import pandas as pd
import os


# =========================
# PAGE CONFIG
# =========================

st.set_page_config(
    page_title="Steinweg Freight Calculator",
    page_icon="🚢",
    layout="wide"
)


st.title("STEINWEG")
st.subheader("Master Freight & Transport Calculator")

st.divider()


# =========================
# EXCEL FILES
# =========================

freight_file = "freight_rates.xlsx"
transport_file = "transport_rates.xlsx"



# =========================
# LOAD DATA
# =========================

@st.cache_data
def load_excel():

    freight = pd.read_excel(
        freight_file,
        engine="openpyxl"
    )

    transport = pd.read_excel(
        transport_file,
        engine="openpyxl"
    )

    return freight, transport



# Check files exist

if not os.path.exists(freight_file):

    st.error("freight_rates.xlsx is missing from GitHub")

    st.stop()


if not os.path.exists(transport_file):

    st.error("transport_rates.xlsx is missing from GitHub")

    st.stop()



freight, transport = load_excel()



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
# SHOW DATABASE STATUS
# =========================

with st.expander("Database Preview"):

    st.write("Freight Rates")

    st.dataframe(
        freight.head()
    )


    st.write("Transport Rates")

    st.dataframe(
        transport.head()
    )



# =========================
# SHIPMENT DETAILS
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



containers = st.number_input(
    "Number of Containers",
    min_value=1,
    value=1
)



st.divider()


# =========================
# TRANSPORT DETAILS
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


    # ---------------------
    # FREIGHT SEARCH
    # ---------------------

    freight_result = freight[

        (freight["POD"] == pod)

        &

        (freight["EQUIP TYPE"] == equipment)

        &

        (freight["PREPAID / COLLECT"] == payment)

    ]


    if freight_result.empty:

        ocean_rate = 0

        st.warning(
            "No freight rate found"
        )

    else:

        ocean_rate = float(
            freight_result.iloc[0]["ALL IN RATE"]
        )



    # ---------------------
    # TRANSPORT SEARCH
    # ---------------------

    transport_result = transport[

        transport["ZONE"] == zone

    ]


    if transport_result.empty:

        transport_rate = 0

        st.warning(
            "No transport rate found"
        )

    else:

        transport_rate = float(
            transport_result.iloc[0]["TOTAL CHARGE"]
        )



    # ---------------------
    # TOTALS
    # ---------------------

    freight_total = ocean_rate * containers

    transport_total = transport_rate * containers


    total = (
        freight_total +
        transport_total
    )



    # ---------------------
    # OUTPUT
    # ---------------------

    st.success(
        "Quote Generated Successfully"
    )


    st.subheader(
        "Cost Breakdown"
    )


    st.write(
        f"Ocean Freight: R {freight_total:,.2f}"
    )


    st.write(
        f"Transport ({zone}): R {transport_total:,.2f}"
    )


    st.divider()


    st.header(
        f"TOTAL LOGISTICS COST: R {total:,.2f}"
    )
