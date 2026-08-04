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

st.divider()


# =========================
# FILES
# =========================

freight_file = "freight_rates.xls"
transport_file = "transport_rates.xls"



# =========================
# LOAD EXCEL FILES
# =========================

@st.cache_data
def load_data():

    freight = pd.read_excel(
        freight_file
    )

    transport = pd.read_excel(
        transport_file
    )

    return freight, transport



# Check files

if not os.path.exists(freight_file):

    st.error("freight_rates.xls is missing from GitHub")

    st.stop()



if not os.path.exists(transport_file):

    st.error("transport_rates.xls is missing from GitHub")

    st.stop()



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
# FREIGHT DATABASE
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


    # Find freight rate

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



    # Find transport rate

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



    # Total

    ocean_total = ocean_rate * containers

    transport_total = transport_rate * containers

    grand_total = (
        ocean_total +
        transport_total
    )



    # =========================
    # RESULTS
    # =========================


    st.success(
        "Quote Generated"
    )


    st.subheader(
        "Cost Breakdown"
    )


    st.write(
        f"Ocean Freight ({containers} containers): R {ocean_total:,.2f}"
    )


    st.write(
        f"Transport ({zone}): R {transport_total:,.2f}"
    )


    st.divider()


    st.header(
        f"TOTAL LOGISTICS COST: R {grand_total:,.2f}"
    )
