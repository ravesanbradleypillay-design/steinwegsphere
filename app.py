import streamlit as st
import pandas as pd
import os


st.set_page_config(
    page_title="Steinweg Landside Calculator",
    page_icon="🚢",
    layout="wide"
)


st.title("STEINWEG")
st.subheader("Master Freight & Transport Calculator")


# FILES

freight_file = "Copy of CSI Africa Buy Sell Rates Manual - (01 Aug 2026) V.1.xls"
transport_file = "CSL - July 2026.xls"



# LOAD EXCEL

@st.cache_data
def load_excel():

    freight = pd.read_excel(
        freight_file
    )

    transport = pd.read_excel(
        transport_file
    )

    return freight, transport



if not os.path.exists(freight_file):
    st.error("Freight Excel file missing")
    st.stop()


if not os.path.exists(transport_file):
    st.error("Transport Excel file missing")
    st.stop()



freight, transport = load_excel()



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



# FIND IMPORTANT COLUMNS AUTOMATICALLY

def find_column(df, words):

    for col in df.columns:

        for word in words:

            if word.lower() in col.lower():
                return col

    return None



pod_col = find_column(
    freight,
    ["POD"]
)


equip_col = find_column(
    freight,
    ["EQUIP"]
)


rate_col = find_column(
    freight,
    ["ALL IN", "RATE"]
)


payment_col = find_column(
    freight,
    ["PREPAID"]
)



zone_col = find_column(
    transport,
    ["ZONE"]
)


charge_col = find_column(
    transport,
    ["TOTAL"]
)



st.divider()


st.header("Shipment Details")


col1,col2,col3 = st.columns(3)


with col1:

    pod = st.selectbox(
        "POD",
        sorted(
            freight[pod_col]
            .dropna()
            .unique()
        )
    )


with col2:

    equipment = st.selectbox(
        "Equipment Type",
        sorted(
            freight[equip_col]
            .dropna()
            .unique()
        )
    )


with col3:

    payment = st.selectbox(
        "Payment Type",
        sorted(
            freight[payment_col]
            .dropna()
            .unique()
        )
    )



st.divider()


zone = st.selectbox(
    "Transport Zone",
    sorted(
        transport[zone_col]
        .dropna()
        .unique()
    )
)



qty = st.number_input(
    "Number of Containers",
    min_value=1,
    value=1
)



if st.button("Calculate"):


    # FREIGHT LOOKUP

    freight_result = freight[

        (freight[pod_col] == pod)

        &

        (freight[equip_col] == equipment)

        &

        (freight[payment_col] == payment)

    ]



    if len(freight_result) > 0:

        ocean = float(
            freight_result.iloc[0][rate_col]
        )

    else:

        ocean = 0



    # TRANSPORT LOOKUP

    transport_result = transport[
        transport[zone_col] == zone
    ]


    if len(transport_result) > 0:

        inland = float(
            transport_result.iloc[0][charge_col]
        )

    else:

        inland = 0



    total = (
        ocean +
        inland
    ) * qty



    st.success(
        "Quote Generated"
    )


    st.write(
        f"Ocean Freight: R {ocean:,.2f}"
    )


    st.write(
        f"Transport ({zone}): R {inland:,.2f}"
    )


    st.divider()


    st.header(
        f"TOTAL: R {total:,.2f}"
    )
