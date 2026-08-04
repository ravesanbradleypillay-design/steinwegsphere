import streamlit as st
import pandas as pd
import os
import re


# =========================================================
# PAGE SETTINGS
# =========================================================

st.set_page_config(
    page_title="SteinwegSphere",
    page_icon="🚢",
    layout="wide",
    initial_sidebar_state="collapsed"
)


# =========================================================
# CLEAN BLUE DESIGN
# =========================================================

st.markdown(
    """
    <style>

    .stApp {
        background:
            linear-gradient(
                135deg,
                #020b18 0%,
                #062344 48%,
                #063b6b 100%
            );
        color: white;
    }

    [data-testid="stHeader"] {
        background: transparent;
    }

    [data-testid="stToolbar"] {
        right: 1rem;
    }

    .block-container {
        max-width: 1400px;
        padding-top: 2rem;
        padding-bottom: 4rem;
    }

    h1, h2, h3 {
        color: white !important;
    }

    .stTextInput label,
    .stSelectbox label,
    .stNumberInput label,
    .stRadio label {
        color: #ccecff !important;
        font-weight: 600 !important;
    }

    .stTextInput input,
    .stNumberInput input {
        background-color: white !important;
        color: #10243b !important;
        border-radius: 10px !important;
        border: 1px solid #4bbdf5 !important;
    }

    div[data-baseweb="select"] > div {
        background-color: white !important;
        color: #10243b !important;
        border-radius: 10px !important;
    }

    div[data-baseweb="select"] * {
        color: #10243b !important;
    }

    .stButton button {
        width: 100%;
        min-height: 44px;
        border-radius: 10px;
        border: 1px solid #35c3ff;
        background:
            linear-gradient(
                90deg,
                #007dc4,
                #00aeea
            );
        color: white;
        font-weight: 700;
    }

    .stButton button:hover {
        border-color: white;
        background:
            linear-gradient(
                90deg,
                #009de5,
                #24cfff
            );
        color: white;
    }

    div[data-testid="stMetric"] {
        background:
            rgba(3, 25, 52, 0.88);
        border:
            1px solid rgba(72, 191, 245, 0.45);
        border-radius: 16px;
        padding: 18px;
    }

    div[data-testid="stMetricLabel"] {
        color: #a9ddf5 !important;
    }

    div[data-testid="stMetricValue"] {
        color: white !important;
    }

    .stDataFrame {
        border-radius: 12px;
        overflow: hidden;
    }

    hr {
        border-color:
            rgba(83, 195, 250, 0.25);
    }

    </style>
    """,
    unsafe_allow_html=True
)


# =========================================================
# TITLE
# =========================================================

st.markdown(
    "<h1 style='text-align:center; margin-bottom:0;'>"
    "C. STEINWEG"
    "</h1>",
    unsafe_allow_html=True
)

st.markdown(
    "<h3 style='text-align:center; color:#55cfff !important; "
    "letter-spacing:6px; margin-top:0;'>"
    "BRIDGE"
    "</h3>",
    unsafe_allow_html=True
)

st.markdown(
    "<p style='text-align:center; color:#b9e8ff; "
    "font-size:17px; margin-bottom:35px;'>"
    "Master Freight & Landside Transport Calculator"
    "</p>",
    unsafe_allow_html=True
)


# =========================================================
# FILE NAMES
# =========================================================

FREIGHT_FILE = "freight rates.xlsx"
TRANSPORT_FILE = "transport rates.xls"


# =========================================================
# CHECK FILES
# =========================================================

if not os.path.exists(FREIGHT_FILE):
    st.error(
        "Freight file not found: "
        "freight rates.xlsx"
    )
    st.stop()


if not os.path.exists(TRANSPORT_FILE):
    st.error(
        "Transport file not found: "
        "transport rates.xls"
    )
    st.stop()


# =========================================================
# LOAD DATA
# =========================================================

@st.cache_data
def load_data():

    freight_data = pd.read_excel(
        FREIGHT_FILE,
        engine="openpyxl"
    )

    transport_data = pd.read_excel(
        TRANSPORT_FILE,
        engine="xlrd",
        header=None
    )

    return freight_data, transport_data


freight, transport = load_data()


# =========================================================
# CLEAN FREIGHT HEADINGS
# =========================================================

freight.columns = (
    freight.columns
    .astype(str)
    .str.strip()
    .str.upper()
)


required_columns = [
    "LINE",
    "COUNTRY",
    "POL",
    "POD",
    "EQUIP TYPE",
    "ALL IN RATE",
    "ROUTING",
    "FREQUENCY",
    "PREPAID / COLLECT"
]


missing_columns = [
    column
    for column in required_columns
    if column not in freight.columns
]


if missing_columns:

    st.error(
        "The following freight columns "
        "were not found: "
        + ", ".join(missing_columns)
    )

    st.write(
        "Columns detected:"
    )

    st.write(
        freight.columns.tolist()
    )

    st.stop()


for column in freight.columns:

    freight[column] = (
        freight[column]
        .fillna("")
        .astype(str)
        .str.strip()
    )


# =========================================================
# HELPER FUNCTIONS
# =========================================================

def clean_text(value):

    if pd.isna(value):
        return ""

    return str(value).strip()


def money_to_number(value):

    if pd.isna(value):
        return None

    if isinstance(
        value,
        (int, float)
    ):
        return float(value)

    text = str(value).strip()

    text = text.replace(
        "USD",
        ""
    )

    text = text.replace(
        "US$",
        ""
    )

    text = text.replace(
        "$",
        ""
    )

    text = text.replace(
        "R",
        ""
    )

    text = text.replace(
        " ",
        ""
    )

    if (
        "," in text
        and "." in text
    ):
        text = text.replace(
            ",",
            ""
        )

    elif "," in text:
        text = text.replace(
            ",",
            "."
        )

    try:
        return float(text)

    except ValueError:
        return None


def get_weight_range(value):

    text = clean_text(value)

    numbers = re.findall(
        r"\d+(?:\.\d+)?",
        text
    )

    if len(numbers) >= 2:

        return (
            float(numbers[0]),
            float(numbers[1])
        )

    return None


def get_logo_url(line_name):

    line = (
        str(line_name)
        .upper()
        .strip()
    )

    logos = {

        "MSC":
        "https://logo.clearbit.com/msc.com",

        "MAERSK":
        "https://logo.clearbit.com/maersk.com",

        "CMA CGM":
        "https://logo.clearbit.com/cmacgm-group.com",

        "CMA-CGM":
        "https://logo.clearbit.com/cmacgm-group.com",

        "HAPAG-LLOYD":
        "https://logo.clearbit.com/hapag-lloyd.com",

        "HAPAG LLOYD":
        "https://logo.clearbit.com/hapag-lloyd.com",

        "COSCO":
        "https://logo.clearbit.com/coscoshipping.com",

        "EVERGREEN":
        "https://logo.clearbit.com/evergreen-marine.com",

        "PIL":
        "https://logo.clearbit.com/pilship.com",

        "ZIM":
        "https://logo.clearbit.com/zim.com",

        "WAN HAI":
        "https://logo.clearbit.com/wanhai.com",

        "YANG MING":
        "https://logo.clearbit.com/yangming.com"

    }

    for key, url in logos.items():

        if key in line:
            return url

    return None


def show_carrier_card(
    column,
    row,
    row_number
):

    line = clean_text(
        row["LINE"]
    )

    rate = money_to_number(
        row["ALL IN RATE"]
    )

    if rate is None:
        rate = 0.0

    routing = clean_text(
        row["ROUTING"]
    )

    frequency = clean_text(
        row["FREQUENCY"]
    )

    equipment = clean_text(
        row["EQUIP TYPE"]
    )

    logo = get_logo_url(
        line
    )

    with column:

        st.markdown(
            """
            <div style="
                background:
                    linear-gradient(
                        145deg,
                        rgba(4,28,58,0.97),
                        rgba(4,70,120,0.90)
                    );
                border:
                    1px solid
                    rgba(75,195,250,0.45);
                border-radius:18px;
                padding:18px;
                min-height:235px;
                box-shadow:
                    0 10px 28px
                    rgba(0,0,0,0.28);
            ">
            """,
            unsafe_allow_html=True
        )

        if logo:

            try:

                st.image(
                    logo,
                    width=105
                )

            except Exception:

                st.markdown(
                    f"### {line}"
                )

        else:

            st.markdown(
                f"### {line}"
            )

        st.markdown(
            f"**{line}**"
        )

        st.markdown(
            f"### USD {rate:,.2f}"
        )

        st.caption(
            f"Equipment: {equipment}"
        )

        st.caption(
            f"Routing: {routing}"
        )

        st.caption(
            f"Frequency: {frequency}"
        )

        if st.button(
            "Select this option",
            key=f"select_{row_number}"
        ):

            st.session_state[
                "selected_freight_row"
            ] = row_number

        st.markdown(
            "</div>",
            unsafe_allow_html=True
        )


# =========================================================
# FREIGHT SEARCH
# =========================================================

st.header(
    "Ocean Freight Search"
)


search_col1, search_col2 = (
    st.columns(2)
)


with search_col1:

    user_pol = st.text_input(
        "Port of Loading (POL)",
        placeholder="Example: Durban"
    )


with search_col2:

    user_pod = st.text_input(
        "Port of Discharge (POD)",
        placeholder="Example: Mundra"
    )


st.caption(
    "Enter all or part of the port name."
)


# =========================================================
# SEARCH ROUTE
# =========================================================

if (
    user_pol.strip()
    and user_pod.strip()
):

    route_matches = freight[

        freight["POL"]
        .str.contains(
            user_pol.strip(),
            case=False,
            na=False
        )

        &

        freight["POD"]
        .str.contains(
            user_pod.strip(),
            case=False,
            na=False
        )

    ].copy()


    if route_matches.empty:

        st.warning(
            "No freight options were found "
            "for this POL and POD."
        )

        st.stop()


    # =====================================================
    # EQUIPMENT FILTER
    # =====================================================

    equipment_list = (
        route_matches[
            "EQUIP TYPE"
        ]
        .replace(
            "",
            pd.NA
        )
        .dropna()
        .unique()
        .tolist()
    )

    equipment_list = sorted(
        equipment_list
    )

    equipment_filter = st.selectbox(
        "Filter by Equipment Type",
        ["All"] + equipment_list
    )


    filtered = (
        route_matches.copy()
    )


    if equipment_filter != "All":

        filtered = filtered[

            filtered[
                "EQUIP TYPE"
            ]

            ==

            equipment_filter

        ]


    # =====================================================
    # REMOVE EXACT DUPLICATES
    # =====================================================

    filtered = (

        filtered

        .drop_duplicates(

            subset=[
                "LINE",
                "EQUIP TYPE",
                "ALL IN RATE",
                "ROUTING",
                "FREQUENCY",
                "PREPAID / COLLECT"
            ]

        )

        .reset_index(
            drop=True
        )

    )


    # Reset selection if a new route
    # produces fewer results.

    if (
        "selected_freight_row"
        in st.session_state
    ):

        if (
            st.session_state[
                "selected_freight_row"
            ]
            >= len(filtered)
        ):

            del st.session_state[
                "selected_freight_row"
            ]


    st.success(
        f"{len(filtered)} "
        "freight option(s) found."
    )


    st.subheader(
        "Available Shipping Lines"
    )


    # =====================================================
    # SHOW THREE CARDS PER ROW
    # =====================================================

    for start in range(
        0,
        len(filtered),
        3
    ):

        columns = st.columns(3)

        for position in range(3):

            row_number = (
                start + position
            )

            if (
                row_number
                >= len(filtered)
            ):
                continue

            show_carrier_card(
                columns[position],
                filtered.iloc[
                    row_number
                ],
                row_number
            )


    # =====================================================
    # SELECTED FREIGHT
    # =====================================================

    if (
        "selected_freight_row"
        not in st.session_state
    ):

        st.info(
            "Select a shipping-line "
            "option above to continue."
        )

        st.stop()


    selected_index = (
        st.session_state[
            "selected_freight_row"
        ]
    )


    selected_row = (
        filtered.iloc[
            selected_index
        ]
    )


    selected_line = clean_text(
        selected_row["LINE"]
    )

    selected_equipment = clean_text(
        selected_row[
            "EQUIP TYPE"
        ]
    )

    selected_routing = clean_text(
        selected_row[
            "ROUTING"
        ]
    )

    selected_frequency = clean_text(
        selected_row[
            "FREQUENCY"
        ]
    )

    selected_rate = money_to_number(
        selected_row[
            "ALL IN RATE"
        ]
    )

    if selected_rate is None:
        selected_rate = 0.0


    st.divider()


    st.subheader(
        "Selected Ocean Freight"
    )


    selected_col1, selected_col2 = (
        st.columns(
            [2, 1]
        )
    )


    with selected_col1:

        st.success(
            f"Selected: {selected_line}"
        )

        st.write(
            f"**Equipment:** "
            f"{selected_equipment}"
        )

        st.write(
            f"**Routing:** "
            f"{selected_routing}"
        )

        st.write(
            f"**Frequency:** "
            f"{selected_frequency}"
        )


    with selected_col2:

        st.metric(
            "Freight Rate",
            f"USD {selected_rate:,.2f}"
        )


    quantity = st.number_input(
        "Number of Containers",
        min_value=1,
        value=1,
        step=1
    )


    freight_total = (
        selected_rate
        * quantity
    )


    # =====================================================
    # TRANSPORT
    # =====================================================

    st.divider()

    st.header(
        "Landside Transport"
    )


    transport_col1, transport_col2, transport_col3 = (
        st.columns(3)
    )


    with transport_col1:

        selected_zone = st.selectbox(
            "Transport Zone",
            [
                "Zone A",
                "Zone B",
                "Zone C"
            ]
        )


    with transport_col2:

        transport_type = st.selectbox(
            "Container Type",
            [
                "6M FCL (20ft)",
                "12M FCL (40ft)"
            ]
        )


    with transport_col3:

        cargo_weight = (
            st.number_input(
                "Cargo Weight (Tons)",
                min_value=0.0,
                value=20.0,
                step=0.1
            )
        )


    # =====================================================
    # TRANSPORT COLUMN MAP
    # =====================================================

    zone_columns = {

        "Zone A": {
            "weight": 0,
            "total": 4
        },

        "Zone B": {
            "weight": 5,
            "total": 9
        },

        "Zone C": {
            "weight": 10,
            "total": 14
        }

    }


    weight_column = (
        zone_columns[
            selected_zone
        ]["weight"]
    )


    total_column = (
        zone_columns[
            selected_zone
        ]["total"]
    )


    if (
        transport_type
        == "6M FCL (20ft)"
    ):

        section_name = "6M FCL"

    else:

        section_name = "12M FCL"


    # =====================================================
    # FIND TRANSPORT SECTION
    # =====================================================

    section_start = None


    for row_number in range(
        len(transport)
    ):

        row_values = (

            transport.iloc[
                row_number
            ]

            .fillna("")

            .astype(str)

            .tolist()

        )


        row_text = (
            " ".join(
                row_values
            )
            .upper()
        )


        if (
            section_name
            in row_text
        ):

            section_start = (
                row_number + 1
            )

            break


    transport_rate = None

    applied_weight_break = (
        "Not found"
    )


    if (
        section_start
        is not None
    ):

        for row_number in range(

            section_start,

            min(
                section_start + 12,
                len(transport)
            )

        ):

            weight_text = clean_text(

                transport.iloc[
                    row_number,
                    weight_column
                ]

            )


            total_value = (

                transport.iloc[
                    row_number,
                    total_column
                ]

            )


            if (
                not weight_text
            ):
                continue


            if (
                "FCL"
                in weight_text.upper()
            ):
                break


            weight_range = (
                get_weight_range(
                    weight_text
                )
            )


            if (
                weight_range
                is None
            ):
                continue


            minimum = (
                weight_range[0]
            )

            maximum = (
                weight_range[1]
            )


            if (
                minimum
                <= cargo_weight
                <= maximum
            ):

                transport_rate = (
                    money_to_number(
                        total_value
                    )
                )

                applied_weight_break = (
                    weight_text
                )

                break


    if (
        transport_rate
        is None
    ):

        transport_total = 0.0

    else:

        transport_total = (
            transport_rate
            * quantity
        )


    # =====================================================
    # FINAL SUMMARY
    # =====================================================

    st.divider()

    st.header(
        "Quote Summary"
    )


    summary_col1, summary_col2 = (
        st.columns(2)
    )


    with summary_col1:

        st.metric(
            "Ocean Freight Total",
            f"USD {freight_total:,.2f}"
        )

        st.write(
            f"**Shipping Line:** "
            f"{selected_line}"
        )

        st.write(
            f"**POL:** "
            f"{user_pol}"
        )

        st.write(
            f"**POD:** "
            f"{user_pod}"
        )

        st.write(
            f"**Equipment:** "
            f"{selected_equipment}"
        )

        st.write(
            f"**Routing:** "
            f"{selected_routing}"
        )

        st.write(
            f"**Frequency:** "
            f"{selected_frequency}"
        )

        st.write(
            f"**Containers:** "
            f"{quantity}"
        )


    with summary_col2:

        if (
            transport_rate
            is None
        ):

            st.warning(
                "No transport rate "
                "was found for this "
                "weight break."
            )

        else:

            st.metric(
                "Transport Total",
                f"R {transport_total:,.2f}"
            )


        st.write(
            f"**Zone:** "
            f"{selected_zone}"
        )

        st.write(
            f"**Transport Type:** "
            f"{transport_type}"
        )

        st.write(
            f"**Cargo Weight:** "
            f"{cargo_weight:,.1f} Tons"
        )

        st.write(
            f"**Weight Break:** "
            f"{applied_weight_break}"
        )

        st.write(
            f"**Containers:** "
            f"{quantity}"
        )


    st.info(
        "Ocean freight is displayed "
        "in USD. Landside transport "
        "is displayed in ZAR. "
        "The currencies are kept "
        "separate."
    )


else:

    st.info(
        "Enter both the POL and POD "
        "to display all available "
        "shipping-line options."
    )
