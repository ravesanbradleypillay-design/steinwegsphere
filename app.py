import streamlit as st
import pandas as pd
import os
import re


# =========================================================
# PAGE SETTINGS
# =========================================================

st.set_page_config(
    page_title="C. Steinweg Bridge",
    page_icon="🚢",
    layout="wide",
    initial_sidebar_state="collapsed"
)


# =========================================================
# BLUE APP DESIGN
# =========================================================

st.markdown(
    """
    <style>

    .stApp {
        background: linear-gradient(
            135deg,
            #020b18 0%,
            #062344 50%,
            #063b6b 100%
        );
        color: white;
    }

    [data-testid="stHeader"] {
        background: transparent;
    }

    .block-container {
        max-width: 1450px;
        padding-top: 2rem;
        padding-bottom: 4rem;
    }

    h1, h2, h3 {
        color: white !important;
    }

    .stTextInput label,
    .stSelectbox label,
    .stNumberInput label {
        color: #d4efff !important;
        font-weight: 650 !important;
    }

    .stTextInput input,
    .stNumberInput input {
        background-color: white !important;
        color: #10243b !important;
        border: 1px solid #4bbdf5 !important;
        border-radius: 10px !important;
    }

    div[data-baseweb="select"] > div {
        background-color: white !important;
        color: #10243b !important;
        border: 1px solid #4bbdf5 !important;
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
        background: linear-gradient(
            90deg,
            #007dc4,
            #00aeea
        );
        color: white;
        font-weight: 700;
    }

    .stButton button:hover {
        border-color: white;
        background: linear-gradient(
            90deg,
            #009de5,
            #24cfff
        );
        color: white;
    }

    div[data-testid="stMetric"] {
        background: rgba(
            3,
            25,
            52,
            0.88
        );

        border: 1px solid rgba(
            72,
            191,
            245,
            0.45
        );

        border-radius: 16px;
        padding: 18px;
    }

    div[data-testid="stMetricLabel"] {
        color: #a9ddf5 !important;
    }

    div[data-testid="stMetricValue"] {
        color: white !important;
    }

    div[data-testid="stVerticalBlockBorderWrapper"] {
        background: rgba(
            3,
            25,
            52,
            0.82
        );

        border: 1px solid rgba(
            75,
            195,
            250,
            0.40
        );

        border-radius: 18px;
    }

    hr {
        border-color: rgba(
            83,
            195,
            250,
            0.25
        );
    }

    </style>
    """,
    unsafe_allow_html=True
)


# =========================================================
# APP HEADER
# =========================================================

st.markdown(
    """
    <h1 style="
        text-align: center;
        margin-bottom: 0;
        font-size: 48px;
    ">
        C. STEINWEG
    </h1>
    """,
    unsafe_allow_html=True
)

st.markdown(
    """
    <h3 style="
        text-align: center;
        color: #55cfff !important;
        letter-spacing: 7px;
        margin-top: 0;
    ">
        BRIDGE
    </h3>
    """,
    unsafe_allow_html=True
)

st.markdown(
    """
    <p style="
        text-align: center;
        color: #b9e8ff;
        font-size: 17px;
        margin-bottom: 35px;
    ">
        Master Ocean Freight & Landside Transport Calculator
    </p>
    """,
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
# LOAD EXCEL FILES
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
# CLEAN FREIGHT DATA
# =========================================================

freight.columns = (
    freight.columns
    .astype(str)
    .str.strip()
    .str.upper()
)


required_columns = [
    "LINE",
    "POL",
    "POD",
    "EQUIP TYPE",
    "ALL IN RATE",
    "ROUTING",
    "FREQUENCY"
]


missing_columns = [

    column

    for column in required_columns

    if column not in freight.columns

]


if missing_columns:

    st.error(
        "Missing freight columns: "
        + ", ".join(missing_columns)
    )

    st.write(
        "Columns found:"
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


def equipment_matches(
    spreadsheet_equipment,
    selected_size
):

    equipment = (
        clean_text(
            spreadsheet_equipment
        )
        .upper()
    )


    if selected_size == "20ft":

        keywords = [
            "20",
            "6M",
            "6 M"
        ]


    else:

        keywords = [
            "40",
            "12M",
            "12 M"
        ]


    return any(

        keyword in equipment

        for keyword in keywords

    )


def carrier_initials(line):

    words = (

        clean_text(line)

        .replace(
            "-",
            " "
        )

        .split()

    )


    if not words:

        return "SL"


    if len(words) == 1:

        return words[0][:3].upper()


    return "".join(

        word[0]

        for word in words[:3]

    ).upper()


# =========================================================
# FREIGHT SEARCH
# =========================================================

st.header(
    "Ocean Freight Search"
)


search_one, search_two, search_three = (
    st.columns(
        [2, 2, 1]
    )
)


with search_one:

    user_pol = st.text_input(
        "Port of Loading (POL)",
        placeholder="Example: Durban"
    )


with search_two:

    user_pod = st.text_input(
        "Port of Discharge (POD)",
        placeholder="Example: Mundra"
    )


with search_three:

    container_size = st.selectbox(
        "Container Size",
        [
            "20ft",
            "40ft"
        ]
    )


st.caption(
    "Enter all or part of the port name. "
    "The container dropdown filters "
    "the available freight rates."
)


# =========================================================
# SEARCH RESULTS
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


    # =====================================================
    # FILTER BY 20FT OR 40FT
    # =====================================================

    route_matches = route_matches[

        route_matches[
            "EQUIP TYPE"
        ].apply(

            lambda value:

            equipment_matches(
                value,
                container_size
            )

        )

    ].copy()


    if route_matches.empty:

        st.warning(
            f"No {container_size} freight "
            "options were found for "
            f"{user_pol} to {user_pod}."
        )

        st.stop()


    # =====================================================
    # REMOVE DUPLICATES
    # =====================================================

    filtered = (

        route_matches

        .drop_duplicates(

            subset=[
                "LINE",
                "EQUIP TYPE",
                "ALL IN RATE",
                "ROUTING",
                "FREQUENCY"
            ]

        )

        .reset_index(
            drop=True
        )

    )


    # =====================================================
    # RESET SELECTION FOR NEW SEARCH
    # =====================================================

    current_search = (

        user_pol.strip().upper(),

        user_pod.strip().upper(),

        container_size

    )


    if (

        st.session_state.get(
            "last_search"
        )

        !=

        current_search

    ):


        st.session_state[
            "last_search"
        ] = current_search


        st.session_state.pop(
            "selected_freight_row",
            None
        )


    st.success(
        f"{len(filtered)} "
        f"{container_size} shipping "
        "option(s) found."
    )


    st.subheader(
        "Available Shipping Lines"
    )


    # =====================================================
    # SHIPPING LINE CARDS
    # =====================================================

    for start in range(
        0,
        len(filtered),
        3
    ):


        card_columns = (
            st.columns(3)
        )


        for position in range(3):


            row_number = (
                start
                + position
            )


            if (
                row_number
                >= len(filtered)
            ):

                continue


            row = (
                filtered.iloc[
                    row_number
                ]
            )


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


            initials = (
                carrier_initials(
                    line
                )
            )


            with card_columns[position]:


                with st.container(
                    border=True
                ):


                    st.markdown(
                        f"## 🚢 {line}"
                    )


                    st.metric(
                        "Ocean Freight Rate",
                        f"USD {rate:,.2f}"
                    )


                    st.write(
                        f"**Routing:** "
                        f"{routing}"
                    )


                    st.write(
                        f"**Frequency:** "
                        f"{frequency}"
                    )


                    st.caption(
                        f"Carrier code: "
                        f"{initials}"
                    )


                    if st.button(

                        "Select Shipping Line",

                        key=(
                            f"select_"
                            f"{row_number}"
                        )

                    ):


                        st.session_state[

                            "selected_freight_row"

                        ] = row_number


                        st.rerun()


    # =====================================================
    # SELECTED FREIGHT
    # =====================================================

    if (

        "selected_freight_row"

        not in

        st.session_state

    ):


        st.info(
            "Select a shipping line "
            "above to continue."
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


    st.header(
        "Selected Ocean Freight"
    )


    selected_one, selected_two = (
        st.columns(
            [2, 1]
        )
    )


    with selected_one:


        st.success(
            f"Selected: "
            f"{selected_line}"
        )


        st.write(
            f"**Container:** "
            f"{container_size}"
        )


        st.write(
            f"**Equipment in Rate Sheet:** "
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


    with selected_two:


        st.metric(
            "Ocean Freight Rate",
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
    # LANDSIDE TRANSPORT
    # =====================================================

    st.divider()


    st.header(
        "Landside Transport"
    )


    transport_one, transport_two, transport_three = (
        st.columns(3)
    )


    with transport_one:


        selected_zone = st.selectbox(
            "Transport Zone",
            [
                "Zone A",
                "Zone B",
                "Zone C"
            ]
        )


    with transport_two:


        if container_size == "20ft":

            transport_type = (
                "6M FCL"
            )

        else:

            transport_type = (
                "12M FCL"
            )


        st.text_input(
            "Transport Equipment",
            value=transport_type,
            disabled=True
        )


    with transport_three:


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


    # =====================================================
    # FIND TRANSPORT SECTION
    # =====================================================

    section_start = None


    for row_number in range(
        len(transport)
    ):


        row_text = (

            " ".join(

                transport.iloc[
                    row_number
                ]

                .fillna("")

                .astype(str)

                .tolist()

            )

            .upper()

        )


        if (

            transport_type

            in

            row_text

        ):


            section_start = (
                row_number
                + 1
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


            weight_text = (

                clean_text(

                    transport.iloc[

                        row_number,

                        weight_column

                    ]

                )

            )


            total_value = (

                transport.iloc[

                    row_number,

                    total_column

                ]

            )


            if not weight_text:

                continue


            if (

                "FCL"

                in

                weight_text.upper()

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

                <=

                cargo_weight

                <=

                maximum

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
    # QUOTE SUMMARY
    # =====================================================

    st.divider()


    st.header(
        "Quote Summary"
    )


    summary_one, summary_two = (
        st.columns(2)
    )


    with summary_one:


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
            f"**Container Size:** "
            f"{container_size}"
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


    with summary_two:


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

        "Ocean freight is shown "

        "in USD and landside "

        "transport is shown "

        "in ZAR. The currencies "

        "are kept separate."

    )


else:


    st.info(

        "Enter the POL and POD, "

        "then select 20ft or 40ft "

        "to view matching "

        "shipping-line rates."

    )
