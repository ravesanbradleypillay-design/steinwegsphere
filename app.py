import streamlit as st
import pandas as pd
import os
import re


# ==================================================
# PAGE CONFIGURATION
# ==================================================

st.set_page_config(
    page_title="SteinwegSphere",
    page_icon="🚢",
    layout="wide",
    initial_sidebar_state="collapsed"
)


# ==================================================
# BLUE LOGISTICS DESIGN
# ==================================================

st.markdown(
    """
    <style>

    .stApp {
        background:
        linear-gradient(
            135deg,
            rgba(3, 15, 35, 0.97),
            rgba(5, 36, 75, 0.94),
            rgba(0, 78, 145, 0.88)
        ),
        url("https://images.unsplash.com/photo-1494412574643-ff11b0a5c1c3?auto=format&fit=crop&w=2200&q=85");

        background-size: cover;
        background-position: center;
        background-attachment: fixed;
    }

    [data-testid="stHeader"] {
        background: rgba(0, 0, 0, 0);
    }

    .main {
        background: transparent;
    }

    .hero {
        text-align: center;
        padding: 35px 20px 28px 20px;
        border-radius: 25px;
        background:
        linear-gradient(
            135deg,
            rgba(4, 19, 42, 0.96),
            rgba(0, 91, 170, 0.78)
        );

        border: 1px solid
        rgba(87, 190, 255, 0.45);

        box-shadow:
        0 15px 50px
        rgba(0, 0, 0, 0.40);

        margin-bottom: 25px;
    }

    .steinweg-title {
        color: white;
        font-size: 48px;
        font-weight: 800;
        letter-spacing: 7px;
        margin-bottom: 0px;
    }

    .bridge-title {
        color: #54c7ff;
        font-size: 21px;
        font-weight: 500;
        letter-spacing: 5px;
        margin-top: 2px;
    }

    .hero-subtitle {
        color: #c9eaff;
        font-size: 16px;
        margin-top: 18px;
    }

    .section-title {
        color: white;
        font-size: 25px;
        font-weight: 700;
        margin-top: 12px;
        margin-bottom: 12px;
    }

    .carrier-card {
        background:
        linear-gradient(
            145deg,
            rgba(7, 28, 58, 0.96),
            rgba(8, 74, 128, 0.80)
        );

        border:
        1px solid
        rgba(80, 190, 255, 0.35);

        border-radius: 18px;

        padding: 18px;

        min-height: 205px;

        box-shadow:
        0 8px 25px
        rgba(0, 0, 0, 0.28);

        margin-bottom: 12px;
    }

    .carrier-name {
        color: white;
        font-size: 22px;
        font-weight: 750;
        margin-top: 8px;
    }

    .rate-text {
        color: #55d0ff;
        font-size: 25px;
        font-weight: 800;
        margin-top: 10px;
    }

    .small-label {
        color: #9bcde7;
        font-size: 12px;
        text-transform: uppercase;
        letter-spacing: 1px;
        margin-top: 10px;
    }

    .small-value {
        color: white;
        font-size: 15px;
        font-weight: 500;
    }

    .summary-card {
        background:
        linear-gradient(
            145deg,
            rgba(4, 24, 51, 0.98),
            rgba(0, 88, 155, 0.80)
        );

        border:
        1px solid
        rgba(78, 201, 255, 0.50);

        border-radius: 20px;

        padding: 25px;

        box-shadow:
        0 12px 35px
        rgba(0, 0, 0, 0.35);
    }

    .summary-heading {
        color: #8bdcff;
        font-size: 14px;
        font-weight: 700;
        letter-spacing: 2px;
    }

    .summary-value {
        color: white;
        font-size: 32px;
        font-weight: 800;
        margin-top: 5px;
    }

    .currency-note {
        color: #b9e7ff;
        font-size: 13px;
        margin-top: 8px;
    }

    .stTextInput input,
    .stNumberInput input,
    .stSelectbox div[data-baseweb="select"] {
        background-color:
        rgba(255, 255, 255, 0.96);

        border-radius: 10px;
    }

    .stButton button {
        width: 100%;
        border-radius: 12px;
        border: none;

        background:
        linear-gradient(
            90deg,
            #0088d1,
            #00b9f2
        );

        color: white;

        font-size: 17px;
        font-weight: 700;

        padding: 12px;

        box-shadow:
        0 7px 20px
        rgba(0, 150, 230, 0.30);
    }

    .stButton button:hover {
        background:
        linear-gradient(
            90deg,
            #00a5ed,
            #32d0ff
        );

        color: white;
    }

    div[data-testid="stMetric"] {
        background:
        rgba(5, 28, 57, 0.82);

        border:
        1px solid
        rgba(80, 190, 255, 0.35);

        padding: 18px;

        border-radius: 16px;
    }

    div[data-testid="stMetricLabel"] {
        color: #a8dcf5;
    }

    div[data-testid="stMetricValue"] {
        color: white;
    }

    </style>
    """,
    unsafe_allow_html=True
)


# ==================================================
# HERO
# ==================================================

st.markdown(
    """
    <div class="hero">

        <div class="steinweg-title">
        C. STEINWEG
        </div>

        <div class="bridge-title">
        BRIDGE
        </div>

        <div class="hero-subtitle">
        Intelligent Freight • Ocean Logistics • Landside Transport
        </div>

    </div>
    """,
    unsafe_allow_html=True
)


# ==================================================
# FILES
# ==================================================

FREIGHT_FILE = "freight rates.xlsx"
TRANSPORT_FILE = "transport rates.xls"


if not os.path.exists(FREIGHT_FILE):

    st.error(
        "Missing file: freight rates.xlsx"
    )

    st.stop()


if not os.path.exists(TRANSPORT_FILE):

    st.error(
        "Missing file: transport rates.xls"
    )

    st.stop()


# ==================================================
# LOAD DATA
# ==================================================

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


# ==================================================
# CLEAN FREIGHT DATA
# ==================================================

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
        "Missing freight columns: "
        +
        ", ".join(missing_columns)
    )

    st.write(
        freight.columns.tolist()
    )

    st.stop()


for column in required_columns:

    freight[column] = (

        freight[column]

        .fillna("")

        .astype(str)

        .str.strip()

    )


# ==================================================
# HELPER FUNCTIONS
# ==================================================

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


    text = str(value)

    text = text.replace(
        "USD",
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
        and
        "." in text
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


def carrier_logo(line):

    line = str(line).upper().strip()


    logo_map = {

        "MSC":
        "https://logo.clearbit.com/msc.com",

        "MEDITERRANEAN SHIPPING COMPANY":
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

        "ONE":
        "https://logo.clearbit.com/one-line.com",

        "OCEAN NETWORK EXPRESS":
        "https://logo.clearbit.com/one-line.com",

        "COSCO":
        "https://logo.clearbit.com/coscoshipping.com",

        "EVERGREEN":
        "https://logo.clearbit.com/evergreen-marine.com",

        "PIL":
        "https://logo.clearbit.com/pilship.com",

        "ZIM":
        "https://logo.clearbit.com/zim.com"

    }


    for key, url in logo_map.items():

        if key in line:

            return url


    return None


# ==================================================
# FREIGHT SEARCH
# ==================================================

st.markdown(
    """
    <div class="section-title">
    Search Ocean Freight
    </div>
    """,
    unsafe_allow_html=True
)


search_one, search_two = (
    st.columns(2)
)


with search_one:

    user_pol = st.text_input(

        "Port of Loading (POL)",

        placeholder=
        "Example: Durban"

    )


with search_two:

    user_pod = st.text_input(

        "Port of Discharge (POD)",

        placeholder=
        "Example: Mundra"

    )


st.caption(
    "Enter part or all of the port name."
)


# ==================================================
# SEARCH RESULTS
# ==================================================

if (
    user_pol.strip()
    and
    user_pod.strip()
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
            "No shipping-line rates were found "
            "for this POL and POD."
        )


    else:


        # ==============================================
        # FILTERS
        # ==============================================

        filter_one, filter_two = (
            st.columns(2)
        )


        with filter_one:

            equipment_options = [

                "All"

            ] + sorted(

                route_matches[
                    "EQUIP TYPE"
                ]

                .replace(
                    "",
                    pd.NA
                )

                .dropna()

                .unique()

            )


            equipment_filter = (

                st.selectbox(

                    "Container Type",

                    equipment_options

                )

            )


        with filter_two:

            payment_options = [

                "All"

            ] + sorted(

                route_matches[
                    "PREPAID / COLLECT"
                ]

                .replace(
                    "",
                    pd.NA
                )

                .dropna()

                .unique()

            )


            payment_filter = (

                st.selectbox(

                    "Payment Type",

                    payment_options

                )

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


        if payment_filter != "All":

            filtered = filtered[

                filtered[
                    "PREPAID / COLLECT"
                ]

                ==

                payment_filter

            ]


        # ==============================================
        # REMOVE DUPLICATE OPTIONS
        # ==============================================

        filtered = (

            filtered

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


        st.markdown(
            """
            <div class="section-title">
            Available Shipping Lines
            </div>
            """,
            unsafe_allow_html=True
        )


        st.success(

            f"{len(filtered)} "

            "shipping option(s) found."

        )


        # ==============================================
        # SHOW CARRIER CARDS
        # ==============================================

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

                    +

                    position

                )


                if row_number >= len(filtered):

                    continue


                row = (

                    filtered.iloc[

                        row_number

                    ]

                )


                line_name = (

                    clean_text(

                        row["LINE"]

                    )

                )


                rate = (

                    money_to_number(

                        row[
                            "ALL IN RATE"
                        ]

                    )

                )


                if rate is None:

                    rate = 0.0


                routing = (

                    clean_text(

                        row["ROUTING"]

                    )

                )


                frequency = (

                    clean_text(

                        row["FREQUENCY"]

                    )

                )


                logo = (

                    carrier_logo(

                        line_name

                    )

                )


                with card_columns[position]:


                    if logo:

                        try:

                            st.image(

                                logo,

                                width=115

                            )

                        except Exception:

                            st.markdown(

                                f"""

                                <div class=
                                "carrier-name">

                                {line_name}

                                </div>

                                """,

                                unsafe_allow_html=True

                            )


                    st.markdown(

                        f"""

                        <div class=
                        "carrier-card">

                        <div class=
                        "carrier-name">

                        {line_name}

                        </div>

                        <div class=
                        "rate-text">

                        USD {rate:,.2f}

                        </div>

                        <div class=
                        "small-label">

                        Routing

                        </div>

                        <div class=
                        "small-value">

                        {routing}

                        </div>

                        <div class=
                        "small-label">

                        Frequency

                        </div>

                        <div class=
                        "small-value">

                        {frequency}

                        </div>

                        </div>

                        """,

                        unsafe_allow_html=True

                    )


        # ==============================================
        # SELECT CARRIER
        # ==============================================

        st.markdown(
            """
            <div class="section-title">
            Select Your Shipping Option
            </div>
            """,
            unsafe_allow_html=True
        )


        option_labels = []


        for index, row in filtered.iterrows():

            rate = (

                money_to_number(

                    row[
                        "ALL IN RATE"
                    ]

                )

            )


            if rate is None:

                rate = 0.0


            option_labels.append(

                f"{row['LINE']} | "

                f"{row['EQUIP TYPE']} | "

                f"USD {rate:,.2f} | "

                f"{row['ROUTING']} | "

                f"{row['FREQUENCY']}"

            )


        selected_option = (

            st.radio(

                "Choose a shipping line",

                options=list(

                    range(

                        len(option_labels)

                    )

                ),

                format_func=

                lambda x:

                option_labels[x],

                horizontal=False

            )

        )


        selected_row = (

            filtered.iloc[

                selected_option

            ]

        )


        selected_rate = (

            money_to_number(

                selected_row[
                    "ALL IN RATE"
                ]

            )

        )


        if selected_rate is None:

            selected_rate = 0.0


        selected_line = (

            clean_text(

                selected_row[
                    "LINE"
                ]

            )

        )


        selected_routing = (

            clean_text(

                selected_row[
                    "ROUTING"
                ]

            )

        )


        selected_frequency = (

            clean_text(

                selected_row[
                    "FREQUENCY"
                ]

            )

        )


        selected_equipment = (

            clean_text(

                selected_row[
                    "EQUIP TYPE"
                ]

            )

        )


        # ==============================================
        # QUANTITY
        # ==============================================

        quantity = (

            st.number_input(

                "Number of Containers",

                min_value=1,

                value=1,

                step=1

            )

        )


        freight_total = (

            selected_rate

            *

            quantity

        )


        st.divider()


        # ==============================================
        # TRANSPORT
        # ==============================================

        st.markdown(
            """
            <div class="section-title">
            Landside Transport
            </div>
            """,
            unsafe_allow_html=True
        )


        transport_one, transport_two, transport_three = (

            st.columns(3)

        )


        with transport_one:

            selected_zone = (

                st.selectbox(

                    "Transport Zone",

                    [

                        "Zone A",

                        "Zone B",

                        "Zone C"

                    ]

                )

            )


        with transport_two:

            transport_type = (

                st.selectbox(

                    "Container Type",

                    [

                        "6M FCL (20ft)",

                        "12M FCL (40ft)"

                    ]

                )

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


        # ==============================================
        # TRANSPORT COLUMN MAP
        # ==============================================

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
            ][
                "weight"
            ]

        )


        total_column = (

            zone_columns[
                selected_zone
            ][
                "total"
            ]

        )


        if transport_type == "6M FCL (20ft)":

            section_name = "6M FCL"

        else:

            section_name = "12M FCL"


        section_rows = []


        for row_number in range(

            len(transport)

        ):


            row_text = " ".join(

                clean_text(value)

                for value in

                transport.iloc[
                    row_number
                ].tolist()

            ).upper()


            if section_name in row_text:

                section_rows.append(

                    row_number

                )


        transport_rate = None

        applied_weight_break = None


        if section_rows:


            start_row = (

                section_rows[0]

                +

                1

            )


            for row_number in range(

                start_row,

                min(

                    start_row + 10,

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


                if weight_text == "":

                    continue


                upper_text = (

                    weight_text.upper()

                )


                if (

                    "FCL"

                    in

                    upper_text

                ):

                    if row_number > start_row:

                        break


                weight_range = (

                    get_weight_range(

                        weight_text

                    )

                )


                if weight_range:


                    minimum = (

                        weight_range[0]

                    )


                    maximum = (

                        weight_range[1]

                    )


                    if (

                        cargo_weight

                        >=

                        minimum

                        and

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


        if transport_rate is None:

            transport_total = 0.0

        else:

            transport_total = (

                transport_rate

                *

                quantity

            )


        # ==============================================
        # FINAL SUMMARY
        # ==============================================

        st.divider()


        st.markdown(
            """
            <div class="section-title">
            Quote Summary
            </div>
            """,
            unsafe_allow_html=True
        )


        summary_one, summary_two = (

            st.columns(2)

        )


        with summary_one:


            st.markdown(

                f"""

                <div class=
                "summary-card">

                <div class=
                "summary-heading">

                OCEAN FREIGHT

                </div>

                <div class=
                "summary-value">

                USD {freight_total:,.2f}

                </div>

                <div class=
                "currency-note">

                {selected_line}

                <br>

                {selected_equipment}

                <br>

                Routing:
                {selected_routing}

                <br>

                Frequency:
                {selected_frequency}

                <br>

                Containers:
                {quantity}

                </div>

                </div>

                """,

                unsafe_allow_html=True

            )


        with summary_two:


            if transport_rate is None:

                transport_display = (
                    "Rate Not Found"
                )

            else:

                transport_display = (

                    f"R {transport_total:,.2f}"

                )


            st.markdown(

                f"""

                <div class=
                "summary-card">

                <div class=
                "summary-heading">

                LANDSIDE TRANSPORT

                </div>

                <div class=
                "summary-value">

                {transport_display}

                </div>

                <div class=
                "currency-note">

                {selected_zone}

                <br>

                {transport_type}

                <br>

                Cargo Weight:
                {cargo_weight:,.1f} Tons

                <br>

                Weight Break:
                {applied_weight_break}

                <br>

                Containers:
                {quantity}

                </div>

                </div>

                """,

                unsafe_allow_html=True

            )


        st.info(

            "Ocean freight is displayed "

            "in USD and landside transport "

            "is displayed in ZAR. "

            "The currencies are kept "

            "separate."

        )


else:


    st.markdown(

        """

        <div style="

        background:
        rgba(4, 27, 57, 0.80);

        border:
        1px solid
        rgba(75, 188, 255, 0.30);

        border-radius: 18px;

        padding: 25px;

        color: #cceeff;

        text-align: center;

        ">

        Enter both the POL and POD

        to display all available

        shipping-line options.

        </div>

        """,

        unsafe_allow_html=True

    )
