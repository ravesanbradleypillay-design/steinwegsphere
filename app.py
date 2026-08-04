
import streamlit as st
import pandas as pd
import os
import re


# ==================================================
# PAGE SETTINGS
# ==================================================

st.set_page_config(
    page_title="Steinweg Master Freight Calculator",
    page_icon="🚢",
    layout="wide"
)

st.title("STEINWEG")
st.subheader("Master Freight & Transport Calculator")

st.caption(
    "Ocean freight is displayed in USD. "
    "Transport is displayed in ZAR."
)

st.divider()


# ==================================================
# FILE NAMES
# ==================================================

FREIGHT_FILE = "freight rates.xlsx"
TRANSPORT_FILE = "transport rates.xls"


# ==================================================
# CHECK FILES
# ==================================================

if not os.path.exists(FREIGHT_FILE):
    st.error(
        f"Missing file: {FREIGHT_FILE}"
    )
    st.stop()

if not os.path.exists(TRANSPORT_FILE):
    st.error(
        f"Missing file: {TRANSPORT_FILE}"
    )
    st.stop()


# ==================================================
# LOAD EXCEL FILES
# ==================================================

@st.cache_data
def load_data():

    freight_data = pd.read_excel(
        FREIGHT_FILE,
        engine="openpyxl"
    )

    # Transport sheet is a wide tariff table,
    # so do not use the first row as headings.
    transport_data = pd.read_excel(
        TRANSPORT_FILE,
        engine="xlrd",
        header=None
    )

    return freight_data, transport_data


freight, transport = load_data()


# ==================================================
# CLEAN FREIGHT HEADINGS
# ==================================================

freight.columns = (
    freight.columns
    .astype(str)
    .str.strip()
    .str.upper()
)


# ==================================================
# REQUIRED FREIGHT COLUMNS
# ==================================================

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
        "The freight spreadsheet is missing: "
        + ", ".join(missing_columns)
    )

    st.write(
        "Columns found:"
    )

    st.write(
        freight.columns.tolist()
    )

    st.stop()


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

    # Handles values such as:
    # 2,500.00
    # 2500.00
    # 2 500,00

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


# ==================================================
# CLEAN FREIGHT DATA
# ==================================================

for column in [
    "LINE",
    "POL",
    "POD",
    "EQUIP TYPE",
    "PREPAID / COLLECT"
]:

    freight[column] = (

        freight[column]

        .fillna("")

        .astype(str)

        .str.strip()

    )


# ==================================================
# OCEAN FREIGHT SEARCH
# ==================================================

st.header(
    "Ocean Freight Search"
)


search_col1, search_col2 = (
    st.columns(2)
)


with search_col1:

    user_pol = st.text_input(
        "POL",
        placeholder=(
            "Enter Port of Loading"
        )
    )


with search_col2:

    user_pod = st.text_input(
        "POD",
        placeholder=(
            "Enter Port of Discharge"
        )
    )


st.caption(
    "You can enter part of the port name. "
    "For example: Durban, Shanghai, "
    "Mundra or Jebel Ali."
)


# ==================================================
# SEARCH FREIGHT ROUTES
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
            "No freight rates were found "
            "for this POL and POD."
        )


    else:

        st.success(

            f"{len(route_matches)} "
            "matching freight rate(s) found."

        )


        select_col1, select_col2, select_col3 = (
            st.columns(3)
        )


        with select_col1:

            shipping_lines = sorted(

                route_matches["LINE"]

                .replace(
                    "",
                    pd.NA
                )

                .dropna()

                .unique()

            )


            selected_line = st.selectbox(

                "Shipping Line",

                shipping_lines

            )


        line_matches = route_matches[

            route_matches["LINE"]
            ==
            selected_line

        ].copy()


        with select_col2:

            equipment_options = sorted(

                line_matches[
                    "EQUIP TYPE"
                ]

                .replace(
                    "",
                    pd.NA
                )

                .dropna()

                .unique()

            )


            selected_equipment = (

                st.selectbox(

                    "Equipment Type",

                    equipment_options

                )

            )


        equipment_matches = (

            line_matches[

                line_matches[
                    "EQUIP TYPE"
                ]

                ==

                selected_equipment

            ]

            .copy()

        )


        with select_col3:

            payment_options = sorted(

                equipment_matches[

                    "PREPAID / COLLECT"

                ]

                .replace(
                    "",
                    pd.NA
                )

                .dropna()

                .unique()

            )


            selected_payment = (

                st.selectbox(

                    "Prepaid / Collect",

                    payment_options

                )

            )


        selected_freight = (

            equipment_matches[

                equipment_matches[

                    "PREPAID / COLLECT"

                ]

                ==

                selected_payment

            ]

            .copy()

        )


        # ==================================================
        # SHOW FREIGHT OPTIONS
        # ==================================================

        st.subheader(
            "Available Shipping Options"
        )


        display_columns = [

            "LINE",

            "POL",

            "POD",

            "EQUIP TYPE",

            "ALL IN RATE",

            "ROUTING",

            "FREQUENCY"

        ]


        st.dataframe(

            selected_freight[

                display_columns

            ],

            use_container_width=True,

            hide_index=True

        )


        # Use the first matching rate.
        # The table above still displays all
        # matching options.

        selected_row = (

            selected_freight.iloc[0]

        )


        freight_rate = (

            money_to_number(

                selected_row[

                    "ALL IN RATE"

                ]

            )

        )


        if freight_rate is None:

            freight_rate = 0.0


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


        # ==================================================
        # CONTAINER QUANTITY
        # ==================================================

        quantity = st.number_input(

            "Number of Containers",

            min_value=1,

            value=1,

            step=1

        )


        freight_total = (

            freight_rate

            *

            quantity

        )


        st.divider()


        # ==================================================
        # TRANSPORT
        # ==================================================

        st.header(
            "Transport"
        )


        transport_col1, transport_col2, transport_col3 = (

            st.columns(3)

        )


        with transport_col1:

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


        with transport_col2:

            transport_type = (

                st.selectbox(

                    "Container / Transport Type",

                    [

                        "6M FCL (20ft)",

                        "12M FCL (40ft)"

                    ]

                )

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


        st.caption(

            "20ft containers use 6M FCL. "

            "40ft containers use 12M FCL."

        )


        # ==================================================
        # TRANSPORT LOOKUP
        # ==================================================

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


            if (

                section_name

                in

                row_text

            ):

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

                    or

                    "BREAKBULK"

                    in

                    upper_text

                ):

                    if (

                        row_number

                        >

                        start_row

                    ):

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

            transport_message = (

                "No matching transport "

                "rate was found."

            )

        else:

            transport_total = (

                transport_rate

                *

                quantity

            )

            transport_message = (

                "Transport rate found."

            )


        # ==================================================
        # FINAL RESULTS
        # ==================================================

        st.divider()

        st.header(
            "Quote Summary"
        )


        result_col1, result_col2 = (

            st.columns(2)

        )


        with result_col1:

            st.subheader(
                "Ocean Freight"
            )


            st.metric(

                "Freight Total",

                f"USD {freight_total:,.2f}"

            )


            st.write(

                f"**Shipping Line:** "

                f"{selected_line}"

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

                f"**Equipment:** "

                f"{selected_equipment}"

            )


            st.write(

                f"**Containers:** "

                f"{quantity}"

            )


        with result_col2:

            st.subheader(
                "Transport"
            )


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


            if applied_weight_break:

                st.write(

                    f"**Weight Break:** "

                    f"{applied_weight_break}"

                )


            st.write(

                f"**Status:** "

                f"{transport_message}"

            )


        st.info(

            "Ocean freight is shown in USD "

            "and transport is shown in ZAR. "

            "They are intentionally not added "

            "together."

        )


else:

    st.info(

        "Enter both POL and POD "

        "to search available shipping lines."

    )
