import streamlit as st
import yfinance as yf
import pandas as pd
from io import BytesIO

st.set_page_config(
    page_title="FNO Last 20 Days",
    layout="wide"
)

st.title("NSE Stocks – Last 20 Trading Days")
st.caption("SMA50 + EMA50 | Date-wise Excel Output")


# ============================================================
# SYMBOL LIST
# ============================================================

symbols = """
FEDERALBNK
KALYANKJIL
BHARATFORG
TECHM
KOTAKBANK
EXIDEIND
NAM-INDIA
SONACOMS
JIOFIN
TCS
DLF
360ONE
PRESTIGE
ICICIBANK
RELIANCE
PERSISTENT
KFINTECH
HINDUNILVR
BRITANNIA
M&M
COLPAL
ADANIENSOL
EICHERMOT
HDFCAMC
MPHASIS
HAVELLS
BAJFINANCE
AXISBANK
TVSMOTOR
HDFCBANK
OBEROIRLTY
HAL
MOTILALOFS
INDIANB
OIL
HCLTECH
RBLBANK
SBICARD
RECLTD
NUVAMA
GRASIM
INDUSINDBK
SBIN
JSWSTEEL
TMPV
LT
SWIGGY
INFY
GODREJCP
BAJAJ-AUTO
OFSS
BPCL
IDFCFIRSTB
POWERGRID
POLICYBZR
PFC
BAJAJFINSV
AMBUJACEM
PNB
LODHA
MARICO
ASIANPAINT
CHOLAFIN
MOTHERSON
ADANIPORTS
IREDA
CDSL
KPITTECH
LICHSGFIN
ITC
BEL
UNOMINDA
SBILIFE
LTM
ASHOKLEY
GODREJPROP
HINDPETRO
NESTLEIND
MARUTI
ONGC
ABCAPITAL
SUZLON
WAAREEENER
HEROMOTOCO
TATASTEEL
TITAN
SIEMENS
ADANIENT
SOLARINDS
ETERNAL
PNBHOUSING
MANAPPURAM
COALINDIA
GAIL
SHRIRAMFIN
PIDILITIND
TATACONSUM
CAMS
DMART
ASTRAL
TATAPOWER
NBCC
IRFC
PHOENIXLTD
LTF
APLAPOLLO
IOC
NHPC
NTPC
BAJAJHLDNG
JINDALSTEL
DALBHARAT
BANKINDIA
UNITDSPR
INDIGO
MUTHOOTFIN
COCHINSHIP
INDHOTEL
SAIL
JSWENERGY
NYKAA
PIIND
UNIONBANK
GODFRYPHLP
RVNL
PETRONET
HINDZINC
DABUR
RADICO
CONCOR
CANBK
BANKBARODA
HDFCLIFE
TRENT
BIOCON
INDUSTOWER
BANDHANBNK
BDL
AUBANK
CIPLA
BHARTIARTL
ULTRACEMCO
MAXHEALTH
BOSCHLTD
APOLLOHOSP
AUROPHARMA
DIVISLAB
SUNPHARMA
ICICIGI
INOXWIND
ZYDUSLIFE
YESBANK
TATAELXSI
DELHIVERY
LICI
SRF
ADANIPOWER
FORCEMOT
WIPRO
NMDC
VBL
MAZDOCK
JUBLFOOD
DRREDDY
FORTIS
IDEA
GMRAIRPORT
PAGEIND
COFORGE
DIXON
HYUNDAI
PGEL
PAYTM
HINDALCO
KAYNES
CROMPTON
TIINDIA
IEX
VMM
SHREECEM
ANGELONE
CGPOWER
ADANIGREEN
UPL
NAUKRI
LAURUSLABS
PREMIERENE
ABB
VEDL
GLENMARK
ICICIPRULI
LUPIN
MCX
VOLTAS
MANKIND
ALKEM
MFSL
CUMMINSIND
BSE
SUPREMEIND
AMBER
BHEL
PATANJALI
KEI
NATIONALUM
BLUESTARCO
POLYCAB
GVT&D
POWERINDIA
TORNTPHARM
""".split()


tickers = [symbol + ".NS" for symbol in symbols]


# ============================================================
# CREATE EXCEL
# ============================================================

def create_excel():

    progress = st.progress(0)
    status = st.empty()

    status.info("Downloading 150 days of NSE data...")

    data = yf.download(
        tickers=tickers,
        period="150d",
        interval="1d",
        auto_adjust=False,
        group_by="ticker",
        threads=True,
        progress=False
    )

    progress.progress(30)

    all_data = []

    status.info("Processing stocks...")

    for i, symbol in enumerate(symbols):

        ticker = symbol + ".NS"

        try:

            if ticker not in data.columns.get_level_values(0):
                continue

            stock_data = data[ticker].copy()

            stock_data = stock_data[
                stock_data["Close"].notna()
            ].copy()

            if stock_data.empty:
                continue

            stock_data = stock_data.reset_index()

            stock_data["Symbol"] = symbol

            # =================================================
            # PREVIOUS CLOSE
            # =================================================

            stock_data["Prev. Close"] = (
                stock_data["Close"].shift(1)
            )

            # =================================================
            # CHANGE
            # =================================================

            stock_data["Chg"] = (
                stock_data["Close"]
                - stock_data["Prev. Close"]
            )

            # =================================================
            # % CHANGE
            # =================================================

            stock_data["%Chg"] = (
                stock_data["Chg"]
                / stock_data["Prev. Close"]
            ) * 100

            # =================================================
            # SMA50
            # =================================================

            stock_data["SMA50"] = (
                stock_data["Close"]
                .rolling(
                    window=50,
                    min_periods=50
                )
                .mean()
            )

            # =================================================
            # EMA50
            # =================================================

            stock_data["EMA50"] = (
                stock_data["Close"]
                .ewm(
                    span=50,
                    adjust=False
                )
                .mean()
            )

            # =================================================
            # CLOSE -> LTP
            # =================================================

            stock_data.rename(
                columns={
                    "Close": "LTP"
                },
                inplace=True
            )

            all_data.append(stock_data)

        except Exception:
            continue

        progress.progress(
            30 + int((i + 1) / len(symbols) * 40)
        )


    if not all_data:
        st.error("No stock data was downloaded.")
        return None


    # ============================================================
    # COMBINE
    # ============================================================

    final_data = pd.concat(
        all_data,
        ignore_index=True
    )


    # ============================================================
    # REMOVE TIMEZONE
    # ============================================================

    final_data["Date"] = pd.to_datetime(
        final_data["Date"]
    )

    try:
        final_data["Date"] = (
            final_data["Date"]
            .dt.tz_localize(None)
        )
    except TypeError:
        pass


    # ============================================================
    # LAST 20 TRADING DAYS
    # ============================================================

    last_20_dates = sorted(
        final_data["Date"]
        .drop_duplicates()
        .tail(20)
    )


    status.info("Creating Excel workbook...")


    # ============================================================
    # EXCEL IN MEMORY
    # ============================================================

    output_file = BytesIO()


    with pd.ExcelWriter(
        output_file,
        engine="openpyxl"
    ) as writer:

        for date in reversed(last_20_dates):

            date_data = final_data[
                final_data["Date"] == date
            ].copy()


            # =================================================
            # REQUIRED COLUMN SEQUENCE
            # =================================================

            output = pd.DataFrame({

                "Symbol":
                    date_data["Symbol"],

                "Open":
                    date_data["Open"],

                "High":
                    date_data["High"],

                "Low":
                    date_data["Low"],

                "Prev. Close":
                    date_data["Prev. Close"],

                "LTP":
                    date_data["LTP"],

                "":
                    "",

                "Chg":
                    date_data["Chg"],

                "%Chg":
                    date_data["%Chg"],

                "Volume":
                    date_data["Volume"],

                "SMA50":
                    date_data["SMA50"],

                "EMA50":
                    date_data["EMA50"]
            })


            # =================================================
            # ROUND PRICE VALUES
            # =================================================

            price_columns = [
                "Open",
                "High",
                "Low",
                "Prev. Close",
                "LTP",
                "Chg",
                "SMA50",
                "EMA50"
            ]


            output[price_columns] = (
                output[price_columns].round(2)
            )


            output["%Chg"] = (
                output["%Chg"].round(2)
            )


            # =================================================
            # ORIGINAL SYMBOL ORDER
            # =================================================

            symbol_order = {
                symbol: i
                for i, symbol in enumerate(symbols)
            }

            output["Order"] = (
                output["Symbol"]
                .map(symbol_order)
            )


            output = (
                output
                .sort_values("Order")
                .drop(columns=["Order"])
            )


            # =================================================
            # SHEET NAME
            # =================================================

            sheet_name = date.strftime("%d %b")


            output.to_excel(
                writer,
                sheet_name=sheet_name,
                index=False
            )


    output_file.seek(0)

    progress.progress(100)

    status.success("✅ Excel created successfully")

    return output_file.getvalue()


# ============================================================
# BUTTON
# ============================================================

if st.button(
    "📊 Download Last 20 Trading Days",
    type="primary"
):

    with st.spinner(
        "Please wait... Downloading and processing NSE data..."
    ):

        excel_data = create_excel()


    if excel_data:

        st.download_button(
            label="⬇️ Download FNO_Last_20_Days_Date_Wise.xlsx",
            data=excel_data,
            file_name="FNO_Last_20_Days_Date_Wise.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            )
        )
