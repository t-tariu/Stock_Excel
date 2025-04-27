from pykrx import stock
from openpyxl import load_workbook
from openpyxl.styles import Font
from datetime import datetime, timedelta

file_path = "stock.xlsx"

wb = load_workbook(file_path)
ws = wb.active

# today date
today = datetime.today()

# pykrx date
def to_yyyymmdd(date):
    return date.strftime("%Y%m%d")

# Find recent date
def get_last_trading_day(start_date, ticker):
    date = start_date
    while True:
        date_str = to_yyyymmdd(date)
        try:
            df = stock.get_market_ohlcv(date_str, date_str, ticker)
            if not df.empty:
                return date_str
        except:
            pass
        date -= timedelta(days=1)

# format_market_cap
def format_market_cap(market_cap_won):
    market_cap_억 = market_cap_won // 100_000_000  # 원 → 억
    if market_cap_억 >= 10_000:  # 1조 = 10,000억
        market_cap_조 = market_cap_억 / 10_000
        return f"{market_cap_조:.2f}조"
    else:
        return f"{market_cap_억}억"

# Write in Excel
for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
    for cell in row:
        ticker = cell.value

        if ticker and isinstance(ticker, str) and len(ticker) == 6 and ticker.isdigit():
            try:
                # recent trade date
                last_trading_day = get_last_trading_day(today, ticker)

                # get data
                ohlcv = stock.get_market_ohlcv(last_trading_day, last_trading_day, ticker)
                price = f"{ohlcv['종가'].values[0]:,}"

                change_value = ohlcv['등락률'].values[0]  
                change_text = f"{change_value:.2f}%"    

                fundamental = stock.get_market_fundamental(last_trading_day, last_trading_day, ticker)
                per = fundamental['PER'].values[0]

                market_cap_data = stock.get_market_cap(last_trading_day, last_trading_day, ticker)
                market_cap_raw = int(market_cap_data['시가총액'].values[0])
                market_cap = format_market_cap(market_cap_raw)

                # 엑셀에 작성
                ws.cell(row=cell.row + 1, column=cell.column, value=price)

                change_cell = ws.cell(row=cell.row + 2, column=cell.column, value=change_text)
                # 등락률 양/음에 따라 색깔 적용
                if change_value > 0:
                    change_cell.font = Font(color="FF0000")  # RED
                elif change_value < 0:
                    change_cell.font = Font(color="0000FF")  # BLUE
                else:
                    change_cell.font = Font(color="000000")  # BLACK

                ws.cell(row=cell.row + 3, column=cell.column, value=per)
                ws.cell(row=cell.row + 4, column=cell.column, value=market_cap)

            except Exception as e:
                print(f"Error processing ticker {ticker}: {e}")

# 저장
wb.save(file_path)
