

"""
Tải dữ liệu THEO THÁNG của Oracle (ORCL) từ 01/2015 đến hiện tại, xuất ra Excel.

Cài thư viện (chạy 1 lần, dùng đúng bản Python bạn đang chạy script):
    /usr/local/bin/python3 -m pip install yfinance pandas openpyxl

Chạy:
    /usr/local/bin/python3 oracle_monthly_to_excel.py
"""

import os
from datetime import date

import pandas as pd
import yfinance as yf

TICKER = "ORCL"
START = "2015-01-01"
END = date.today().isoformat()
OUTPUT = os.path.join(os.path.expanduser("~/Desktop"), "ORCL_monthly_2015_to_now.xlsx")


def main():
    print(f"Đang tải dữ liệu THÁNG của {TICKER} từ {START} đến {END} ...")
    df = yf.download(
        TICKER,
        start=START,
        end=END,
        interval="1mo",      # dữ liệu theo tháng
        auto_adjust=False,   # giữ cả 'Close' và 'Adj Close'
        progress=False,
    )

    if df.empty:
        raise SystemExit("Không tải được dữ liệu. Kiểm tra Internet rồi thử lại.")

    # yfinance bản mới trả về cột nhiều tầng -> làm phẳng
    if isinstance(df.columns, pd.MultiIndex):
        df.columns = df.columns.get_level_values(0)

    df.index = pd.to_datetime(df.index).tz_localize(None)
    df.index.name = "Month"
    df = df.reset_index()

    cols = ["Month", "Open", "High", "Low", "Close", "Adj Close", "Volume"]
    df = df[[c for c in cols if c in df.columns]]

    # % thay đổi giá đóng cửa so với tháng trước
    df["Monthly_Return_%"] = (df["Adj Close"].pct_change() * 100).round(2)

    # Tổng hợp theo năm
    yearly = (
        df.assign(Year=df["Month"].dt.year)
        .groupby("Year")
        .agg(
            Open_dau_nam=("Open", "first"),
            Close_cuoi_nam=("Close", "last"),
            Cao_nhat=("High", "max"),
            Thap_nhat=("Low", "min"),
            Tong_Volume=("Volume", "sum"),
        )
        .round(2)
        .reset_index()
    )
    yearly["Thay_doi_%"] = (
        (yearly["Close_cuoi_nam"] / yearly["Open_dau_nam"] - 1) * 100
    ).round(2)

    with pd.ExcelWriter(OUTPUT, engine="openpyxl", datetime_format="yyyy-mm") as writer:
        df.to_excel(writer, sheet_name="Monthly", index=False)
        yearly.to_excel(writer, sheet_name="Yearly summary", index=False)

        for ws in writer.book.worksheets:
            ws.freeze_panes = "A2"
            for col in ws.columns:
                width = max(len(str(c.value)) if c.value is not None else 0 for c in col)
                ws.column_dimensions[col[0].column_letter].width = min(width + 3, 22)

    print(f"Xong! {len(df)} tháng dữ liệu đã lưu tại: {OUTPUT}")


if __name__ == "__main__":
    main()
    
    """
Tải dữ liệu THEO THÁNG của NVIDIA (NVDA) từ 01/2015 đến hiện tại, xuất ra Excel.

Cài thư viện (chạy 1 lần, dùng đúng bản Python bạn đang chạy script):
    /usr/local/bin/python3 -m pip install yfinance pandas openpyxl

Chạy:
    /usr/local/bin/python3 nvidia_monthly_to_excel.py
"""

import os
from datetime import date

import pandas as pd
import yfinance as yf

TICKER = "NVDA"
START = "2015-01-01"
END = date.today().isoformat()
OUTPUT = os.path.join(os.path.expanduser("~/Desktop"), "NVDA_monthly_2015_to_now.xlsx")


def main():
    print(f"Đang tải dữ liệu THÁNG của {TICKER} từ {START} đến {END} ...")
    df = yf.download(
        TICKER,
        start=START,
        end=END,
        interval="1mo",      # dữ liệu theo tháng
        auto_adjust=False,   # giữ cả 'Close' và 'Adj Close'
        progress=False,
    )

    if df.empty:
        raise SystemExit("Không tải được dữ liệu. Kiểm tra Internet rồi thử lại.")

    # yfinance bản mới trả về cột nhiều tầng -> làm phẳng
    if isinstance(df.columns, pd.MultiIndex):
        df.columns = df.columns.get_level_values(0)

    df.index = pd.to_datetime(df.index).tz_localize(None)
    df.index.name = "Month"
    df = df.reset_index()

    cols = ["Month", "Open", "High", "Low", "Close", "Adj Close", "Volume"]
    df = df[[c for c in cols if c in df.columns]]

    # % thay đổi giá đóng cửa so với tháng trước
    df["Monthly_Return_%"] = (df["Adj Close"].pct_change() * 100).round(2)

    # Tổng hợp theo năm
    yearly = (
        df.assign(Year=df["Month"].dt.year)
        .groupby("Year")
        .agg(
            Open_dau_nam=("Open", "first"),
            Close_cuoi_nam=("Close", "last"),
            Cao_nhat=("High", "max"),
            Thap_nhat=("Low", "min"),
            Tong_Volume=("Volume", "sum"),
        )
        .round(2)
        .reset_index()
    )
    yearly["Thay_doi_%"] = (
        (yearly["Close_cuoi_nam"] / yearly["Open_dau_nam"] - 1) * 100
    ).round(2)

    with pd.ExcelWriter(OUTPUT, engine="openpyxl", datetime_format="yyyy-mm") as writer:
        df.to_excel(writer, sheet_name="Monthly", index=False)
        yearly.to_excel(writer, sheet_name="Yearly summary", index=False)

        for ws in writer.book.worksheets:
            ws.freeze_panes = "A2"
            for col in ws.columns:
                width = max(len(str(c.value)) if c.value is not None else 0 for c in col)
                ws.column_dimensions[col[0].column_letter].width = min(width + 3, 22)

    print(f"Xong! {len(df)} tháng dữ liệu đã lưu tại: {OUTPUT}")


if __name__ == "__main__":
    main()