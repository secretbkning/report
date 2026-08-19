"""해외법인 손익 리포트 계산 로직.

Dummy Data_강병규.xlsx > Sheet1 을 읽어 Report_P1 / Report_P2 (엑셀 버전)와
동일한 정의로 당월/누계 기준 계획·실적·차이·차이내역을 계산한다.
엑셀 버전(build_stage1/2/3.py, Report_P1/P2/Calc 시트)과 계산 정의를 반드시
동일하게 유지할 것 - 여기 또는 거기 한쪽만 바뀌면 두 버전 숫자가 어긋난다.
"""

from __future__ import annotations

import datetime as dt
from pathlib import Path

import pandas as pd

DATA_PATH = Path(__file__).with_name("Dummy Data_강병규.xlsx")

CORPS = [
    "중국1", "중국2", "중국3", "중국4",
    "인도1", "인도2", "인도3",
    "유럽1", "유럽2", "유럽3", "유럽4",
    "미주1", "미주2", "미주3", "미주4",
]


def load_raw(path: Path = DATA_PATH) -> pd.DataFrame:
    """Sheet1 원본 데이터 로드. 헤더는 3행(0-index 2), 데이터는 4행부터."""
    df = pd.read_excel(path, sheet_name="Sheet1", header=2)
    df["년월"] = pd.to_datetime(df["년월"])
    return df


def period_options(df: pd.DataFrame) -> list[str]:
    """데이터에 존재하는 (년월 x 당월/누계) 조합 목록 - 월이 늘어나도 자동 반영."""
    months = sorted(df["년월"].dropna().unique())
    opts = []
    for m in months:
        ts = pd.Timestamp(m)
        label = f"{ts.year}-{ts.month:02d}"
        opts.append(f"{label} 당월")
        opts.append(f"{label} 누계")
    return opts


def parse_period(selected: str) -> tuple[int, int, str]:
    ym, mode = selected.split(" ")
    year, month = ym.split("-")
    return int(year), int(month), mode


def period_range(year: int, month: int, mode: str) -> tuple[pd.Timestamp, pd.Timestamp]:
    end = pd.Timestamp(dt.date(year, month, 1))
    start = pd.Timestamp(dt.date(year, 1, 1)) if mode == "누계" else end
    return start, end


def _safe_div(a: float, b: float) -> float:
    return a / b if b else 0.0


def _sum_item(df: pd.DataFrame, start, end, plan_actual: str, item: str, corp: str) -> float:
    mask = (
        (df["년월"] >= start)
        & (df["년월"] <= end)
        & (df["계획실적"] == plan_actual)
        & (df["구분"] == item)
    )
    return float(df.loc[mask, corp].sum())  # NaN(공란)은 sum()에서 자동으로 0 취급


def build_calc(df: pd.DataFrame, start, end) -> pd.DataFrame:
    """법인별 원천 pull + 파생값(롤마진, 제경비) - 엑셀의 숨김 Calc 시트와 동일."""
    rows = []
    for corp in CORPS:
        raw = {
            "계획판매량": _sum_item(df, start, end, "계획", "판매량", corp),
            "실적판매량": _sum_item(df, start, end, "실적", "판매량", corp),
            "계획매출액": _sum_item(df, start, end, "계획", "매출액", corp),
            "실적매출액": _sum_item(df, start, end, "실적", "매출액", corp),
            "계획재료비": _sum_item(df, start, end, "계획", " - 재료비", corp),
            "실적재료비": _sum_item(df, start, end, "실적", " - 재료비", corp),
            "계획노무비": _sum_item(df, start, end, "계획", " - 노무비", corp),
            "실적노무비": _sum_item(df, start, end, "실적", " - 노무비", corp),
            "계획경비": _sum_item(df, start, end, "계획", " - 경비", corp),
            "실적경비": _sum_item(df, start, end, "실적", " - 경비", corp),
            "계획판관비": _sum_item(df, start, end, "계획", "판관비", corp),
            "실적판관비": _sum_item(df, start, end, "실적", "판관비", corp),
            "계획영업외": _sum_item(df, start, end, "계획", "영업외", corp),
            "실적영업외": _sum_item(df, start, end, "실적", "영업외", corp),
        }
        raw["계획롤마진"] = _safe_div(raw["계획매출액"], raw["계획판매량"]) - _safe_div(
            raw["계획재료비"], raw["계획판매량"]
        )
        raw["실적롤마진"] = _safe_div(raw["실적매출액"], raw["실적판매량"]) - _safe_div(
            raw["실적재료비"], raw["실적판매량"]
        )
        raw["계획제경비"] = raw["계획노무비"] + raw["계획경비"] + raw["계획판관비"]
        raw["실적제경비"] = raw["실적노무비"] + raw["실적경비"] + raw["실적판관비"]
        raw["법인"] = corp
        rows.append(raw)
    return pd.DataFrame(rows).set_index("법인")


def build_page1(df: pd.DataFrame, start, end) -> pd.DataFrame:
    rows = []
    for corp in CORPS:
        plan_qty = _sum_item(df, start, end, "계획", "판매량", corp)
        plan_rev = _sum_item(df, start, end, "계획", "매출액", corp)
        plan_ni = _sum_item(df, start, end, "계획", "경상이익", corp)
        act_qty = _sum_item(df, start, end, "실적", "판매량", corp)
        act_rev = _sum_item(df, start, end, "실적", "매출액", corp)
        act_ni = _sum_item(df, start, end, "실적", "경상이익", corp)
        rows.append(
            {
                "법인": corp,
                "계획-판매량": plan_qty,
                "계획-매출액": plan_rev,
                "계획-경상이익": plan_ni,
                "실적-판매량": act_qty,
                "실적-매출액": act_rev,
                "실적-경상이익": act_ni,
                "차이-경상이익": act_ni - plan_ni,
            }
        )
    p1 = pd.DataFrame(rows)
    total = p1.drop(columns="법인").sum()
    total["법인"] = "합계"
    p1 = pd.concat([p1, pd.DataFrame([total])], ignore_index=True)
    return p1


def build_page2(page1: pd.DataFrame, calc: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for corp in CORPS:
        c = calc.loc[corp]
        ni_diff = float(page1.loc[page1["법인"] == corp, "차이-경상이익"].iloc[0])
        qty_var = c["계획롤마진"] * (c["실적판매량"] - c["계획판매량"])
        margin_var = (c["실적롤마진"] - c["계획롤마진"]) * c["실적판매량"]
        cost_var = c["계획제경비"] - c["실적제경비"]
        nonop_var = c["실적영업외"] - c["계획영업외"]
        rows.append(
            {
                "법인": corp,
                "경상이익차이": ni_diff,
                "물량차이": qty_var,
                "롤마진차이": margin_var,
                "제경비차이": cost_var,
                "영업외차이": nonop_var,
            }
        )
    p2 = pd.DataFrame(rows)
    total = p2.drop(columns="법인").sum()
    total["법인"] = "합계"
    p2 = pd.concat([p2, pd.DataFrame([total])], ignore_index=True)
    p2["검증(차이내역합-경상이익차이)"] = (
        p2["물량차이"] + p2["롤마진차이"] + p2["제경비차이"] + p2["영업외차이"] - p2["경상이익차이"]
    ).round(6)
    return p2


def compute_report(selected_period: str, path: Path = DATA_PATH):
    df = load_raw(path)
    year, month, mode = parse_period(selected_period)
    start, end = period_range(year, month, mode)
    page1 = build_page1(df, start, end)
    calc = build_calc(df, start, end)
    page2 = build_page2(page1, calc)
    return page1, page2
