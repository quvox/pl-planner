from typing import Tuple
import datetime
from dateutil.relativedelta import relativedelta


MAPPING1 = {"plan": "計画", "performance": "実績"}
MAPPING2 = {"計画": "plan", "実績": "performance"}


def get_yyyymm(dt: datetime.datetime) -> str:
    return f"{dt.year}/{dt.month}"


def convert_from_yyyymm(yyyymm: str) -> datetime.datetime:
    y = int(yyyymm[:4])
    if "/" in yyyymm:
        m = int(yyyymm[5:])
    else:
        m = int(yyyymm[4:])
    return datetime.datetime(y, m, 1)


def get_term_start_month(dt: datetime.datetime, settlement_month: int, months_after=0) -> datetime.datetime:
    """指定した年月(dt)、またはそこから指定月数(months_before)だけ後の年月が属する期の期初の年月を返す"""
    d = dt + relativedelta(months=months_after)
    d = datetime.datetime(d.year, d.month, 1)
    te = datetime.datetime(d.year, settlement_month, 1)  # 期末
    if (d - te).days < 0:
        return te - relativedelta(months=11)  # 期初
    return te + relativedelta(months=1)


def get_term_end_month(dt: datetime.datetime, settlement_month: int, months_after=0) -> datetime.datetime:
    """指定した年月(dt)、またはそこから指定月数(months_after)だけ後の年月が属する期の期末の年月を返す"""
    d = dt + relativedelta(months=months_after)
    d = datetime.datetime(d.year, d.month, 1)
    te = datetime.datetime(d.year, settlement_month, 1)  # 期末
    if (d - te).days > 0:
        te = datetime.datetime(te.year+1, settlement_month, 1)  # 期末
    return te


def create_header_labels(start: datetime.datetime, end: datetime.datetime, settlement_month: int) -> list[str]:
    header = list()
    delta_month = abs(end.year - start.year)*12 + abs(end.month - start.month) + 1  # startからendまでの月数
    if end.month - start.month < 0: delta_month -= 2
    fiscal_years = 0
    for dm in range(delta_month):
        dt = start + relativedelta(months=dm)
        header.append(f"{dt.year}/{dt.month:02d}")
        if dt.month == settlement_month:
            fiscal_years += 1
            header.append(f"{dt.year}.{dt.month:02d}決算")
    return header
