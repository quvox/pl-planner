from typing import Union, Tuple
import os
import sys
import datetime
import openpyxl

from . import common
from pldata import ProfitData, LossData, ProfitDataItem, LabelManager, MonthlyData
from excel import utils, table


def read_data_file(file_path: str, data_store: Union[dict, None], label_mgr: LabelManager) -> Tuple[datetime.datetime, datetime.datetime]:
    start_dt = None
    end_dt = None
    business = os.path.splitext(os.path.basename(file_path))[0]
    if business not in data_store["definition"]:
        # 設定ファイルに定義されていないものは無視する
        return start_dt, end_dt

    print(f" - reading: {business}.xlsx")
    workbook = openpyxl.load_workbook(file_path)
    for ws_name in workbook.sheetnames:  # 計画,実績
        ws = workbook[ws_name]

        data_store.setdefault(common.MAPPING2[ws_name], {}).setdefault(business, {}).setdefault("profit", {})
        data_store.setdefault(common.MAPPING2[ws_name], {}).setdefault(business, {}).setdefault("loss", {})

        # 一番上の表を読み込む
        pos = utils.find_column_numbers([business], ws)
        if business not in pos:
            print(f"XXX ファイル:{business}.xlsxの{ws_name}シートが不正です")
            continue
        start_dt, end_dt = _parse_data(ws, data_store, business, ws_name, label_mgr)

    return start_dt, end_dt


def _parse_data(ws: any, ds: dict, business: str, ws_name: str, label_mgr: LabelManager) -> Tuple[datetime.datetime, datetime.datetime]:
    """表を読み込む"""
    data_store = ds[common.MAPPING2[ws_name]][business]
    profit_label_num = len(label_mgr.get_all(business, "profit"))
    loss_label_num = len(label_mgr.get_all(business, "loss"))
    start_dt = None
    end_dt = None

    tbl = table.SingleTable(ws, (1, 2), 3)  # テーブルの起点(左上がA2のセル)、3セルを行ラベル用に使う

    # ヘッダ行を読む
    tbl.read_as_header()  # 1行目をヘッダとして読む
    for hdr in tbl.headers:
        if "決算" in hdr: continue
        end_dt = common.convert_from_yyyymm(hdr)
        if start_dt is None:
            start_dt = common.convert_from_yyyymm(hdr)

    # 売上サブテーブルを読む
    sales_tbl = tbl.add_sub_table("sales")
    sales_tbl.setup_dummy_aggregate_row()
    sales_tbl.read_as_row_labels(row_size=profit_label_num+1)  # 集計行を含めるため+1
    data_cols = sales_tbl.get_all_data()
    # -- 数値データを読み込む
    for yyyymm, data in data_cols.items():
        if "決算" in yyyymm: continue  # 決算列は無視して良い
        if start_dt is None:
            start_dt = common.convert_from_yyyymm(yyyymm)
        end_dt = common.convert_from_yyyymm(yyyymm)
        monthly_data = list()  # type: list[ProfitData]
        for dat in data:
            row_label = label_mgr.get(business, "profit", name=dat["label"][0])
            monthly_data.append(ProfitData(row_label, dat["value"]))

        if yyyymm not in data_store["profit"]:
            data_store["profit"][yyyymm] = MonthlyData(yyyymm, monthly_data)
        else:
            data_store["profit"][yyyymm].merge(monthly_data)

    tbl.add_blank_row()

    # 経費サブテーブルを読む
    expense_tbl = tbl.add_sub_table("expense")
    expense_tbl.setup_dummy_aggregate_row()
    expense_tbl.read_as_row_labels(row_size=loss_label_num+1)  # 集計行を含めるため+1
    data_cols = expense_tbl.get_all_data()
    # -- 数値データを読み込む
    for yyyymm, data in data_cols.items():
        if "決算" in yyyymm: continue  # 決算列は無視して良い
        monthly_data = list()  # type: list[LossData]
        for dat in data:
            row_label = label_mgr.get(business, "loss", group=dat["label"][0], account=dat["label"][1], category=dat["label"][2])
            monthly_data.append(LossData(row_label, dat["value"]))

        if yyyymm not in data_store["loss"]:
            data_store["loss"][yyyymm] = MonthlyData(yyyymm, monthly_data)
        else:
            data_store["loss"][yyyymm].merge(monthly_data)
    tbl.add_blank_row()

    return start_dt, end_dt


def update(data_store: dict):
    """データの集計や、全社共通シートに記載された経費を按分して各事業に振り分けたりする"""
    _divide_common_expense(data_store)
    _calculate_all_earnings(data_store)


def _make_monthly_data(store: dict, label: Union[str, None]=None):
    for yyyymm, data in store.items():
        if isinstance(data, list):
            store[yyyymm] = MonthlyData(yyyymm, list(map(lambda x: ProfitData(ProfitDataItem(label), x), data)))
        elif isinstance(data, dict):
            store[yyyymm] = MonthlyData(yyyymm, list(map(lambda x: ProfitData(ProfitDataItem(x[0]), x[1]), data.items())))
        else:
            store[yyyymm] = MonthlyData(yyyymm, [ProfitData(ProfitDataItem(label), data)])


def _sum_all_rows(data: MonthlyData):
    total = 0
    for d in data.rows:
        if d.value is not None:
            total += d.value
    return total


def _divide_common_expense(data_store: dict):
    memo_rest = {}  # 全社共通のシートの勘定科目の値を各事業に按分した時の残り

    # 按分率が設定されている項目を見つける
    for business, definitions in data_store["definition"].items():
        if business == "全社共通": continue
        for item in definitions["loss"]:
            if item.ratio is None or item.ratio == "": continue
            memo_rest.setdefault(item.account, 1)
            memo_rest[item.account] -= item.ratio
    if len(memo_rest) == 0:
        return

    for account, rest in memo_rest.items():
        if rest > 0.0000001 or rest < -0.0000001:  # 丸目誤差対策
            print(f"XXX 勘定科目[{account}]の全事業の按分率の合計が{(1 - rest):05f}になっています。ちょうど1になるように設定してください")
            sys.exit(1)
        memo_rest[account] = 0  # 丸目誤差対策

    # 按分率を適用して値を書き換える
    for typ in ["plan", "performance"]:
        if typ not in data_store: continue
        origin = dict()
        for yyyymm, monthly_data in data_store[typ]["全社共通"]["loss"].items():
            if "決算" in yyyymm: continue
            for d in monthly_data.rows:
                if d.label is None or d.label.account not in memo_rest.keys() or d.value is None:
                    continue
                # 共通（按分元）に残りを設定する
                #d.rest_value = d.value * memo_rest[d.label.account]
                d.rest_value = 0  # 全部分配する
                # 共通（按分元）の値を後の処理のために保存しておく
                origin.setdefault(yyyymm, {})[d.label.account] = d.value

        for business, data in data_store[typ].items():
            if business == "全社共通": continue
            for yyyymm, monthly_data in data["loss"].items():
                if "決算" in yyyymm: continue
                if yyyymm not in origin: continue
                for d in monthly_data.rows:
                    if d.label is None or d.label.ratio is None or d.label.account not in origin[yyyymm]:
                        continue
                    d.value = origin[yyyymm][d.label.account] * d.label.ratio


def _calculate_all_earnings(data_store: dict):
    """全ての事業について利益を計算する"""
    for typ in ["plan", "performance"]:
        if typ not in data_store: continue
        for business, data in data_store[typ].items():
            data_store[typ][business]["earnings"] = dict()
            for yyyymm, monthly_data in data["profit"].items():
                if "決算" in yyyymm: continue
                for d in monthly_data.rows:
                    if d.label is None or d.value is None:
                        continue
                    data_store[typ][business]["earnings"][yyyymm] = d.value

            for yyyymm, monthly_data in data["loss"].items():
                if "決算" in yyyymm: continue
                for d in monthly_data.rows:
                    if d.label is None or d.value is None:
                        continue
                    data_store[typ][business]["earnings"].setdefault(yyyymm, 0)
                    data_store[typ][business]["earnings"][yyyymm] -= d.value

            # 月ごとのデータはMonthlyDataオブジェクト出なければならないので変換する
            _make_monthly_data(data_store[typ][business]["earnings"], "利益")


def aggregate_fixval(ws_type: str, business: str, header_row: list[str], data_store: dict):
    """変動費・固定費の集計結果を表にする"""
    expense = ["固定費", "変動費"]
    #expense_categories = list(set(map(lambda x: x.category, data_store["definition"][business]["loss"])))

    # 出力データの形を作る
    result = {"loss": {}, "variable_ratio": {}}
    total_sales = {}

    # 売上は事業の売上項目ごと、経費は変動費・固定費で集約する
    for yyyymm in header_row:
        if "決算" in yyyymm:
            continue
        if yyyymm in data_store[ws_type][business]["loss"]:
            for row in data_store[ws_type][business]["loss"][yyyymm].rows:
                if row.value is None: continue
                # 変動費・固定費 (あとで形式をProfitData型に変更する必要がある）
                result["loss"].setdefault(yyyymm, {}).setdefault(row.label.fixval, 0)
                result["loss"][yyyymm][row.label.fixval] += row.value

        if yyyymm in data_store[ws_type][business]["profit"]:
            for m in data_store[ws_type][business]["profit"][yyyymm].rows:
                if m.value is not None:
                    total_sales.setdefault(yyyymm, 0)
                    total_sales[yyyymm] += m.value

    # 変動費・固定費の情報の型をProfitData型に変更。変動比率の計算
    for yyyymm, dat in result["loss"].items():
        for label, value in result["loss"][yyyymm].items():
            if label == "変動費" and yyyymm in total_sales and total_sales[yyyymm] > 0:
                ratio = int(value/total_sales[yyyymm] * 10000)/10000
                result["variable_ratio"][yyyymm] = ratio

    # 月ごとのデータはMonthlyDataオブジェクト出なければならないので変換する
    _make_monthly_data(result["loss"])  # 与えたデータ(result["loss"]の中身が(label, value)のタプルになっている場合は、第２引数を指定しない

    return result, expense


def aggregate_category(ws_type: str, business: str, header_row: list[str], data_store: dict):
    """経費カテゴリ別の集計結果を表にする"""
    expense = list(set(map(lambda x: x.category, data_store["definition"][business]["loss"])))

    # 出力データの形を作る
    result = {"profit": {}, "loss": {}}

    # 売上は、変動費・固定費の表のところで計算済みなので、ここでは経費カテゴリのみ集約する
    for yyyymm in header_row:
        if "決算" in yyyymm:
            continue
        if yyyymm in data_store[ws_type][business]["loss"]:
            for row in data_store[ws_type][business]["loss"][yyyymm].rows:
                if row.value is None: continue
                # 変動費・固定費 (あとで形式をProfitData型に変更する必要がある）
                result["loss"].setdefault(yyyymm, {}).setdefault(row.label.category, 0)
                result["loss"][yyyymm][row.label.category] += row.value

    # 月ごとのデータはMonthlyDataオブジェクト出なければならないので変換する
    _make_monthly_data(result["loss"])  # 与えたデータ(result["loss"]の中身が(label, value)のタプルになっている場合は、第２引数を指定しない

    return result, expense


def aggregate_all_business(data_store: dict, header_row: list[str]):
    """全事業のprofit/lossを結合して一つにまとめる
    経費は経費グループごとにまとめる
    """
    # 経費グループリスト、売上リストを見つける
    sales_list = list()
    expense_group = list()
    for business, conf in data_store["definition"].items():
        sales_list.extend(filter(lambda y: y is not None, map(lambda x: x.name, conf["profit"])))
        expense_group.extend(map(lambda x: x.group, conf["loss"]))
    expense_group = list(set(expense_group))

    # 出力データの形を作る
    result = dict()

    # 売上はそれぞれの事業の売上項目ごと、経費は経費グループで集約する
    for typ in ["plan", "performance"]:
        result.setdefault(typ, {"profit": {}, "loss": {}, "earnings": {}})
        if typ not in data_store: continue
        for yyyymm in header_row:
            if "決算" in yyyymm:
                continue
            for business, dat in data_store[typ].items():
                if yyyymm in dat["profit"]:
                    for row in dat["profit"][yyyymm].rows:
                        if row.value is None: continue
                        result[typ]["profit"].setdefault(yyyymm, {}).setdefault(row.label.name, 0)
                        result[typ]["profit"][yyyymm][row.label.name] += row.value
                        result[typ]["earnings"].setdefault(yyyymm, 0)
                        result[typ]["earnings"][yyyymm] += row.value
                if yyyymm in dat["loss"]:
                    for row in dat["loss"][yyyymm].rows:
                        if row.value is None: continue
                        val = row.value if row.rest_value is None else row.rest_value
                        result[typ]["loss"].setdefault(yyyymm, {}).setdefault(row.label.group, 0)
                        result[typ]["loss"][yyyymm][row.label.group] += val
                        result[typ]["earnings"].setdefault(yyyymm, 0)
                        result[typ]["earnings"][yyyymm] -= val

        _make_monthly_data(result[typ]["profit"])
        _make_monthly_data(result[typ]["loss"])
        _make_monthly_data(result[typ]["earnings"], "利益")

    return result, sales_list, expense_group

