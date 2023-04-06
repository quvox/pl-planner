from typing import Union, Dict
import os
import datetime

from . import common, data
from pldata import ProfitData, LossData, ProfitDataItem, LossDataItem, LabelManager, MonthlyData
from excel import utils, styles, table


def build_business_books(directory: str, data_store: Union[dict, None], start_dt: datetime.datetime, end_dt: datetime.datetime):
    """事業別ファイルを作成する
    保存済みのデータが存在するならそのデータで埋め、なければ空白にしてスタイルだけを設定する
    """
    settlement_month = data_store["config"].get("決算月", 3)  # type: int
    header_row = common.create_header_labels(start_dt, end_dt, settlement_month)

    workbooks = dict()
    for typ in ["plan", "performance"]:
        for business, data_def in data_store["definition"].items():
            data_store.setdefault(typ, {}).setdefault(business, {}).setdefault("profit", {})
            data_store.setdefault(typ, {}).setdefault(business, {}).setdefault("loss", {})
            data_store.setdefault(typ, {}).setdefault(business, {}).setdefault("earnings", {})

            # ワークブック、ワークシートの作成
            if business not in workbooks:
                workbooks[business] = utils.create_new_workbook()
            workbooks[business].create_sheet(title=common.MAPPING1[typ])
            ws = workbooks[business][common.MAPPING1[typ]]

            tbl = create_main_table(ws, typ, business, header_row, data_store)
            create_fixval_table(tbl, typ, business, header_row, data_store)

            # ヘッダ、ラベル部分を出力する
            tbl.create_frame()

            # シート全体に渡って幅を自動調整する
            utils.auto_adjust_column_width(ws)

            # シートの行と列の表示を固定する
            ws.freeze_panes = "D3"

            if "Sheet" in workbooks[business]:
                workbooks[business].remove(workbooks[business]["Sheet"])  # 最初から存在するシートは不要なので削除する
            workbooks[business].save(os.path.join(directory, f"{business}.xlsx"))


def create_pl_book(directory: str, data_store: dict, start_dt: datetime.datetime, end_dt: datetime.datetime):
    """全社統合版のP/L表を作る"""
    settlement_month = data_store["config"].get("決算月", 3)  # type: int
    header_row = common.create_header_labels(start_dt, end_dt, settlement_month)

    result, sales_list, expense_list = data.aggregate_all_business(data_store, header_row)

    # ワークブック、ワークシートの作成
    wb = utils.create_new_workbook()
    for typ in ["plan", "performance"]:
        wb.create_sheet(title=common.MAPPING1[typ])
        ws = wb[common.MAPPING1[typ]]

        # テーブルを作成する
        create_aggregated_pl_tables(ws, result[typ], header_row, sales_list, expense_list)

        # シート全体に渡って幅を自動調整する
        utils.auto_adjust_column_width(ws)

        # シートの行と列の表示を固定する
        ws.freeze_panes = "D3"

    wb.remove(wb["Sheet"])  # 最初から存在するシートは不要なので削除する
    wb.save(os.path.join(directory, "事業計画.xlsx"))


def create_main_table(ws: any, ws_type: str, business: str, header_row: list[str], data_store: dict) -> table.SingleTable:
    data_def = data_store["definition"][business]

    # 表タイトル（事業名）
    utils.set_style_and_value(ws.cell(row=1, column=1),
                              business,
                              {"style": styles.title_style, "border": styles.border_hair_box})

    tbl = table.SingleTable(ws, (1, 2), 3)  # テーブルの起点(左上がA2のセル)、3セルを行ラベル用に使う
    # -- 年月のヘッダ行を設定する
    tbl.set_headers(header_row, {"style": styles.header_date_style, "border": styles.border_box})

    # 売上サブテーブルを作成する
    sales_tbl = tbl.add_sub_table("sales")
    row_labels = create_item_labels(data_def["profit"])
    sales_tbl.set_row_labels(row_labels, {"style": styles.column_label2_style, "border": styles.border_box}, True)
    sales_tbl.set_aggregation_row("売上総計", {"style": styles.table_aggregated_style, "border": styles.border_box, "format": styles.number_format})
    tbl.add_blank_row()

    # 経費サブテーブルを作成する
    expense_tbl = tbl.add_sub_table("expense")
    row_labels = create_item_labels(data_def["loss"])
    expense_tbl.set_row_labels(row_labels, {"style": styles.column_label_style, "border": styles.border_box})
    expense_tbl.set_aggregation_row("経費総計", {"style": styles.table_aggregated_style, "border": styles.border_box, "format": styles.number_format})
    tbl.add_blank_row()

    # 利益サブテーブルを作成する
    earnings_tbl = tbl.add_sub_table("earnings")
    earnings_tbl.set_row_labels([("利益",)], {"style": styles.table_aggregated2_style, "border": styles.border_box}, True)  # 最後の引数をTrueにすると、セル結合する

    # 表の中にデータを入れる、またはデータがないならスタイルだけ設定する
    _create_table_body(sales_tbl, data_store[ws_type][business]["profit"])
    _create_table_body(expense_tbl, data_store[ws_type][business]["loss"])
    _create_table_body(earnings_tbl, data_store[ws_type][business]["earnings"], {"style": styles.table_aggregated2_style, "border": styles.border_box, "format": styles.number_format})

    return tbl


def create_fixval_table(tbl: table.SingleTable, ws_type: str, business: str, header_row: list[str], data_store: dict):
    """変動費・固定費の集計結果を表にする"""
    data_def = data_store["definition"][business]
    ws = tbl.ws

    tbl.add_blank_row()
    tbl.add_blank_row()

    # タイトル
    utils.set_style_and_value(ws.cell(row=tbl.get_max_row()+1, column=1),
                              "変動費・固定費/カテゴリ別分析",
                              {"style": styles.table_main2_style, "border": styles.border_hair_box})

    fixval_result, expense_list = data.aggregate_fixval(ws_type, business, header_row, data_store)
    category_result, expense_category_list = data.aggregate_category(ws_type, business, header_row, data_store)

    # 売上のサブテーブル
    tbl1 = tbl.add_sub_table("fixval_sales")
    row_labels = create_item_labels(data_def["profit"])
    tbl1.set_row_labels(row_labels, {"style": styles.column_label2_style, "border": styles.border_box}, True)
    tbl1.set_aggregation_row("売上総計", {"style": styles.table_aggregated_style, "border": styles.border_box, "format": styles.number_format})
    tbl.add_blank_row()

    # 経費(変動費・固定費別)のサブテーブル
    tbl2 = tbl.add_sub_table("fixval_expense")
    row_labels = create_item_label_list(expense_list)  # tupleのリストにしないとlabelの一致判定で失敗して情報が表示されない
    tbl2.set_row_labels(row_labels, {"style": styles.column_label2_style, "border": styles.border_box}, True)
    tbl2.set_aggregation_row("経費総計", {"style": styles.table_aggregated_style, "border": styles.border_box, "format": styles.number_format})
    tbl.add_single_row("variable_ratio", ["変動比率"], {"style": styles.column_label2_style, "border": styles.border_box}, True)
    variable_ratio_row_num = tbl.get_max_row()-1
    tbl.add_blank_row()

    # 経費(カテゴリ別)のサブテーブル
    tbl3 = tbl.add_sub_table("category_expense")
    row_labels = create_item_label_list(expense_category_list)  # tupleのリストにしないとlabelの一致判定で失敗して情報が表示されない
    tbl3.set_row_labels(row_labels, {"style": styles.column_label2_style, "border": styles.border_box}, True)
    tbl3.set_aggregation_row("経費総計", {"style": styles.table_aggregated_style, "border": styles.border_box, "format": styles.number_format})
    tbl.add_blank_row()

    # 表の中にデータを入れる、またはデータがないならスタイルだけ設定する
    _create_table_body(tbl1, data_store[ws_type][business]["profit"])
    _create_table_body(tbl2, fixval_result["loss"])
    tbl.put_data_in_row(variable_ratio_row_num, fixval_result["variable_ratio"], {"style": styles.table_main_style, "border": styles.border_box, "format": styles.percentage_format})
    _create_table_body(tbl3, category_result["loss"])


def create_aggregated_pl_tables(ws: any, result: dict, header_row: list[str], sales_label_list: list, expense_label_list: list) -> table.SingleTable:
    # 表タイトル（事業名）
    utils.set_style_and_value(ws.cell(row=1, column=1),
                              "事業計画",
                              {"style": styles.title_style, "border": styles.border_hair_box})

    tbl = table.SingleTable(ws, (1, 2), 1)  # テーブルの起点(左上がA2のセル)、3セルを行ラベル用に使う
    # -- 年月のヘッダ行を設定する
    tbl.set_headers(header_row, {"style": styles.header_date_style, "border": styles.border_box})

    # 売上サブテーブルを作成する
    sales_tbl = tbl.add_sub_table("sales")
    row_labels = create_item_label_list(sales_label_list)  # tupleのリストにしないとlabelの一致判定で失敗して情報が表示されない
    sales_tbl.set_row_labels(row_labels, {"style": styles.column_label_style, "border": styles.border_box}, True)
    sales_tbl.set_aggregation_row("売上総計", {"style": styles.table_aggregated_style, "border": styles.border_box, "format": styles.number_format})
    tbl.add_blank_row()

    # 経費サブテーブルを作成する
    expense_tbl = tbl.add_sub_table("expense")
    row_labels = create_item_label_list(expense_label_list)  # tupleのリストにしないとlabelの一致判定で失敗して情報が表示されない
    expense_tbl.set_row_labels(row_labels, {"style": styles.column_label_style, "border": styles.border_box})
    expense_tbl.set_aggregation_row("経費総計", {"style": styles.table_aggregated_style, "border": styles.border_box, "format": styles.number_format})
    tbl.add_blank_row()

    # 利益サブテーブルを作成する
    earnings_tbl = tbl.add_sub_table("earnings")
    earnings_tbl.set_row_labels([("利益",)], {"style": styles.table_aggregated2_style, "border": styles.border_box}, True)  # 最後の引数をTrueにすると、セル結合する

    # ヘッダ、ラベル部分を出力する
    tbl.create_frame()

    # 表の中にデータを入れる、またはデータがないならスタイルだけ設定する
    _create_table_body(sales_tbl, result["profit"])
    _create_table_body(expense_tbl, result["loss"])
    _create_table_body(earnings_tbl, result["earnings"], {"style": styles.table_aggregated2_style, "border": styles.border_box, "format": styles.number_format})

    return tbl


def _create_table_body(tbl: table.SubTable, data: Dict[str, MonthlyData], style_main=None):
    """中身の数字の部分を埋める、またはデータがなければスタイルだけを設定する"""
    fiscal_years = 0
    style_def_aggregation = {"style": styles.table_yellow_style, "border": styles.border_box, "format": styles.number_format}
    if style_main is None:
        style_def_main = {"style": styles.table_main_style, "border": styles.border_box, "format": styles.number_format}
    else:
        style_def_main = style_main

    for i, yyyymm in enumerate(tbl.parent.headers):
        if "決算" in yyyymm:
            # 集計列を入れる(その会計年度のデータのSUMの式を入れる）
            if i + fiscal_years < 12:
                tbl.put_column_sum(i, 0, i - 1, style_def_aggregation)
            else:
                tbl.put_column_sum(i, i - 12, i - 1, style_def_aggregation)
            fiscal_years += 1
        else:
            # データがあればデータを入れる（data[yyyymm]がNoneならスタイルだけ設定する）
            values = data.get(yyyymm)
            if values is not None:
                values = list(map(lambda x: {"value": x.value, "label": x.label.tuple()}, filter(lambda x: x.label is not None, values.rows)))
            tbl.put_data_in_column(i, values, style_def_main)


def _fill_color_for_divided_entries(data_store: dict, tables: list[table.SingleTable]):
    """共通に計上された費用を、他事業に按分する行に色を塗る"""
    # 按分される勘定科目を見つける(同一勘定科目で按分されることに注意)
    divided = set()
    for definitions in data_store["definition"].values():  # keys=business
        for d in definitions["loss"]:
            if d.ratio is not None and d.ratio != "":
                divided.add(d.account)

    style_def = {"style": styles.table_light_gray_style, "border": styles.border_box, "format": styles.number_format}
    # 共通シートの該当行を灰色にする
    for tbl in tables:
        for i in range(1, tbl.get_max_row()):
            account = tbl.get_table_cell(i, 1).value
            if account not in divided:
                continue
            tbl.set_row_style(i, style_def)


def create_item_labels(labels: list[Union[ProfitDataItem, LossDataItem]]) -> list:
    """ProfitDataIteまたはLossDataItemのリストを、row_labelタプルのリストに変換する"""
    result = list()
    for i, label in enumerate(labels):
        result.append(label.tuple())
    return result


def create_item_label_list(label_list: list) -> list:
    """文字列のリストをrow_labelタプルのリストに変換する"""
    # list(map(lambda x: tuple([x]), expense_category_list))
    result = list()
    for i, label in enumerate(label_list):
        if isinstance(label, list):
            result.append(tuple(label))
        else:
            result.append((label,))
    return result

