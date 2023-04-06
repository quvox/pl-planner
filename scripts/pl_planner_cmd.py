from typing import Union
import os
import sys
import json
import glob
import datetime
from argparse import ArgumentParser

sys.path.append("./libs")
from libs import config, data, build_table, pldata, common


def _parser():
    usage = 'python {} [-n num] [--help]'.format(os.path.basename(__file__))
    argparser = ArgumentParser(usage=usage)
    argparser.add_argument('-d', '--directory', type=str, default="../data", help='directory where excel files are located')
    argparser.add_argument('-c', '--create', action="store_true", default=False, help='create/update profit/loss excel files')
    argparser.add_argument('-s', '--start', type=str, help='start month (YYYYMM)')
    argparser.add_argument('-e', '--end', type=str, help='end month (YYYYMM)')
    return argparser.parse_args()


def calc_period(start: str, end: str, data_store: dict):
    """与えられた開始月、終了月を含む期の期初と期末までのdatetimeを返す"""
    now = datetime.datetime.today()
    settlement_month = data_store["config"].get("決算月", 3)
    if start is None:
        # startの指定がない場合は直近2期分
        start_dt = common.get_term_start_month(now, settlement_month)
    else:
        dt = common.convert_from_yyyymm(start)
        start_dt = common.get_term_start_month(dt, settlement_month)
    if end is None:
        # endの指定がない場合は今の期の期末
        end_dt = common.get_term_end_month(now, settlement_month, 12)
    else:
        dt = common.convert_from_yyyymm(end)
        end_dt = common.get_term_end_month(dt, settlement_month)

    return start_dt, end_dt


def read_data_store(directory: str) -> tuple[dict, pldata.LabelManager]:
    mgr = pldata.LabelManager()
    file_path = os.path.join(directory, "store.json")
    if not os.path.exists(file_path):
        return {}, mgr

    # JSONデータ内の表定義の情報をクラスオブジェクトに変更する
    with open(file_path) as f:
        data_store = json.load(f)

    for business, conf in data_store["definition"].items():
        for i in range(len(conf["profit"])):
            conf["profit"][i] = pldata.ProfitDataItem(**conf["profit"][i])
            mgr.add(business, "profit", conf["profit"][i])
        for i in range(len(conf["loss"])):
            conf["loss"][i] = pldata.LossDataItem(**conf["loss"][i])
            mgr.add(business, "loss", conf["loss"][i])

    # JSONデータ内の計画情報/実績情報を月毎のMonthlyDataオブジェクトに変更する
    for typ in ["plan", "performance"]:
        if typ not in data_store: continue
        for business, data in data_store[typ].items():
            for yyyymm, monthly_data in data["profit"].items():
                monthly_data_list = list(map(lambda x: pldata.ProfitData(mgr.get(business, "profit", name=x["label"][0]), x["value"]), monthly_data))
                data["profit"][yyyymm] = pldata.MonthlyData(yyyymm, monthly_data_list)
            for yyyymm, monthly_data in data["loss"].items():
                monthly_data_list = list(map(lambda x: _make_loss_data(business, mgr, x), monthly_data))
                data["loss"][yyyymm] = pldata.MonthlyData(yyyymm, monthly_data_list)
    return data_store, mgr


def _make_loss_data(business: str, mgr: pldata.LabelManager, x: dict) -> pldata.LossData:
    d = pldata.LossData(mgr.get(business, "loss", group=x["label"][0], account=x["label"][1], category=x["label"][2]), x["value"])
    if "rest_value" in x and x["rest_value"] is not None:
        d.rest_value = x["rest_value"]
    return d


def get_file_paths(directory: str, data_store: dict) -> Union[str, list[str]]:
    """事業別ファイルと全社共通ファイルを読み込む"""
    result = list()
    for business in data_store["definition"].keys():
        fp = os.path.join(directory, f"{business}.xlsx")
        if os.path.exists(fp):
            result.append(fp)
    return result


def convert_proc(obj):
    if hasattr(obj, "obj"):  # TODO: class名にパスが入ってしまい正しく判定できない
        return obj.obj()
    elif hasattr(obj, "list_monthly_data"):
        return obj.list_monthly_data()
    raise TypeError


if __name__ == '__main__':
    args = _parser()
    start_dt = None
    end_dt = None

    if not os.path.exists(args.directory):
        print("XXX no such directory:", args.directory)
        sys.exit(-1)
    print("*** データディレクトリ：", args.directory)

    # データストアファイル（過去の入力情報）を読み込む
    store, label_mgr = read_data_store(args.directory)

    # 設定ファイルを読み込む
    config.read_config_file(os.path.join(args.directory, "設定.xlsx"), store, label_mgr)

    # 事業別ファイル、全社共通ファイルを読み込む
    files = get_file_paths(args.directory, store)
    for fp in files:
        start_dt, end_dt = data.read_data_file(fp, store, label_mgr)  # 戻り値はエクセルに含まれているデータの期間

    # 集計期間を期初からにする。引数で与えられていたら、そちらの設定を優先する
    if args.start is not None or args.end is not None or start_dt is None or end_dt is None:
        start_dt, end_dt = calc_period(args.start, args.end, store)
        print(">>>>>>>>", start_dt, end_dt)

    # 共通シートに記載された経費を按分して各事業に振り分ける
    data.update(store)

    # データをJSONで保存する（過去の分も結合して保存する）
    with open(os.path.join(args.directory, "store.json"), "w") as f:
        json.dump(store, f, default=pldata.convert_proc)

    # 全社共通、事業別ファイルを生成または更新する
    # データストアファイル（jsonファイル）があり、入力済みデータがあるならそれもprofit,lossファイルに書き込む
    build_table.build_business_books(args.directory, store, start_dt, end_dt)

    # 集計して一つの情報に統合し、全社統合版PL表エクセルを書き出す
    build_table.create_pl_book(args.directory, store, start_dt, end_dt)
