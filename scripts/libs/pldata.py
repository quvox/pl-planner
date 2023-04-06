from typing import Union


class ProfitDataItem:
    def __init__(self, name: str, memo=""):
        self.name = name
        self.memo = memo

    def __eq__(self, other):
        return self.name == other.name

    def __hash__(self):
        return hash((self.name,))

    def obj(self):
        return {"name": self.name, "memo": self.memo}

    def tuple(self) -> tuple:
        return tuple([self.name])


class ProfitData:
    def __init__(self, label: ProfitDataItem, value: any):
        self.label = label
        self.value = value


class LossDataItem:
    def __init__(self, group: str, account: str, category: str, fixval: str, ratio: str, memo=""):
        self.group = group
        self.account = account
        self.category = category
        self.fixval = fixval
        self.ratio = ratio
        self.memo = memo

    def __eq__(self, other):
        return self.group == other.group and self.account == other.account and self.category == other.category

    def __hash__(self):
        return hash((self.group, self.account, self.category))

    def obj(self):
        return {"group": self.group, "account": self.account, "category": self.category, "fixval": self.fixval, "ratio": self.ratio, "memo": self.memo}

    def tuple(self) -> tuple:
        return tuple([self.group, self.account, self.category])


class LossData:
    def __init__(self, label: LossDataItem, value: any):
        self.label = label
        self.value = value
        self.rest_value = None  # 全社共通のシートでだけ利用する。valueの値を他事業に按分した時の残りの分


class MonthlyData:
    def __init__(self, yyyymm: str, rows: list[Union[LossData, ProfitData]]):
        self.yyyymm = yyyymm
        self.rows = rows

    def list_monthly_data(self):
        if self.rows is None or len(self.rows) == 0:
            return []
        rows = list(filter(lambda x: x.label is not None, self.rows))
        if hasattr(rows[0], "rest_value"):
            return list(map(lambda x: {"label": x.label.tuple(), "value": x.value, "rest_value": x.rest_value}, rows))
        else:
            return list(map(lambda x: {"label": x.label.tuple() if x.label is not None else "", "value": x.value}, rows))

    def find_account(self, account: str) -> Union[LossData, ProfitData]:
        return next((r for r in self.rows if r.label.account == account), None)

    def merge(self, new_rows: list[Union[LossData, ProfitData]]):
        """重複を排除しながらマージする"""
        for r in new_rows:
            if r.label is None:
                continue
            idx = next((i for i, d in enumerate(self.rows) if d.label is not None and r.label.tuple() == d.label.tuple()), None)
            if idx is None:
                self.rows.append(r)
            else:
                self.rows[idx] = r


class LabelManager:
    """ProfitDataItemやLossDataItemのリストを管理し、Item解決の問い合わせに答える"""
    def __init__(self):
        self.items = dict()

    def add(self, business: str, typ: str, item: Union[ProfitDataItem, LossDataItem]):
        """行ラベルを登録する
        Args:
            business (str): 事業名(=エクセルファイル名)
            typ (str): profit/loss
            item (Union[ProfitDataItem, LossDataItem])): エクセルから読み取ったラベル情報
        """
        if typ == "profit":
            if item.name is None: return
        elif typ == "loss":
            if item.account is None: return
        self.items.setdefault(business, {}).setdefault(typ, set()).add(item)

    def get_all(self, business, typ: str):
        return list(self.items[business][typ])

    def get(self, business: str, typ: str, **kwargs):
        if typ not in self.items[business]:
            return
        for item in self.items[business][typ]:
            flag = False
            for k, v in kwargs.items():
                if getattr(item, k) == v:
                    flag = True
                else:
                    flag = False
                    break
            if flag:
                return item
        return


def convert_proc(obj):
    if hasattr(obj, "obj"):  # TODO: class名にパスが入ってしまい正しく判定できない
        return obj.obj()
    elif hasattr(obj, "list_monthly_data"):
        return obj.list_monthly_data()
    print(obj)
    raise TypeError
