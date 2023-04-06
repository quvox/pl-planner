from typing import Tuple, Dict, Union

from . import utils


def get_data(data_list: list, col_value: any):
    for d in data_list:
        if d["label"] == col_value:
            return d["value"]
    return


class SingleTable:
    """シート内の一つの表を表すクラス"""

    def __init__(self, worksheet: any, left_top_pos: Tuple[int, int], row_label_column_num=1):
        self.ws = worksheet
        self.left_top_pos = left_top_pos

        # 年月のヘッダ部分
        self.headers = []     # type: list[str]
        self.header_style = {}  # type: Dict[str, any]  # key = style, border

        # 表の行ラベルの上のヘッダ部分（年月のヘッダの左）
        self.row_label_column_num = row_label_column_num  # 行ラベルの列数（通常は１だと思う）
        self.row_label_headers = []  # type: list[str]
        self.row_labels_style = {}  # type: Dict[str, any]  # key = style, border

        self.sub_tables = {}  # type: Dict[str, SubTable]
        self.table_structure = []  # type: List[Union[str, SubTable]]

    def get_table_cell(self, row: int, column: int):
        """表全体の中の位置を指定して、そのセルを得る。左上のセルをrow=0,column=0とする"""
        return self.ws.cell(row=self.left_top_pos[1]+row, column=self.left_top_pos[0]+column)

    def get_max_row(self):
        """表の一番下の行の行番号を返す"""
        num = 1  # 最初にヘッダがあるので1から始める
        for n in self.table_structure:
            if n == "" or n is None:  # 空行
                num += 1
            else:
                num += self.sub_tables[n].row_size
        return num

    def add_sub_table(self, name: str) -> 'SubTable':
        row_num = self.get_max_row()
        st = SubTable(self.ws, self, self.left_top_pos[1]+row_num, self.left_top_pos[0])
        self.table_structure.append(name)
        self.sub_tables[name] = st
        return st

    def add_single_row(self, name: str, row_header: list, style: Union[Dict[str, any], None] = None, merge=False) -> 'SubTable':
        row_num = self.get_max_row()
        st = SubTable(self.ws, self, self.left_top_pos[1]+row_num, self.left_top_pos[0])
        st.set_row_labels(list(map(lambda x: [(x)], row_header)), style, merge)
        self.table_structure.append(name)
        self.sub_tables[name] = st
        return st

    def add_blank_row(self, merge=True):
        """空行を入れる
        Args:
            merge (bool): Trueなら行全体(headerがあるところまで)のセルを結合する
        """
        if merge:
            self.table_structure.append("")    # 空行(列全体をマージ)を追加
        else:
            self.table_structure.append(None)  # 空行を追加

    def add_black_row_after(self, name: str, merge=True):
        """指定したサブテーブルの後ろに空行を追加する"""
        idx = next((i for i, s in enumerate(self.table_structure) if s == name), None)
        if idx is None:
            self.add_blank_row(merge)
        elif merge:
            self.table_structure.insert(idx+1, "")
        else:
            self.table_structure.insert(idx+1, None)

    def set_headers(self, headers: list[str], style: Union[Dict[str, any], None] = None):
        self.headers = headers
        if style is not None:
            self.header_style = style

    def set_row_label_headers(self, headers: list[str], style: Union[Dict[str, any], None] = None):
        for i in range(len(headers)):
            cell = self.ws.cell(row=self.left_top_pos[1], column=i+1)
            utils.set_style_and_value(cell, headers[i], style)

    def create_frame(self):
        """ヘッダと行ラベルなど、周りの部分を構築する"""
        # ヘッダ行を作る
        # -- 行ラベルの上の部分
        for i in range(self.row_label_column_num):
            cell = self.ws.cell(row=self.left_top_pos[1], column=self.left_top_pos[0]+i)
            if len(self.row_label_headers) >= i+1:
                utils.set_style_and_value(cell, self.row_label_headers[i], self.header_style)
            else:
                utils.set_style_and_value(cell, "", self.header_style)
        # -- 年月の部分
        for i, hdr in enumerate(self.headers):
            cell = self.ws.cell(row=self.left_top_pos[1], column=self.left_top_pos[0]+self.row_label_column_num+i)
            utils.set_style_and_value(cell, hdr, self.header_style)

        # 行ラベル（一番左の列）を作る。サブテーブルを作成する（行ラベルと集計行の設定）
        row_num = self.left_top_pos[1] + 1  # ヘッダの次の行から始める
        for nm in self.table_structure:
            if nm == "":
                utils.merge_cells(self.ws, (self.left_top_pos[0]+0, row_num), (self.left_top_pos[0]+len(self.headers), row_num))
                row_num += 1
            elif nm is None:
                row_num += 1
            else:
                self.sub_tables[nm].create_frame()  # サブテーブルの作成
                row_num += self.sub_tables[nm].row_size

    def put_row_sum(self, row_num: int, sum_start_row: int, sum_end_row: int, style_defs: Union[dict, None]):
        """指定した行に指定した行(sum_start_row)から行(sum_end_row)までのSUM式を記入する

        Args:
            row_num (int): 表のボディ部の一番上を1としたときの行番号
            sum_start_row (int): 集計開始行の行番号（表のボディ部の一番上を1）
            sum_end_row (int): 集計終了行の行番号（表のボディ部の一番上を1）
            style_defs (Union[dict, None]): {"style": ..., "border": ..., "format": ...}
        """
        for i in range(len(self.headers)):
            cell = self.ws.cell(row=self.left_top_pos[1]+row_num, column=self.row_label_column_num+i+1)
            start = utils.get_cell_coordinate(self.ws, row=self.left_top_pos[1]+sum_start_row, column=self.row_label_column_num+i+1)
            end = utils.get_cell_coordinate(self.ws, row=self.left_top_pos[1]+sum_end_row, column=self.row_label_column_num+i+1)
            utils.set_style_and_value(cell, f"=SUM({start}:{end})", style_defs)

    def put_data_in_row(self, row_num: int, data: dict, style_defs: Union[dict, None]):
        """指定した行に指定した行にデータを記入する

        Args:
            row_num (int): 表のボディ部の一番上を1としたときの行番号
            data (dict): データ列
            style_defs (Union[dict, None]): {"style": ..., "border": ..., "format": ...}
        """
        for i in range(len(self.headers)):
            cell = self.ws.cell(row=self.left_top_pos[1]+row_num, column=self.row_label_column_num+i+1)
            if self.headers[i] in data:
                utils.set_style_and_value(cell, data[self.headers[i]], style_defs)
            else:
                utils.set_style_and_value(cell, None, style_defs)

    def put_data_at(self, column_num: int, row_num: int, value: any, style_defs: Union[dict, None]):
        """指定した位置（表のメインボディの左上からの位置）にデータを記入する"""
        cell = self.ws.cell(row=self.left_top_pos[1]+row_num, column=self.row_label_column_num+column_num+1)
        utils.set_style_and_value(cell, value, style_defs)

    def set_row_style(self, row_num: int, style_defs: Union[dict, None]):
        """指定した行にスタイルを設定する

        Args:
            row_num (int): 表のボディ部の一番上を1としたときにの行番号
            style_defs (Union[dict, None]): {"style": ..., "border": ..., "format": ...}
        """
        for i in range(len(self.headers)):
            cell = self.ws.cell(row=self.left_top_pos[1]+row_num, column=self.row_label_column_num+i+1)
            utils.set_style(cell, style_defs)

    def read_as_header(self, row=0):
        """指定された行をヘッダ行として読み込む
        Args:
            row (int): 表の一番上の行を0行目とした時の行番号
        """
        row += self.left_top_pos[1]
        col = self.left_top_pos[0] + self.row_label_column_num

        while True:
            cell = self.ws.cell(row=row, column=col)
            if cell.value is None or cell.value == "":
                break
            self.headers.append(cell.value)
            col += 1


class SubTable:
    """SingleTable内に置くサブテーブルのインスタンス"""
    def __init__(self, worksheet: any, parent: SingleTable, start_row: int, left_pos: int):
        self.ws = worksheet
        self.parent = parent
        self.start_row = start_row
        self.left_pos = left_pos
        self.row_size = 0

        self.row_labels = []  # type: list[list]
        self.row_labels_style = None
        self.merge_row_label_cells = False

        self.label_aggregate_right_column = None  # 文字列を設定すると、最終列に合計の列を追加する
        self.style_aggregate_right_column = {}    # type: Dict[str, any]  # key = style, border
        self.label_aggregate_row = None    # 文字列を設定すると、先頭行または最終行に合計の行を追加する
        self.style_aggregate_row = {}      # type: Dict[str, any]  # key = style, border
        self.aggregate_row_top = True   # Trueならサブテーブルの最下行、Falseなら際上行に合計行を置く

    def set_row_labels(self, labels: list[list], style: Union[Dict[str, any], None] = None, merge=False):
        self.row_labels = labels
        self.row_size = len(labels)
        if self.label_aggregate_row is not None:
            self.row_size += 1
        if style is not None:
            self.row_labels_style = style
        self.merge_row_label_cells = merge

    def setup_dummy_aggregate_row(self, at_top=True):
        """エクセルを読み込むためのインスタンスに与えるダミーの設定"""
        self.aggregate_row_top = at_top
        self.label_aggregate_row = []

    def get_offset(self):
        if self.label_aggregate_row is None:
            return 0, -1
        if self.aggregate_row_top:
            row_offset = 1  # データ行は2行目から
            agg_row = 0     # 集計行は1行目
        else:
            row_offset = 0  # データ行は1行目から
            agg_row = self.row_size  # 集計行は最終行の1つ次の行
        return row_offset, agg_row

    def set_right_aggregation_column(self, label: str, style: Union[Dict[str, any], None] = None):
        self.label_aggregate_right_column = label
        if style is not None:
            self.style_aggregate_right_column = style

    def set_aggregation_row(self, labels: str, style: Union[Dict[str, any], None] = None, at_top=True):
        self.label_aggregate_row = labels
        if style is not None:
            self.style_aggregate_row = style
        self.aggregate_row_top = at_top

        # 表サイズ（行数）の再計算
        self.row_size = len(self.row_labels)
        if self.label_aggregate_row is not None:
            self.row_size += 1

    def create_frame(self):
        """行ラベルと集計行を構築する"""
        row_offset, agg_row = self.get_offset()

        # 行ラベル
        for i, labels in enumerate(self.row_labels):
            for k in range(self.parent.row_label_column_num):
                cell = self.ws.cell(row=self.start_row+row_offset+i, column=self.left_pos+k)
                if len(labels) > k:
                    utils.set_style_and_value(cell, labels[k], self.row_labels_style)
                else:
                    utils.set_style_and_value(cell, None, self.row_labels_style)
            if self.merge_row_label_cells:
                utils.merge_cells(self.ws, (self.left_pos, self.start_row+row_offset+i), (self.left_pos+self.parent.row_label_column_num-1, self.start_row+row_offset+i))

        # 合計列(一番右に配置)
        if self.label_aggregate_right_column is not None:
            for i in range(self.start_row+row_offset, self.start_row+self.row_size):
                cell = self.ws.cell(row=i, column=self.left_pos+self.parent.row_label_column_num+len(self.parent.headers)+1)
                utils.set_style_and_value(cell, '', self.style_aggregate_right_column)  # TODO: 列要素にSUMを入れる

        # 合計行
        if self.label_aggregate_row is not None:
            # 最左のラベルのカラム
            cell = self.ws.cell(row=self.start_row+agg_row, column=self.left_pos)
            utils.set_style_and_value(cell, self.label_aggregate_row, self.style_aggregate_row)

            if self.parent.row_label_column_num > 1:
                for i in range(1, self.parent.row_label_column_num):
                    cell = self.ws.cell(row=self.start_row+agg_row, column=self.left_pos+i)
                    utils.set_style_and_value(cell, None, self.style_aggregate_row)
                utils.merge_cells(self.ws, (self.left_pos, self.start_row+agg_row), (self.left_pos+self.parent.row_label_column_num-1, self.start_row+agg_row))
            # 合計値のセル
            self.put_row_sum(agg_row, row_offset, self.row_size+row_offset-2, self.style_aggregate_row)

    def set_row_style(self, row_num: int, style_defs: Union[dict, None]):
        """指定した行にスタイルを設定する

        Args:
            row_num (int): 表のボディ部の一番上を1としたときにの行番号
            style_defs (Union[dict, None]): {"style": ..., "border": ..., "format": ...}
        """
        self.parent.set_row_style(self.start_row+row_num-1, style_defs)

    def set_col_style(self, col_num: int, style_defs: Union[dict, None]):
        """指定した列にスタイルを設定する

        Args:
            col_num (int): 表のボディ部の一番左を1としたときにの行番号
            style_defs (Union[dict, None]): {"style": ..., "border": ..., "format": ...}
        """
        for i in range(self.start_row, self.start_row+self.row_size):
            cell = self.ws.cell(row=i, column=self.parent.row_label_column_num+col_num)
            utils.set_style(cell, style_defs)

    def put_data_in_column(self, column_num: int, data_list: Union[list, None], style_defs: Union[dict, None]):
        """指定した列にデータ列を記入する
        Args:
            column_num (int): ボディ部最初の列を0とした時に、何番目の列かを表す
            data_list (Union[list, None]): [{"label": (列ラベルのカラムの文字列,,,), "value": 数値}, {...},,,]
            style_defs (Union[dict, None]): {"style": ..., "border": ..., "format": ...}
        """
        row_offset, agg_row = self.get_offset()
        if agg_row > -1:
            # 合計行があるときは、そこをセットする
            self.put_row_sum(agg_row, 1, len(self.row_labels), self.style_aggregate_row)

        for i in range(len(self.row_labels)):
            cell = self.ws.cell(row=self.start_row+row_offset+i, column=self.parent.row_label_column_num+column_num+1)  # +1はヘッダ行の分
            if data_list is not None:
                value = get_data(data_list, self.row_labels[i])
                utils.set_style_and_value(cell, value, style_defs)
            else:
                utils.set_style_and_value(cell, None, style_defs)

    def put_column_sum(self, column_num: int, sum_start_column: int, sum_end_column: int, style_defs: Union[dict, None]):
        """指定した列に指定した列(sum_start_column)から列(sum_end_column)までのSUM式を記入する"""
        row_offset, agg_row = self.get_offset()
        for i in range(len(self.row_labels)):
            cell = self.ws.cell(row=self.start_row+row_offset+i, column=self.parent.row_label_column_num+column_num+1)  # +1はヘッダ行の分
            start = utils.get_cell_coordinate(self.ws, row=self.start_row+row_offset+i, column=self.parent.row_label_column_num+sum_start_column+1)
            end = utils.get_cell_coordinate(self.ws, row=self.start_row+row_offset+i, column=self.parent.row_label_column_num+sum_end_column+1)
            utils.set_style_and_value(cell, f"=SUM({start}:{end})", style_defs)

    def put_data_in_row(self, row_num: int, data_list: list, style_defs: Union[dict, None]):
        """指定した行にデータ列を記入する"""
        for i in range(len(self.parent.headers)):
            cell = self.ws.cell(row=self.start_row+row_num, column=self.parent.row_label_column_num+i+1)
            if data_list is not None:
                value = get_data(data_list, self.row_labels[i])
                utils.set_style_and_value(cell, value, style_defs)
            else:
                utils.set_style_and_value(cell, None, style_defs)

    def put_row_sum(self, row_num: int, sum_start_row: int, sum_end_row: int, style_defs: Union[dict, None]):
        """指定したサブテーブル内の行に指定した行(sum_start_row)から行(sum_end_row)までのSUM式を記入する（集計行の作成）

        Args:
            row_num (int): サブテーブルの一番上を0としたときの集計行の行番号
            sum_start_row (int): 集計開始行の行番号（サブテーブルの一番上を0）
            sum_end_row (int): 集計終了行の行番号（サブテーブルの一番上を0）
            style_defs (Union[dict, None]): {"style": ..., "border": ..., "format": ...}
        """
        subtable_top = self.start_row-2
        self.parent.put_row_sum(row_num+subtable_top, sum_start_row+subtable_top, sum_end_row+subtable_top, style_defs)

    def put_data_at(self, column_num: int, row_num: int, value: any, style_defs: Union[dict, None]):
        """指定した位置（サブテーブルの左上からの位置）にデータを記入する"""
        cell = self.ws.cell(row=self.start_row+row_num, column=self.parent.row_label_column_num+column_num+1)
        utils.set_style_and_value(cell, value, style_defs)

    def read_as_row_labels(self, row_size: int):
        """指定された行数を一つのサブテーブルだとみなして、行ラベルを読む

        Args:
            row_size (int): サブテーブルの行数（集計行も含む）
        """
        self.row_size = row_size
        row_offset, agg_row = self.get_offset()
        if agg_row > -1:
            row_size -= 1   # 集計行のラベルは取得しなくて良いので１減らす
        for i in range(row_size):
            result = list()
            for k in range(self.parent.row_label_column_num):
                cell = self.ws.cell(row=self.start_row+i+row_offset, column=k+1)
                result.append(cell.value)
            self.row_labels.append(result)

    def get_all_data(self):
        """サブテーブルのボディ部のデータを読み込む"""
        row_offset, agg_row = self.get_offset()
        result = dict()
        for k, yyyymm in enumerate(self.parent.headers):
            result[yyyymm] = []
            for i, labels in enumerate(self.row_labels):
                cell = self.ws.cell(row=self.start_row+i+row_offset, column=k+self.parent.row_label_column_num+self.left_pos)
                result[yyyymm].append({"label": labels, "value": cell.value})
        return result
