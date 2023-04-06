from openpyxl.styles import Protection, NamedStyle, Font, PatternFill, Color, Alignment
from openpyxl.styles.borders import Border, Side
from openpyxl.styles.numbers import builtin_format_code, builtin_format_id

BASE_FONT = "Arial"
HEADER_FONT = "ＭＳ Ｐゴシック"


def gen_named_style(title: str, font=BASE_FONT, bg_color="ffffff", fg_color="000000", size=14, alignment="right", bold=False):
    return NamedStyle(name=title,
                      font=Font(name=font, size=size, color=fg_color, bold=bold),
                      fill=PatternFill(patternType='solid', fgColor=Color(rgb=bg_color)),
                      alignment=Alignment(horizontal=alignment, vertical='bottom'))


# フォントやセル背景のスタイル
title_style = gen_named_style("title1", size=20, alignment="left", bold=True)
table_main_style = gen_named_style("table_main")
table_main2_style = gen_named_style("table_main2", alignment="left")
table_yellow_style = gen_named_style("table_yellow", bg_color="fcf4dd")
table_blue_style = gen_named_style("table_blue", bg_color="dce5f1")
table_light_gray_style = gen_named_style("table_light_gray", bg_color="dfdfdf")
header_date_style = gen_named_style("header_date", font=HEADER_FONT, bg_color="ebf1de", alignment="left")
column_label_style = gen_named_style("column_label", font=HEADER_FONT, bg_color="ded9c4", alignment="left")
column_label2_style = gen_named_style("column_label2", font=HEADER_FONT, bg_color="ded9c4", alignment="right")
table_aggregated_style = gen_named_style("table_aggregated", font=BASE_FONT, bg_color="dce5f1")
table_aggregated2_style = gen_named_style("table_aggregated2", font=BASE_FONT, bg_color="fcd5b4")


# 罫線
side0 = Side(style='hair', color='000000')
border_hair_box = Border(top=side0, bottom=side0, left=side0, right=side0)

side1 = Side(style='thin', color='000000')
border_box = Border(top=side1, bottom=side1, left=side1, right=side1)

# 数値フォーマット
number_format = builtin_format_code(38)  # マイナスは赤字になる。カンマ区切り
percentage_format = builtin_format_code(10)  # パーセント表示（小数第２位まで）
