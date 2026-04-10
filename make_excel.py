import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ===== 共通スタイル =====
YELLOW  = PatternFill("solid", fgColor="FFFF00")
ORANGE  = PatternFill("solid", fgColor="FFC000")
LBLUE   = PatternFill("solid", fgColor="DDEEFF")
LGREEN  = PatternFill("solid", fgColor="E2EFDA")
LGRAY   = PatternFill("solid", fgColor="F2F2F2")
WHITE   = PatternFill("solid", fgColor="FFFFFF")
HEADER  = PatternFill("solid", fgColor="4472C4")
THEADER = PatternFill("solid", fgColor="70AD47")

def bold(size=10, color="000000", white=False):
    c = "FFFFFF" if white else color
    return Font(bold=True, size=size, color=c)

def normal(size=10, color="000000"):
    return Font(size=size, color=color)

def center(wrap=False):
    return Alignment(horizontal="center", vertical="center", wrap_text=wrap)

def right():
    return Alignment(horizontal="right", vertical="center")

def left(wrap=False):
    return Alignment(horizontal="left", vertical="center", wrap_text=wrap)

thin = Side(style="thin")
med  = Side(style="medium")

def all_border(t=thin, b=thin, l=thin, r=thin):
    return Border(top=t, bottom=b, left=l, right=r)

def set_cell(ws, row, col, value, fill=None, font=None, align=None, fmt=None, border=None):
    c = ws.cell(row=row, column=col, value=value)
    if fill:   c.fill = fill
    if font:   c.font = font
    if align:  c.alignment = align
    if fmt:    c.number_format = fmt
    if border: c.border = border
    return c

NUM = '#,##0'
NUM0 = '0.00'

# =====================================================================
# 1. 売上計画検討表
# =====================================================================
def make_uriage(wb):
    ws = wb.create_sheet("売上計画検討表")
    ws.sheet_view.showGridLines = False

    # タイトル
    ws.merge_cells("A1:P1")
    set_cell(ws,1,1,"■ 売上計画検討表", font=bold(14), align=left())

    # 入力エリア
    ws.merge_cells("A2:B2")
    set_cell(ws,2,1,"【入力エリア】※黄色セルを入力", font=bold(10), align=left())
    set_cell(ws,3,1,"席数",      font=bold(), align=center(), fill=LGRAY, border=all_border())
    set_cell(ws,3,2, 15,         font=bold(), align=center(), fill=YELLOW, border=all_border(), fmt=NUM)
    set_cell(ws,3,3,"平均客単価",font=bold(), align=center(), fill=LGRAY, border=all_border())
    set_cell(ws,3,4, 3750,       font=bold(), align=center(), fill=YELLOW, border=all_border(), fmt=NUM)
    set_cell(ws,3,5,"(円/名)",   font=normal(9), align=left())

    ws.merge_cells("A4:B4")
    set_cell(ws,4,1,"※売上高 ＝ 席数 × 客単価 × 回転数 × 日数", font=normal(9,"666666"), align=left())

    # 月別ヘッダー
    months = list(range(1,13))
    day_types = ["平日","金曜日","土曜日","日祝"]
    col_headers = ["回転数","日数","売上高"]

    # 曜日別デフォルト日数（各月）
    # 月: [平日, 金, 土, 日祝]
    default_days = {
        1:[17,4,4,6], 2:[15,4,4,5], 3:[20,5,5,1],
        4:[20,4,4,2], 5:[18,4,4,5], 6:[21,4,4,1],
        7:[22,4,4,1], 8:[21,4,4,2], 9:[21,4,4,1],
        10:[22,5,4,0],11:[20,4,4,2],12:[20,4,4,3],
    }
    # 回転数デフォルト（曜日別）
    default_rot = {"平日":0.62, "金曜日":0.91, "土曜日":0.71, "日祝":0.53}

    row = 6
    # 月ごとに縦並び
    for m in months:
        # 月ヘッダー
        ws.merge_cells(f"A{row}:C{row}")
        set_cell(ws,row,1,f"{m}月", font=bold(11, white=True), align=center(), fill=HEADER, border=all_border())
        set_cell(ws,row,4,"回転数", font=bold(9, white=True), align=center(), fill=HEADER, border=all_border())
        set_cell(ws,row,5,"日数",   font=bold(9, white=True), align=center(), fill=HEADER, border=all_border())
        set_cell(ws,row,6,"売上高", font=bold(9, white=True), align=center(), fill=HEADER, border=all_border())
        row += 1

        total_days_row = []
        total_sales_row = []
        for i, dt in enumerate(day_types):
            d = default_days[m][i]
            r = default_rot[dt]
            set_cell(ws,row,1, dt, font=normal(), align=center(), fill=LGRAY, border=all_border())
            ws.merge_cells(f"B{row}:C{row}")
            set_cell(ws,row,2,"", border=all_border())
            rot_cell = set_cell(ws,row,4, r,  font=normal(), align=center(), fill=YELLOW, border=all_border(), fmt=NUM0)
            day_cell = set_cell(ws,row,5, d,  font=normal(), align=center(), fill=YELLOW, border=all_border(), fmt=NUM)
            # 売上高 = $B$3 * $D$3 * 回転数 * 日数
            sales_formula = f"=$B$3*$D$3*D{row}*E{row}"
            set_cell(ws,row,6, sales_formula, font=normal(), align=right(), fill=LGREEN, border=all_border(), fmt=NUM)
            total_days_row.append(f"E{row}")
            total_sales_row.append(f"F{row}")
            row += 1

        # 合計行
        set_cell(ws,row,1,"合計", font=bold(), align=center(), fill=ORANGE, border=all_border())
        ws.merge_cells(f"B{row}:C{row}")
        set_cell(ws,row,2,"", fill=ORANGE, border=all_border())
        set_cell(ws,row,4,"", fill=ORANGE, border=all_border())
        set_cell(ws,row,5,f"=SUM({','.join(total_days_row)})", font=bold(), align=center(), fill=ORANGE, border=all_border(), fmt=NUM)
        set_cell(ws,row,6,f"=SUM({','.join(total_sales_row)})", font=bold(), align=right(), fill=ORANGE, border=all_border(), fmt=NUM)
        row += 2

    # 列幅
    ws.column_dimensions["A"].width = 10
    ws.column_dimensions["B"].width = 6
    ws.column_dimensions["C"].width = 6
    ws.column_dimensions["D"].width = 8
    ws.column_dimensions["E"].width = 6
    ws.column_dimensions["F"].width = 12

# =====================================================================
# 2. 人件費検討
# =====================================================================
def make_labor(wb):
    ws = wb.create_sheet("人件費検討")
    ws.sheet_view.showGridLines = False

    ws.merge_cells("A1:N1")
    set_cell(ws,1,1,"■ 人件費検討", font=bold(14), align=left())

    # --- シフト表 ---
    ws.merge_cells("A2:N2")
    set_cell(ws,2,1,"【シフト検討表】　営業時間：17時〜29時（12時間）　深夜割増：22時〜 ×1.25", font=bold(10), align=left())

    hours = list(range(17,30))
    headers = ["スタッフ","時給","金額/日"] + [f"{h}時" for h in hours] + ["合計(h)"]

    for ci, h in enumerate(headers, 1):
        fill = HEADER if ci <= 3 or ci == len(headers) else LGRAY
        fc   = bold(9, white=True) if ci <= 3 or ci == len(headers) else bold(9)
        set_cell(ws,3,ci,h, font=fc, align=center(), fill=fill, border=all_border())

    # 繁忙期
    ws.merge_cells("A4:N4")
    set_cell(ws,4,1,"▼ 繁忙期（金・土）", font=bold(10), align=left(), fill=PatternFill("solid",fgColor="FFE0B2"))

    # シフト: 前半(17-23) = Owner+1, 後半(23-29) = 2名
    # スタッフA=Owner(前半), B(前半), C(後半), D(後半)
    staff_busy = [
        ("A（オーナー）", "―",    "―",     [1]*6 + [0]*6),   # 17-23
        ("B（バイト）",   1200, None,      [1,1,1,1,1,1,0,0,0,0,0,0]),  # 17-23: 5h通常+1h深夜
        ("C（バイト）",   1200, None,      [0,0,0,0,0,0,1,1,1,1,1,1]),  # 23-29: 6h深夜
        ("D（バイト）",   1200, None,      [0,0,0,0,0,0,1,1,1,1,1,1]),  # 23-29: 6h深夜
    ]
    # 金額計算: B: 5*1200+1*1500=7500, C/D: 6*1500=9000
    busy_amounts = {"B（バイト）":7500, "C（バイト）":9000, "D（バイト）":9000}

    row = 5
    total_busy = 0
    for name, wage, amt, shifts in staff_busy:
        a = busy_amounts.get(name, "―")
        if isinstance(a, int): total_busy += a
        set_cell(ws,row,1,name, font=normal(), align=center(), fill=LGRAY, border=all_border())
        set_cell(ws,row,2,wage, font=normal(), align=center(), fill=LGRAY, border=all_border())
        set_cell(ws,row,3,a,    font=normal(), align=right(),  fill=LGREEN,border=all_border(), fmt=NUM if isinstance(a,int) else None)
        for ci, s in enumerate(shifts, 4):
            fill = PatternFill("solid",fgColor="4472C4") if s else WHITE
            set_cell(ws,row,ci,"●" if s else "", font=Font(color="FFFFFF" if s else "000000",size=9), align=center(), fill=fill, border=all_border())
        set_cell(ws,row,16,sum(shifts), font=bold(), align=center(), fill=LGRAY, border=all_border())
        row += 1

    # 繁忙期合計
    set_cell(ws,row,1,"計", font=bold(), align=center(), fill=ORANGE, border=all_border())
    set_cell(ws,row,2,"",   fill=ORANGE, border=all_border())
    set_cell(ws,row,3,total_busy, font=bold(), align=right(), fill=ORANGE, border=all_border(), fmt=NUM)
    for ci in range(4,17):
        set_cell(ws,row,ci,"", fill=ORANGE, border=all_border())
    row += 2

    # 通常期
    ws.merge_cells(f"A{row}:N{row}")
    set_cell(ws,row,1,"▼ 通常期（月〜木・日）", font=bold(10), align=left(), fill=PatternFill("solid",fgColor="E3F2FD"))
    row += 1

    staff_normal = [
        ("A（オーナー）", "―",   "―",    [1]*6+[0]*6),
        ("B（バイト）",   1200, 7500,   [1,1,1,1,1,1,0,0,0,0,0,0]),
        ("C（バイト）",   1200, 9000,   [0,0,0,0,0,0,1,1,1,1,1,1]),
        ("D（バイト）",   1200, "―",   [0]*12),
    ]
    normal_amounts = {"B（バイト）":7500, "C（バイト）":9000}
    total_normal = sum(normal_amounts.values())

    for name, wage, amt, shifts in staff_normal:
        set_cell(ws,row,1,name, font=normal(), align=center(), fill=LGRAY, border=all_border())
        set_cell(ws,row,2,wage, font=normal(), align=center(), fill=LGRAY, border=all_border())
        a = normal_amounts.get(name, "―")
        set_cell(ws,row,3,a, font=normal(), align=right(), fill=LGREEN, border=all_border(), fmt=NUM if isinstance(a,int) else None)
        for ci, s in enumerate(shifts, 4):
            fill = PatternFill("solid",fgColor="1565C0") if s else WHITE
            set_cell(ws,row,ci,"●" if s else "", font=Font(color="FFFFFF" if s else "000000",size=9), align=center(), fill=fill, border=all_border())
        set_cell(ws,row,16,sum(shifts), font=bold(), align=center(), fill=LGRAY, border=all_border())
        row += 1

    set_cell(ws,row,1,"計", font=bold(), align=center(), fill=ORANGE, border=all_border())
    set_cell(ws,row,2,"",   fill=ORANGE, border=all_border())
    set_cell(ws,row,3,total_normal, font=bold(), align=right(), fill=ORANGE, border=all_border(), fmt=NUM)
    for ci in range(4,17):
        set_cell(ws,row,ci,"", fill=ORANGE, border=all_border())
    row += 3

    # --- 月次人件費集計 ---
    ws.merge_cells(f"A{row}:G{row}")
    set_cell(ws,row,1,"【月次人件費集計】", font=bold(11), align=left())
    row += 1

    # ヘッダー
    for ci, h in enumerate(["月","区分","給与額/日","日数","給与総額"],1):
        set_cell(ws,row,ci,h, font=bold(9,white=True), align=center(), fill=HEADER, border=all_border())
    row += 1

    months_days = {
        1: {"平日":17,"金曜":4,"土曜":4,"日祝":6},
        2: {"平日":15,"金曜":4,"土曜":4,"日祝":5},
        3: {"平日":20,"金曜":5,"土曜":5,"日祝":1},
        4: {"平日":20,"金曜":4,"土曜":4,"日祝":2},
        5: {"平日":18,"金曜":4,"土曜":4,"日祝":5},
        6: {"平日":21,"金曜":4,"土曜":4,"日祝":1},
        7: {"平日":22,"金曜":4,"土曜":4,"日祝":1},
        8: {"平日":21,"金曜":4,"土曜":4,"日祝":2},
        9: {"平日":21,"金曜":4,"土曜":4,"日祝":1},
        10:{"平日":22,"金曜":5,"土曜":4,"日祝":0},
        11:{"平日":20,"金曜":4,"土曜":4,"日祝":2},
        12:{"平日":20,"金曜":4,"土曜":4,"日祝":3},
    }
    # 金・土は繁忙期=25500/日, 平日・日祝は通常期=16500/日
    day_wages = {"平日":16500,"金曜":25500,"土曜":25500,"日祝":16500}

    for m in range(1,13):
        d = months_days[m]
        month_total_rows = []
        first = True
        for dtype, days in d.items():
            wage = day_wages[dtype]
            total = wage * days
            if first:
                set_cell(ws,row,1,f"{m}月", font=bold(), align=center(), fill=LGRAY, border=all_border())
                first = False
            else:
                set_cell(ws,row,1,"", fill=LGRAY, border=all_border())
            set_cell(ws,row,2,dtype,  font=normal(), align=center(), fill=LGRAY, border=all_border())
            set_cell(ws,row,3,wage,   font=normal(), align=right(),  fill=YELLOW,border=all_border(), fmt=NUM)
            set_cell(ws,row,4,days,   font=normal(), align=center(), fill=YELLOW,border=all_border(), fmt=NUM)
            set_cell(ws,row,5,total,  font=normal(), align=right(),  fill=LGREEN,border=all_border(), fmt=NUM)
            month_total_rows.append(f"E{row}")
            row += 1
        # 月合計
        set_cell(ws,row,1,f"{m}月合計", font=bold(), align=center(), fill=ORANGE, border=all_border())
        ws.merge_cells(f"B{row}:D{row}")
        set_cell(ws,row,2,"", fill=ORANGE, border=all_border())
        set_cell(ws,row,5,f"=SUM({','.join(month_total_rows)})", font=bold(), align=right(), fill=ORANGE, border=all_border(), fmt=NUM)
        row += 1

    # 列幅
    for c, w in [(1,14),(2,10),(3,12),(4,8),(5,12),(6,6),(7,6),(8,6),(9,6),(10,6),(11,6),(12,6),(13,6),(14,6),(15,6),(16,8)]:
        ws.column_dimensions[get_column_letter(c)].width = w

# =====================================================================
# 3. 月次損益計画書
# =====================================================================
def make_monthly_pl(wb):
    ws = wb.create_sheet("月次損益計画書")
    ws.sheet_view.showGridLines = False

    ws.merge_cells("A1:O1")
    set_cell(ws,1,1,"■ 月次損益計画書（開業後1年）", font=bold(14), align=left())
    ws.merge_cells("M2:O2")
    set_cell(ws,2,13,"（単位／円）", font=normal(9), align=right())

    months = [f"{m}月" for m in range(1,13)] + ["合計"]
    headers = ["項目","区分"] + months

    for ci, h in enumerate(headers, 1):
        set_cell(ws,3,ci,h, font=bold(9,white=True), align=center(), fill=HEADER, border=all_border())

    # 売上月次（売上計画検討表のデータをもとに設定）
    monthly_shisha = [1000000,1000000,1300000,1300000,1400000,1400000,1500000,1500000,1500000,1600000,1600000,1700000]
    monthly_drink  = [int(s*0.3) for s in monthly_shisha]

    row = 4
    def add_row(label, sub, values, fill=WHITE, bold_=False, formula_col=None, indent=False):
        nonlocal row
        lbl = "　" + label if indent else label
        set_cell(ws,row,1,lbl, font=bold(9) if bold_ else normal(9), align=left(), fill=fill, border=all_border())
        set_cell(ws,row,2,sub, font=normal(9), align=center(), fill=fill, border=all_border())
        for ci, v in enumerate(values, 3):
            if isinstance(v, str):
                set_cell(ws,row,ci,v, font=bold(9) if bold_ else normal(9), align=right(), fill=fill, border=all_border(), fmt=NUM)
            else:
                set_cell(ws,row,ci,v if v is not None else 0,
                         font=bold(9) if bold_ else normal(9), align=right(), fill=fill, border=all_border(), fmt=NUM)
        # 合計列
        if formula_col:
            set_cell(ws,row,15,formula_col, font=bold(9) if bold_ else normal(9), align=right(), fill=fill if fill!=WHITE else LGRAY, border=all_border(), fmt=NUM)
        else:
            total = sum(v for v in values if isinstance(v,(int,float)))
            set_cell(ws,row,15,total, font=bold(9) if bold_ else normal(9), align=right(), fill=fill if fill!=WHITE else LGRAY, border=all_border(), fmt=NUM)
        row += 1
        return row - 1

    # 売上
    r_shisha = add_row("シーシャ売上","", monthly_shisha, fill=WHITE, indent=True)
    r_drink  = add_row("ドリンク売上","", monthly_drink,  fill=WHITE, indent=True)
    # 売上合計
    shisha_total = sum(monthly_shisha)
    drink_total  = sum(monthly_drink)
    monthly_sales = [monthly_shisha[i]+monthly_drink[i] for i in range(12)]
    r_sales = add_row("売上合計","A", monthly_sales, fill=LGREEN, bold_=True)

    # 原価
    shisha_cost = [int(s*0.30) for s in monthly_shisha]
    drink_cost  = [int(s*0.20) for s in monthly_drink]
    add_row("シーシャ原価","原価30%", shisha_cost, fill=WHITE, indent=True)
    add_row("ドリンク原価","原価20%", drink_cost,  fill=WHITE, indent=True)
    monthly_cost = [shisha_cost[i]+drink_cost[i] for i in range(12)]
    r_cost = add_row("売上原価合計","B", monthly_cost, fill=LGRAY, bold_=True)

    # 粗利
    monthly_gross = [monthly_sales[i]-monthly_cost[i] for i in range(12)]
    r_gross = add_row("売上総利益（粗利）","C=A-B", monthly_gross, fill=LGREEN, bold_=True)

    # 販管費
    rent    = [300000]*12
    labor   = [int(m*0.57) for m in [1,1,1,1,1,1,1,1,1,1,1,1]]  # placeholder
    # 実際の人件費（月によって異なる）
    labor_actual = [451200,534400,664200,534400,534400,534400,609400,609400,534400,593900,534400,593900]
    owner_sal = [220000]*12
    util    = [80000]*12
    adv     = [30000,30000,20000,20000,20000,10000,10000,10000,10000,10000,10000,10000]
    misc    = [50000]*12

    add_row("役員報酬","固定", owner_sal, fill=WHITE, indent=True)
    add_row("人件費（バイト）","変動", labor_actual, fill=WHITE, indent=True)
    add_row("賃借料","固定", rent, fill=WHITE, indent=True)
    add_row("水道光熱費","固定", util, fill=WHITE, indent=True)
    add_row("広告宣伝費","変動", adv, fill=WHITE, indent=True)
    add_row("消耗品・雑費","変動", misc, fill=WHITE, indent=True)

    monthly_sga = [owner_sal[i]+labor_actual[i]+rent[i]+util[i]+adv[i]+misc[i] for i in range(12)]
    r_sga = add_row("販管費合計","D", monthly_sga, fill=LGRAY, bold_=True)

    # 営業利益
    monthly_op = [monthly_gross[i]-monthly_sga[i] for i in range(12)]
    r_op = add_row("営業利益","E=C-D", monthly_op, fill=LGREEN, bold_=True)

    # 支払利息
    interest = [15000]*12
    add_row("支払利息","",interest, fill=WHITE, indent=True)

    # 経常利益
    monthly_net = [monthly_op[i]-interest[i] for i in range(12)]
    add_row("経常利益","F=E-支払利息", monthly_net, fill=ORANGE, bold_=True)

    # 列幅
    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 10
    for c in range(3,16):
        ws.column_dimensions[get_column_letter(c)].width = 11

# =====================================================================
# 4. 年次損益計画書
# =====================================================================
def make_annual_pl(wb):
    ws = wb.create_sheet("年次損益計画書")
    ws.sheet_view.showGridLines = False

    ws.merge_cells("A1:H1")
    set_cell(ws,1,1,"■ 年次損益計画書（3ヵ年）", font=bold(14), align=left())
    ws.merge_cells("G2:H2")
    set_cell(ws,2,7,"（単位／千円）", font=normal(9), align=right())

    headers = ["項目","区分","1年目\n金額","1年目\n構成比","2年目\n金額","2年目\n構成比","3年目\n金額","3年目\n構成比","備考"]
    for ci, h in enumerate(headers,1):
        set_cell(ws,3,ci,h, font=bold(9,white=True), align=center(wrap=True), fill=HEADER, border=all_border())
    ws.row_dimensions[3].height = 30

    row = 4
    def add_annual(label, sub, y1, y2, y3, fill=WHITE, bold_=False, indent=False, note=""):
        nonlocal row
        lbl = "　"+label if indent else label
        s1 = sum(y1) if isinstance(y1,list) else y1
        s2 = sum(y2) if isinstance(y2,list) else y2
        s3 = sum(y3) if isinstance(y3,list) else y3
        set_cell(ws,row,1,lbl,  font=bold(9) if bold_ else normal(9), align=left(), fill=fill, border=all_border())
        set_cell(ws,row,2,sub,  font=normal(9), align=center(), fill=fill, border=all_border())
        set_cell(ws,row,3,round(s1/1000), font=bold(9) if bold_ else normal(9), align=right(), fill=fill, border=all_border(), fmt=NUM)
        set_cell(ws,row,4,"100.0%" if "合計" in label else "", font=normal(9), align=center(), fill=fill, border=all_border())
        set_cell(ws,row,5,round(s2/1000), font=bold(9) if bold_ else normal(9), align=right(), fill=fill, border=all_border(), fmt=NUM)
        set_cell(ws,row,6,"100.0%" if "合計" in label else "", font=normal(9), align=center(), fill=fill, border=all_border())
        set_cell(ws,row,7,round(s3/1000), font=bold(9) if bold_ else normal(9), align=right(), fill=fill, border=all_border(), fmt=NUM)
        set_cell(ws,row,8,"100.0%" if "合計" in label else "", font=normal(9), align=center(), fill=fill, border=all_border())
        set_cell(ws,row,9,note, font=normal(8,"666666"), align=left(wrap=True), fill=fill, border=all_border())
        row += 1

    # 年次売上データ
    y1_shisha = 15000000; y2_shisha = 18000000; y3_shisha = 20000000
    y1_drink  = int(y1_shisha*0.3); y2_drink=int(y2_shisha*0.3); y3_drink=int(y3_shisha*0.3)
    y1_sales  = y1_shisha+y1_drink; y2_sales=y2_shisha+y2_drink; y3_sales=y3_shisha+y3_drink

    add_annual("シーシャ売上","", y1_shisha, y2_shisha, y3_shisha, indent=True)
    add_annual("ドリンク売上","", y1_drink,  y2_drink,  y3_drink,  indent=True)
    add_annual("売上合計","A",    y1_sales,  y2_sales,  y3_sales,  fill=LGREEN, bold_=True)

    y1_cost=int(y1_sales*0.27); y2_cost=int(y2_sales*0.27); y3_cost=int(y3_sales*0.27)
    add_annual("売上原価","B",    y1_cost, y2_cost, y3_cost, fill=LGRAY, bold_=True, note="原価率27%（シーシャ30%・ドリンク20%）")

    y1_gross=y1_sales-y1_cost; y2_gross=y2_sales-y2_cost; y3_gross=y3_sales-y3_cost
    add_annual("売上総利益","C=A-B", y1_gross, y2_gross, y3_gross, fill=LGREEN, bold_=True)

    y1_owner=2640000; y2_owner=2640000; y3_owner=2640000
    y1_labor=6864000; y2_labor=7200000; y3_labor=7200000
    y1_rent =3600000; y2_rent =3600000; y3_rent =3600000
    y1_util = 960000; y2_util = 960000; y3_util = 960000
    y1_adv  = 240000; y2_adv  = 180000; y3_adv  = 120000
    y1_misc = 600000; y2_misc = 600000; y3_misc = 600000

    add_annual("役員報酬","固定",     y1_owner, y2_owner, y3_owner, indent=True, note="月22万円×12ヶ月")
    add_annual("人件費（バイト）","変動", y1_labor, y2_labor, y3_labor, indent=True, note="月平均57万円")
    add_annual("賃借料","固定",       y1_rent,  y2_rent,  y3_rent,  indent=True, note="月30万円×12ヶ月")
    add_annual("水道光熱費","固定",   y1_util,  y2_util,  y3_util,  indent=True, note="月8万円×12ヶ月")
    add_annual("広告宣伝費","変動",   y1_adv,   y2_adv,   y3_adv,   indent=True)
    add_annual("消耗品・雑費","変動", y1_misc,  y2_misc,  y3_misc,  indent=True)

    y1_sga=y1_owner+y1_labor+y1_rent+y1_util+y1_adv+y1_misc
    y2_sga=y2_owner+y2_labor+y2_rent+y2_util+y2_adv+y2_misc
    y3_sga=y3_owner+y3_labor+y3_rent+y3_util+y3_adv+y3_misc
    add_annual("販管費合計","D", y1_sga, y2_sga, y3_sga, fill=LGRAY, bold_=True)

    y1_op=y1_gross-y1_sga; y2_op=y2_gross-y2_sga; y3_op=y3_gross-y3_sga
    add_annual("営業利益","E=C-D", y1_op, y2_op, y3_op, fill=LGREEN, bold_=True)

    interest = 1050000  # 700万×2.5%×6年平均
    y1_int=180000; y2_int=162000; y3_int=144000
    add_annual("支払利息","", y1_int, y2_int, y3_int, indent=True, note="700万×2.5%（元本逓減）")

    y1_net=y1_op-y1_int; y2_net=y2_op-y2_int; y3_net=y3_op-y3_int

    # 返済
    repay = 840000  # 700万/10年×12ヶ月
    add_annual("経常利益","F", y1_net, y2_net, y3_net, fill=ORANGE, bold_=True)

    row += 1
    set_cell(ws,row,1,"【参考指標】", font=bold(10), align=left())
    row += 1
    indicators = [
        ("原価率（目標30%以下）",f"{round(y1_cost/y1_sales*100,1)}%",f"{round(y2_cost/y2_sales*100,1)}%",f"{round(y3_cost/y3_sales*100,1)}%"),
        ("実質人件費率（目標55%以内）",f"{round((y1_owner+y1_labor)/y1_sales*100,1)}%",f"{round((y2_owner+y2_labor)/y2_sales*100,1)}%",f"{round((y3_owner+y3_labor)/y3_sales*100,1)}%"),
        ("損益分岐点売上高",f"{round((y1_sga)/(1-y1_cost/y1_sales)/10000)}万円",f"{round((y2_sga)/(1-y2_cost/y2_sales)/10000)}万円",f"{round((y3_sga)/(1-y3_cost/y3_sales)/10000)}万円"),
        ("借入返済額（年）","840万円","840万円","840万円"),
        ("本返済後手残り（概算）",f"{round((y1_net-8400000)/10000)}万円",f"{round((y2_net-8400000)/10000)}万円",f"{round((y3_net-8400000)/10000)}万円"),
    ]
    for label, v1, v2, v3 in indicators:
        set_cell(ws,row,1,label, font=normal(9), align=left(), fill=LGRAY, border=all_border())
        set_cell(ws,row,3,v1,    font=bold(9),   align=center(), fill=YELLOW, border=all_border())
        set_cell(ws,row,5,v2,    font=bold(9),   align=center(), fill=YELLOW, border=all_border())
        set_cell(ws,row,7,v3,    font=bold(9),   align=center(), fill=YELLOW, border=all_border())
        row += 1

    # 列幅
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 10
    for c in [3,4,5,6,7,8]:
        ws.column_dimensions[get_column_letter(c)].width = 12
    ws.column_dimensions["I"].width = 30

# =====================================================================
# MAIN
# =====================================================================
wb = openpyxl.Workbook()
wb.remove(wb.active)  # デフォルトシート削除

make_uriage(wb)
make_labor(wb)
make_monthly_pl(wb)
make_annual_pl(wb)

path = "/home/user/kuon_shisha/シーシャバー_創業融資資料.xlsx"
wb.save(path)
print(f"保存完了: {path}")
