# -*- coding: utf-8 -*-
"""部位维护表.xlsx 生成脚本 — 义体记录项目"""
try:
    import openpyxl
except ImportError:
    import subprocess, sys
    subprocess.check_call([sys.executable, "-m", "pip", "install", "--quiet", "openpyxl>=3.1.0"])
    import openpyxl

from openpyxl.styles import Font, PatternFill, Alignment, Side, Border
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule
from openpyxl.utils import get_column_letter

def xl_color(css_hex: str) -> str:
    value = css_hex.removeprefix("#").upper()
    if len(value) != 6:
        raise ValueError(f"Expected #RRGGBB, got: {css_hex}")
    return "FF" + value

XL_HEADER_BG = xl_color("#4472C4")
XL_INPUT_BG = xl_color("#D9E2F3")
XL_BORDER = xl_color("#BFBFBF")
XL_GREEN_BG = xl_color("#C6EFCE"); XL_GREEN_FG = xl_color("#006100")
XL_RED_BG = xl_color("#FFC7CE");  XL_RED_FG = xl_color("#9C0006")
XL_YELLOW_BG = xl_color("#FFEB9C"); XL_YELLOW_FG = xl_color("#9C6500")

thin = Side(style="thin", color=XL_BORDER)
border_all = Border(left=thin, right=thin, top=thin, bottom=thin)
font_header = Font(bold=True, color="FFFFFFFF", size=11)
font_body = Font(size=11)
fill_header = PatternFill("solid", fgColor=XL_HEADER_BG)
fill_input = PatternFill("solid", fgColor=XL_INPUT_BG)
align_c = Alignment(horizontal="center", vertical="center")
align_l = Alignment(horizontal="left", vertical="center", wrap_text=True)

# (部位ID, 部位名称, 所属视图, 身体区域, 对称性, 备注)
PARTS = [
    # 头部 — 正面
    ("head",        "头",     "正面", "头部", "中线", ""),
    ("left_eye",    "左眼",   "正面", "头部", "左",   ""),
    ("right_eye",   "右眼",   "正面", "头部", "右",   ""),
    ("nose",        "鼻",     "正面", "头部", "中线", ""),
    ("mouth",       "嘴",     "正面", "头部", "中线", ""),
    # 颈部
    ("neck_front",  "颈前",   "正面", "颈部", "中线", ""),
    ("neck_back",   "后颈",   "背面", "颈部", "中线", ""),
    # 头部 — 背面
    ("back_head",   "后脑",   "背面", "头部", "中线", ""),
    ("hair",        "头发",   "双面", "头部", "中线", "v0.0.7 新增，正/背面各一套图形，同一份记录"),
    # 耳（双面）
    ("left_ear",    "左耳",   "双面", "头部", "左",   ""),
    ("right_ear",   "右耳",   "双面", "头部", "右",   ""),
    # 躯干 — 正面
    ("chest",       "胸",     "正面", "躯干", "中线", ""),
    ("abdomen",     "腹",     "正面", "躯干", "中线", ""),
    ("hips",        "髋部",   "正面", "躯干", "中线", ""),
    ("genitals",    "生殖器", "正面", "躯干", "中线", "v0.0.3新增，小长方形表示"),
    # 躯干 — 背面
    ("upper_back",  "上背",   "背面", "躯干", "中线", ""),
    ("lower_back",  "腰",     "背面", "躯干", "中线", ""),
    ("left_glute",  "左臀",   "背面", "躯干", "左",   ""),
    ("right_glute", "右臀",   "背面", "躯干", "右",   ""),
    ("anus",        "肛门",   "背面", "躯干", "中线", "v0.0.4新增，小长方形，臀部中间靠下"),
    # 肩（双面）
    ("left_shoulder",  "左肩",  "双面", "躯干", "左", ""),
    ("right_shoulder", "右肩",  "双面", "躯干", "右", ""),
    # 上肢（双面）
    ("left_upper_arm",  "左大臂", "双面", "上肢", "左", ""),
    ("right_upper_arm", "右大臂", "双面", "上肢", "右", ""),
    ("left_elbow",      "左肘",   "双面", "上肢", "左", ""),
    ("right_elbow",     "右肘",   "双面", "上肢", "右", ""),
    ("left_forearm",    "左小臂", "双面", "上肢", "左", ""),
    ("right_forearm",   "右小臂", "双面", "上肢", "右", ""),
    ("left_wrist",      "左腕",   "双面", "上肢", "左", ""),
    ("right_wrist",     "右腕",   "双面", "上肢", "右", ""),
    ("left_palm",       "左手掌", "正面", "上肢", "左", ""),
    ("right_palm",      "右手掌", "正面", "上肢", "右", ""),
    ("left_hand_back",  "左手背", "背面", "上肢", "左", ""),
    ("right_hand_back", "右手背", "背面", "上肢", "右", ""),
    ("left_fingers",    "左手指", "双面", "上肢", "左", ""),
    ("right_fingers",   "右手指", "双面", "上肢", "右", ""),
    # 下肢（双面）
    ("left_thigh",  "左大腿", "双面", "下肢", "左", ""),
    ("right_thigh", "右大腿", "双面", "下肢", "右", ""),
    ("left_knee",   "左膝",   "双面", "下肢", "左", ""),
    ("right_knee",  "右膝",   "双面", "下肢", "右", ""),
    ("left_calf",   "左小腿", "双面", "下肢", "左", ""),
    ("right_calf",  "右小腿", "双面", "下肢", "右", ""),
    ("left_ankle",  "左脚踝", "双面", "下肢", "左", ""),
    ("right_ankle", "右脚踝", "双面", "下肢", "右", ""),
    ("left_instep", "左脚背", "双面", "下肢", "左", "v0.0.8 脚部改为三层：脚背在上层，脚底在脚背下方，脚趾在脚背前端；三层不分正反都显示"),
    ("right_instep","右脚背", "双面", "下肢", "右", "v0.0.8 同上"),
    ("left_sole",   "左脚底", "双面", "下肢", "左", "v0.0.8 改名（原「左脚底」），改为双面，紧贴脚背下方"),
    ("right_sole",  "右脚底", "双面", "下肢", "右", "v0.0.8 同上"),
    ("left_toes",   "左脚趾", "双面", "下肢", "左", "v0.0.8 位于脚背前端"),
    ("right_toes",  "右脚趾", "双面", "下肢", "右", "v0.0.8 同上"),
]

HEADERS = ["序号", "部位ID", "部位名称", "所属视图", "身体区域", "对称性", "状态", "备注"]
COL_WIDTHS = [6, 16, 10, 10, 10, 8, 8, 36]

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "部位清单"

for ci, (h, w) in enumerate(zip(HEADERS, COL_WIDTHS), start=1):
    cell = ws.cell(row=1, column=ci, value=h)
    cell.font = font_header
    cell.fill = fill_header
    cell.alignment = align_c
    cell.border = border_all
    ws.column_dimensions[get_column_letter(ci)].width = w
ws.row_dimensions[1].height = 22

data_start, data_end = 2, 1 + len(PARTS)          # 2..44
reserved_rows = 5
row_end = data_end + reserved_rows                 # 49

for i, (pid, name, view, region, side, note) in enumerate(PARTS):
    r = data_start + i
    vals = [i + 1, pid, name, view, region, side, "已有", note]
    for ci, v in enumerate(vals, start=1):
        cell = ws.cell(row=r, column=ci, value=v)
        cell.font = font_body
        cell.border = border_all
        cell.alignment = align_l if ci == 8 else align_c

for r in range(data_end + 1, row_end + 1):         # 预留录入行
    for ci in range(1, 9):
        cell = ws.cell(row=r, column=ci)
        cell.font = font_body
        cell.border = border_all
        cell.fill = fill_input
        cell.alignment = align_l if ci == 8 else align_c

dv_view = DataValidation(type="list", formula1='"正面,背面,双面"', allow_blank=True)
dv_region = DataValidation(type="list", formula1='"头部,颈部,躯干,上肢,下肢"', allow_blank=True)
dv_side = DataValidation(type="list", formula1='"左,右,中线"', allow_blank=True)
dv_status = DataValidation(type="list", formula1='"已有,新增,删除"', allow_blank=True)
for dv, col in ((dv_view, "D"), (dv_region, "E"), (dv_side, "F"), (dv_status, "G")):
    ws.add_data_validation(dv)
    dv.add(f"{col}{data_start}:{col}{row_end}")

status_range = f"G{data_start}:G{row_end}"
ws.conditional_formatting.add(status_range, CellIsRule(
    operator="equal", formula=['"新增"'],
    fill=PatternFill("solid", fgColor=XL_GREEN_BG), font=Font(color=XL_GREEN_FG)))
ws.conditional_formatting.add(status_range, CellIsRule(
    operator="equal", formula=['"删除"'],
    fill=PatternFill("solid", fgColor=XL_RED_BG), font=Font(color=XL_RED_FG)))

ws.freeze_panes = "A2"
ws.auto_filter.ref = f"A1:H{row_end}"

# ---- 填写说明 sheet ----
ws2 = wb.create_sheet("填写说明")
ws2.column_dimensions["A"].width = 14
ws2.column_dimensions["B"].width = 80
title = ws2.cell(row=1, column=1, value="部位维护表 — 使用说明")
title.font = Font(bold=True, size=13, color=XL_HEADER_BG[2:])
notes = [
    ("新增部位", "在「部位清单」末尾浅蓝色空白行填写：部位ID、部位名称、所属视图、身体区域、对称性，状态选「新增」。可整行复制格式后向下扩展。"),
    ("删除部位", "将该部位所在行的「状态」改为「删除」，不需要物理删行。"),
    ("修改部位", "直接修改「部位名称」，并在「备注」写明改动原因（如：改名/换视图）。"),
    ("部位ID规范", "小写英文加下划线，如 left_wrist；左右对称部位用 left_ / right_ 前缀；中线部位不加前缀。新增部位若不确定ID可留空由AI生成。"),
    ("所属视图", "正面=只在正面人体图显示；背面=只在背面人体图显示；双面=正反两图都显示（如四肢、耳）。"),
    ("状态列", "已有=当前网页已实现；新增=待AI加入网页；删除=待AI从网页移除。"),
    ("更新流程", "在表中改完后，把表格发给AI（或在对话中说“部位表已更新”），AI会据此修改网页。"),
]
for i, (k, v) in enumerate(notes, start=3):
    c1 = ws2.cell(row=i, column=1, value=k)
    c1.font = Font(bold=True)
    c1.alignment = Alignment(horizontal="left", vertical="center")
    c2 = ws2.cell(row=i, column=2, value=v)
    c2.font = font_body
    c2.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    ws2.row_dimensions[i].height = 34

wb.properties.title = "部位维护表"
wb.save(r"F:\AIproject\义体记录\文档\部位维护表.xlsx")
print("OK 部位维护表")
