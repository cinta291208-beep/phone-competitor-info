#!/usr/bin/env python3
"""生成"手机竞品基础信息收集"的标准 Excel 模板。

结构：列 A=分组（前置/后置/屏幕/其他），列 B=字段名，之后每列一个机型。
冻结表头，分组合并底色标识，字段行含说明。

用法:
    python3 build_template.py <输出xlsx路径> [机型列数]

依赖: openpyxl
"""

import sys
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# 字段结构: (分组, [(字段名, 说明), ...])  共 24 字段
FIELDS = [
    ("前置", [
        ("前置摄像头", "数量/传感器/像素/光圈"),
        ("照片分辨率", "前置最大拍照分辨率"),
        ("摄像分辨率", "前置视频最大规格"),
        ("拍摄功能", "人像/美颜/夜景等"),
        ("人脸识别", "是否支持及类型"),
    ]),
    ("后置", [
        ("主摄", "传感器/像素/光圈/防抖"),
        ("长焦", "变焦倍数/像素/光圈"),
        ("超广角", "视场角/像素/光圈"),
        ("照片分辨率", "后置最大拍照分辨率"),
        ("摄像分辨率", "后置视频最大规格"),
        ("拍摄功能", "夜景/人像/电影模式等"),
    ]),
    ("屏幕", [
        ("尺寸", "对角线英寸"),
        ("分辨率", "如 3200×1440"),
        ("色彩", "色深/色域"),
        ("类型", "OLED/AMOLED 等"),
        ("亮度", "峰值亮度 nit"),
        ("护眼", "PWM 调光/认证"),
        ("显示技术", "LTPO/刷新率"),
        ("保护玻璃材质", "大猩猩/昆仑等"),
        ("触控", "采样率/刷新率"),
    ]),
    ("其他", [
        ("宣传点", "核心卖点"),
        ("官网介绍链接", "必填"),
        ("博主测评链接", "博主+平台+链接"),
        ("补充信息", "其他"),
    ]),
]

HEADER_FILL = PatternFill("solid", fgColor="4472C4")
GROUP_FILL = PatternFill("solid", fgColor="D9E2F3")
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
THIN = Side(style="thin", color="999999")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def build(path: str, model_cols: int = 3) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "竞品信息"

    # 表头行
    ws.cell(1, 1, "分组")
    ws.cell(1, 2, "字段")
    for i in range(model_cols):
        ws.cell(1, 3 + i, f"机型{i + 1}")
    for c in range(1, 3 + model_cols):
        cell = ws.cell(1, c)
        cell.fill = HEADER_FILL
        cell.font = Font(bold=True, color="FFFFFF")
        cell.alignment = CENTER
        cell.border = BORDER

    # 字段行
    row = 2
    for group, fields in FIELDS:
        for name, _note in fields:
            g = ws.cell(row, 1, group)
            n = ws.cell(row, 2, name)
            g.fill = GROUP_FILL
            g.alignment = LEFT
            n.alignment = LEFT
            g.border = BORDER
            n.border = BORDER
            for i in range(model_cols):
                c = ws.cell(row, 3 + i, "")
                c.alignment = LEFT
                c.border = BORDER
            row += 1

    # 列宽 / 行高 / 冻结
    ws.column_dimensions["A"].width = 9
    ws.column_dimensions["B"].width = 16
    for i in range(model_cols):
        ws.column_dimensions[get_column_letter(3 + i)].width = 46
    ws.row_dimensions[1].height = 22
    for r in range(2, row):
        ws.row_dimensions[r].height = 34
    ws.freeze_panes = "C2"

    wb.save(path)
    total = sum(len(f) for _, f in FIELDS)
    print(f"模板已生成: {path}（{total} 个字段，预留 {model_cols} 个机型列）")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("用法: python3 build_template.py <输出xlsx路径> [机型列数]")
        sys.exit(1)
    cols = int(sys.argv[2]) if len(sys.argv) > 2 else 3
    build(sys.argv[1], cols)
