# -*- coding: utf-8 -*-
"""
converter.py
------------
Chuyển đổi BOM kỹ thuật (dạng cây: Cụm > Chi tiết > Vật tư) sang
BOM ERP (dạng bảng phẳng: STT / QTSX / Mã sản phẩm / Loại / ... / Công đoạn).

Cách dùng nhanh:
    from converter import load_technical_bom, convert_to_erp, DEFAULT_RULES
    header, nodes = load_technical_bom("BOM_ky_thuat.xlsx")
    rows = convert_to_erp(header, nodes, DEFAULT_RULES)
    # rows là list[dict], mỗi dict có đủ các cột ERP

Ghi chú quan trọng (đọc kỹ trước khi tin tưởng 100% kết quả):
    File BOM kỹ thuật chỉ lưu thông tin THIẾT KẾ (kích thước, vật liệu,
    khối lượng). File BOM ERP còn lưu thêm THÔNG TIN SẢN XUẤT (QTSX,
    Công đoạn, việc tách 1 chi tiết thành nhiều bước gia công như
    CUT -> ĐÓNG PEM -> WELD...). Thông tin sản xuất này KHÔNG có sẵn
    trong BOM kỹ thuật, nên công cụ chỉ có thể "đoán" theo từ khoá
    (xem RULES["process_keywords"]). Kỹ sư/QC vẫn cần rà soát lại
    cột "Công đoạn" và việc gộp/tách các cụm phụ kiện (kit) trước khi
    nạp vào ERP thật.
"""
from __future__ import annotations

import re
import json
from dataclasses import dataclass, field
from typing import Optional

import openpyxl
from openpyxl.utils import get_column_letter


ERP_COLUMNS = [
    "STT", "QTSX", "Mã sản phẩm", "Loại", "Tên sản phẩm", "SL sản phẩm",
    "ĐVT sp", "Mã cũ", "Mã NVL", "Tên NVL", "ĐVT (NVL)", "Số lượng",
    "ĐVT Kho", "Số lượng kho", "Số lượng setup", "% Tỉ lệ hao hụt", "Công đoạn",
]

RAW_MATERIALS = {"steel", "inox", "aluminium", "aluminum", "copper", "zinc"}

DEFAULT_RULES = {
    # Shortcode khách hàng dùng để ghép mã: "SG - <KH> - <code>"
    "customer_shortcode": "STV",
    # Các tiền tố mã vẽ nội bộ (drawing code) sẽ được gắn "<Loại> - KH - <code>"
    # Mã KHÔNG nằm trong danh sách này (vd FA, D, PM, PR, STX, số thuần...)
    # được xem là mã hàng mua/thương mại -> giữ nguyên, không thêm tiền tố.
    "drawing_prefixes": ["GZ", "ME", "CA", "CE", "RA"],
    # Từ khoá trong mô tả -> Công đoạn gợi ý (khớp từ trên xuống, không phân biệt hoa/thường,
    # không dấu cũng được vì có bỏ dấu trước khi so khớp)
    "process_keywords": [
        ["tu giu", "ĐÓNG PEM"],
        ["pem", "ĐÓNG PEM"],
        ["vermiculite", "AS"],
        ["duc", "Mua"],
        ["kinh", "Mua"],
        ["nhan dan", "Packing"],
        ["nhan ", "Packing"],
        ["tui ", "Packing"],
        ["huong dan", "Packing"],
        ["bao hanh", "Packing"],
        ["catalogue", "Packing"],
        ["hop ", "Packing"],
        ["gang tay", "Packing"],
    ],
    # Tên các nhóm (Section, cột A) được coi là "vật tư đóng gói" -> Công đoạn mặc định = Packing
    "packaging_sections": ["VẬT TƯ ĐÓNG GÓI", "VẬT TƯ CONTAINER"],
    "default_process_sheet_metal": "CUT",
    "default_process_purchased": "Mua",
    "default_unit_piece": "Cái",
    "default_unit_weight": "Kg",
}


def _strip_accents(s: str) -> str:
    import unicodedata
    if s is None:
        return ""
    nfkd = unicodedata.normalize("NFKD", s)
    return "".join(c for c in nfkd if not unicodedata.combining(c)).lower()


def _norm(v):
    if v is None:
        return None
    if isinstance(v, str):
        v = v.strip()
        return v if v != "" else None
    return v


@dataclass
class TechNode:
    row: int
    level: int  # 1=Cụm(section col A) is level 0 actually; see loader for real semantics
    section_no: Optional[str]
    section_name: Optional[str]
    cum_no: Optional[str]          # cột B
    part_no: Optional[str]         # cột C
    detail_no: Optional[str]       # cột D
    ma: Optional[str]
    mo_ta: Optional[str]
    length: Optional[float]
    width: Optional[float]
    thickness: Optional[float]
    material: Optional[str]        # cột "Loại vật liệu" (Steel/Inox/.../Ea/Set)
    density: Optional[float]
    wastage: Optional[float]
    qty: Optional[float]
    raw_weight: Optional[float]
    finished_weight: Optional[float]
    children: list = field(default_factory=list)


def load_technical_bom(path: str, sheet_name: str = "Form"):
    """Đọc file BOM kỹ thuật (đúng layout mẫu Stovax) và trả về:
        header: dict {khach_hang, du_an, revision, nguoi_lap}
        sections: list[dict] mỗi section có {name, no, items:[TechNode cấp Cụm...]}
    """
    wb = openpyxl.load_workbook(path, data_only=True)
    if sheet_name not in wb.sheetnames:
        sheet_name = wb.sheetnames[0]
    ws = wb[sheet_name]

    header = {"khach_hang": None, "du_an": None, "revision": None, "nguoi_lap": None}
    header_row = None
    for r in range(1, 15):
        a = _norm(ws.cell(row=r, column=1).value)
        if not isinstance(a, str):
            a = None
        if a and "khách hàng" in a.lower():
            header["khach_hang"] = _norm(ws.cell(row=r, column=6).value) or _norm(ws.cell(row=r, column=2).value)
        if a and "dự án" in a.lower():
            header["du_an"] = _norm(ws.cell(row=r, column=6).value) or _norm(ws.cell(row=r, column=2).value)
        if a and "revision" in a.lower():
            header["revision"] = _norm(ws.cell(row=r, column=6).value) or _norm(ws.cell(row=r, column=2).value)
        if a and "người lập" in a.lower():
            header["nguoi_lap"] = _norm(ws.cell(row=r, column=6).value) or _norm(ws.cell(row=r, column=2).value)
        if a and ("mục" == a.strip().lower()):
            header_row = r

    if header_row is None:
        header_row = 5  # fallback theo file mẫu

    data_start = header_row + 2  # bỏ dòng đơn vị phụ (mm/inch...)

    # Tìm dòng TOTAL để biết điểm dừng
    last_row = ws.max_row
    end_row = last_row
    for r in range(data_start, last_row + 1):
        a = _norm(ws.cell(row=r, column=1).value)
        if isinstance(a, str) and a.strip().upper() == "TOTAL":
            end_row = r
            break

    sections = []
    cur_section = None
    cur_cum = None       # node cấp Cụm hiện tại (level B)
    cur_part = None      # node cấp Part hiện tại (level C)

    for r in range(data_start, end_row):
        a = _norm(ws.cell(row=r, column=1).value)
        b = _norm(ws.cell(row=r, column=2).value)
        c = _norm(ws.cell(row=r, column=3).value)
        d = _norm(ws.cell(row=r, column=4).value)
        ma = _norm(ws.cell(row=r, column=5).value)
        mo_ta = _norm(ws.cell(row=r, column=6).value)
        length = _norm(ws.cell(row=r, column=7).value)
        width = _norm(ws.cell(row=r, column=9).value)
        thickness = _norm(ws.cell(row=r, column=13).value)
        material = _norm(ws.cell(row=r, column=15).value)
        density = _norm(ws.cell(row=r, column=16).value)
        wastage = _norm(ws.cell(row=r, column=17).value)
        qty = _norm(ws.cell(row=r, column=18).value)
        raw_w = _norm(ws.cell(row=r, column=19).value)
        fin_w = _norm(ws.cell(row=r, column=22).value)

        if a is None and b is None and c is None and d is None and ma is None and mo_ta is None:
            continue  # dòng trống

        if a is not None:
            # Dòng mở Section mới (Section 1 = sản phẩm chính; 2..n = nhóm vật tư phụ)
            cur_section = {
                "no": a, "name": mo_ta or ma, "cums": [],
            }
            sections.append(cur_section)
            cur_cum = None
            cur_part = None
            continue

        if cur_section is None:
            # phòng trường hợp file không có dòng section rõ ràng
            cur_section = {"no": "1", "name": None, "cums": []}
            sections.append(cur_section)

        if b is not None:
            node = TechNode(
                row=r, level=2, section_no=cur_section["no"], section_name=cur_section["name"],
                cum_no=b, part_no=None, detail_no=None, ma=ma, mo_ta=mo_ta,
                length=length, width=width, thickness=thickness, material=material,
                density=density, wastage=wastage, qty=qty, raw_weight=raw_w, finished_weight=fin_w,
            )
            cur_section["cums"].append(node)
            cur_cum = node
            cur_part = None
            continue

        if c is not None:
            node = TechNode(
                row=r, level=3, section_no=cur_section["no"], section_name=cur_section["name"],
                cum_no=cur_cum.cum_no if cur_cum else None, part_no=c, detail_no=None,
                ma=ma, mo_ta=mo_ta, length=length, width=width, thickness=thickness,
                material=material, density=density, wastage=wastage, qty=qty,
                raw_weight=raw_w, finished_weight=fin_w,
            )
            if cur_cum is not None:
                cur_cum.children.append(node)
            cur_part = node
            continue

        if d is not None:
            node = TechNode(
                row=r, level=4, section_no=cur_section["no"], section_name=cur_section["name"],
                cum_no=cur_cum.cum_no if cur_cum else None,
                part_no=cur_part.part_no if cur_part else None, detail_no=d,
                ma=ma, mo_ta=mo_ta, length=length, width=width, thickness=thickness,
                material=material, density=density, wastage=wastage, qty=qty,
                raw_weight=raw_w, finished_weight=fin_w,
            )
            if cur_part is not None:
                cur_part.children.append(node)
            elif cur_cum is not None:
                cur_cum.children.append(node)
            continue

    return header, sections


def _guess_process(mo_ta: str, material: str, rules: dict, is_sheet_metal: bool, section_name: str = None) -> str:
    text = _strip_accents(mo_ta or "")
    if section_name:
        for pkg in rules.get("packaging_sections", []):
            if _strip_accents(pkg) in _strip_accents(section_name):
                return "Packing"
    for kw, proc in rules.get("process_keywords", []):
        if _strip_accents(kw) in text:
            return proc
    if is_sheet_metal:
        return rules.get("default_process_sheet_metal", "CUT")
    return rules.get("default_process_purchased", "Mua")


def _code_with_prefix(ma: str, loai: str, rules: dict) -> str:
    """Ghép mã theo quy ước FG/SG/PT - KH - <code>.
    Nếu mã KHÔNG khớp danh sách drawing_prefixes -> coi là mã hàng mua, giữ nguyên."""
    if not ma:
        return ma
    ma_clean = ma.strip()
    prefix_letters = re.match(r"[A-Za-zÀ-ỹ]+", ma_clean)
    letters = prefix_letters.group(0).upper() if prefix_letters else ""
    if letters in [p.upper() for p in rules.get("drawing_prefixes", [])]:
        kh = rules.get("customer_shortcode", "STV")
        return f"{loai} - {kh} - {ma_clean}"
    return ma_clean


def _is_raw_material_row(node: TechNode) -> bool:
    if node.material and node.material.strip().lower() in RAW_MATERIALS:
        return True
    return False


def convert_to_erp(header: dict, sections: list, rules: dict = None,
                    product_code: str = None, product_name: str = None) -> list:
    """Trả về list[dict] theo đúng cột ERP_COLUMNS (dict key = tên cột)."""
    rules = rules or DEFAULT_RULES
    rows = []

    def new_row(**kw):
        r = {c: None for c in ERP_COLUMNS}
        r.update(kw)
        rows.append(r)
        return r

    # Dòng A - Thành phẩm (FG)
    du_an = header.get("du_an") or ""
    pcode = product_code or (du_an.split("_")[0].strip() if du_an else "PCODE")
    pname = product_name or (du_an.split("_", 1)[1].strip() if "_" in du_an else du_an) or "SẢN PHẨM"
    kh = rules.get("customer_shortcode", "STV")

    new_row(**{
        "STT": "A", "Mã sản phẩm": f"FG - {kh} - {pcode}",
        "Loại": "THÀNH PHẨM", "Tên sản phẩm": pname, "SL sản phẩm": 1,
    })

    stt_counter = 0

    for section in sections:
        is_main_section = (str(section.get("no")) == "1")
        for cum in section["cums"]:
            stt_counter += 1
            stt = str(stt_counter)

            if is_main_section:
                # Cụm lắp ráp chính -> dòng SG cấp cao
                new_row(**{
                    "STT": stt,
                    "Mã sản phẩm": _code_with_prefix(cum.ma, "SG", rules),
                    "Loại": "BTP",
                    "Tên sản phẩm": cum.mo_ta,
                    "SL sản phẩm": cum.qty or 1,
                    "ĐVT sp": rules.get("default_unit_piece", "Cái"),
                    "Công đoạn": "WD",
                })

                sub_idx = 0
                for part in cum.children:
                    sub_idx += 1
                    sub_stt = f"{stt}.{sub_idx}"

                    if not part.children:
                        # Leaf: chi tiết có sẵn dữ liệu vật liệu ngay trên dòng của nó
                        is_sheet = _is_raw_material_row(part)
                        new_row(**{
                            "STT": sub_stt,
                            "QTSX": "QTSX_01" if is_sheet else None,
                            "Mã sản phẩm": _code_with_prefix(part.ma, "SG", rules),
                            "Loại": "BTP",
                            "Tên sản phẩm": part.mo_ta,
                            "SL sản phẩm": part.qty or 1,
                            "ĐVT sp": rules.get("default_unit_piece", "Cái"),
                            "Tên NVL": (f"{part.material} {part.thickness}mm" if is_sheet and part.material else part.mo_ta),
                            "ĐVT (NVL)": rules.get("default_unit_weight", "Kg") if is_sheet else rules.get("default_unit_piece", "Cái"),
                            "Số lượng": part.raw_weight if is_sheet else part.qty,
                            "% Tỉ lệ hao hụt": (round((1 - float(part.wastage)) * 100, 2)
                                                if isinstance(part.wastage, (int, float)) and part.wastage <= 1 else None),
                            "Công đoạn": _guess_process(part.mo_ta, part.material, rules, is_sheet),
                        })
                    else:
                        # Nhóm chi tiết: dòng SG/PT cha không mang vật liệu, vật liệu ở các dòng con
                        new_row(**{
                            "STT": sub_stt,
                            "Mã sản phẩm": _code_with_prefix(part.ma, "SG", rules),
                            "Loại": "BTP",
                            "Tên sản phẩm": part.mo_ta,
                            "SL sản phẩm": part.qty or 1,
                            "ĐVT sp": rules.get("default_unit_piece", "Cái"),
                        })
                        op_idx = 0
                        for i, child in enumerate(part.children):
                            op_idx += 1
                            is_sheet = _is_raw_material_row(child)
                            if i == 0 and is_sheet:
                                code = _code_with_prefix(child.ma, "PT", rules)
                                if not code.upper().endswith(".01"):
                                    code = f"{code}.01" if " - " in code else f"{code}.01"
                                new_row(**{
                                    "QTSX": f"QTSX_{op_idx:02d}",
                                    "Mã sản phẩm": code,
                                    "SL sản phẩm": child.qty or 1,
                                    "ĐVT sp": rules.get("default_unit_piece", "Cái"),
                                    "Tên NVL": (f"{child.material} {child.thickness}mm" if child.material else child.mo_ta),
                                    "ĐVT (NVL)": rules.get("default_unit_weight", "Kg"),
                                    "Số lượng": child.raw_weight,
                                    "Công đoạn": _guess_process(child.mo_ta, child.material, rules, True),
                                })
                            else:
                                new_row(**{
                                    "QTSX": f"QTSX_{op_idx:02d}",
                                    "Mã sản phẩm": _code_with_prefix(part.ma, "PT", rules),
                                    "Loại": "BTP",
                                    "Tên sản phẩm": child.mo_ta,
                                    "SL sản phẩm": child.qty or 1,
                                    "ĐVT sp": rules.get("default_unit_piece", "Cái"),
                                    "Mã NVL": child.ma,
                                    "Tên NVL": child.mo_ta,
                                    "ĐVT (NVL)": rules.get("default_unit_piece", "Cái"),
                                    "Số lượng": child.qty,
                                    "Công đoạn": _guess_process(child.mo_ta, child.material, rules, False),
                                })
            else:
                # Nhóm vật tư phụ (sơn, phụ kiện, đóng gói, container) -> mỗi item 1 dòng STT riêng
                is_sheet = _is_raw_material_row(cum)
                new_row(**{
                    "STT": stt,
                    "Mã sản phẩm": _code_with_prefix(cum.ma, "PT", rules) if cum.ma else None,
                    "Loại": "BTP",
                    "Tên sản phẩm": cum.mo_ta,
                    "SL sản phẩm": cum.qty or 1,
                    "ĐVT sp": rules.get("default_unit_piece", "Cái"),
                    "Tên NVL": (f"{cum.material} {cum.thickness}mm" if is_sheet and cum.material else None),
                    "ĐVT (NVL)": rules.get("default_unit_weight", "Kg") if is_sheet else None,
                    "Số lượng": cum.raw_weight if is_sheet else None,
                    "Công đoạn": _guess_process(cum.mo_ta, cum.material, rules, is_sheet, section.get("name")),
                })
                # nếu có children (hiếm ở nhóm phụ) thì liệt kê tiếp không STT
                for child in cum.children:
                    new_row(**{
                        "Mã sản phẩm": _code_with_prefix(cum.ma, "PT", rules) if cum.ma else None,
                        "Loại": "BTP",
                        "Tên sản phẩm": child.mo_ta,
                        "SL sản phẩm": child.qty or 1,
                        "ĐVT sp": rules.get("default_unit_piece", "Cái"),
                        "Mã NVL": child.ma,
                        "Tên NVL": child.mo_ta,
                        "ĐVT (NVL)": rules.get("default_unit_piece", "Cái"),
                        "Số lượng": child.qty,
                        "Công đoạn": _guess_process(child.mo_ta, child.material, rules, False, section.get("name")),
                    })

    return rows


LEGEND_TEXT = [
    (1, "Cấu trúc mã Thành phẩm", "FG – KKK – PCODE"),
    (3, "Thành phần / Ký hiệu / Ý nghĩa", None),
    (4, "Nhóm / FG / Finished Goods", None),
    (5, "Khách hàng / KKK / shortcode KH (mã viết tắt của khách hàng)", None),
    (6, "Mã sản phẩm / PCODE / Mã sản phẩm KH hoặc mã nội bộ", None),
    (9, "Cấu trúc mã Bán thành phẩm", "SG– KKK – PCODE –P"),
]


def write_erp_workbook(rows: list, out_path: str, product_pcode: str = "PCODE",
                        customer_shortcode: str = "STV"):
    """Ghi ra file .xlsx theo đúng cấu trúc BOM ERP mẫu (chú giải mã ở đầu + bảng dữ liệu)."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    def put(r, c, v):
        ws.cell(row=r, column=c, value=v)

    put(1, 3, "Cấu trúc mã Thành phẩm"); put(1, 5, "FG – KKK – PCODE")
    put(3, 3, "Thành phần"); put(3, 4, "Ký hiệu"); put(3, 5, "Ý nghĩa")
    put(4, 3, "Nhóm"); put(4, 4, "FG"); put(4, 5, "Finished Goods")
    put(5, 3, "Khách hàng"); put(5, 4, "KKK"); put(5, 5, "shortcode KH (mã viết tắt của khách hàng)")
    put(6, 3, "Mã sản phẩm"); put(6, 4, "PCODE"); put(6, 5, "Mã sản phẩm KH hoặc mã nội bộ")

    put(9, 3, "Cấu trúc mã Bán thành phẩm"); put(9, 5, "SG– KKK – PCODE –P")
    put(11, 3, "Thành phần"); put(11, 4, "Ý nghĩa"); put(11, 5, "Quy định")
    put(12, 3, "SG"); put(12, 4, "Thành phẩm / Bán thành phẩm"); put(12, 5, "SG = thay FG")
    put(13, 3, "KKK"); put(13, 4, "shortcode KH (mã viết tắt của khách hàng)"); put(13, 5, "Theo chuẩn FG")
    put(14, 3, "PCODE"); put(14, 4, "Mã sản phẩm"); put(14, 5, "Giữ nguyên")
    put(15, 3, "P"); put(15, 4, "Công đoạn"); put(15, 5, "1 ký tự – bắt buộc, ở cuối")

    put(17, 3, "P"); put(17, 4, "Công đoạn (VN)"); put(17, 5, "English")
    codes = [("C", "Cắt", "Cutting"), ("B", "Chấn / Uốn", "Bending"), ("W", "Hàn + Mài", "Welding & Grinding"),
             ("M", "Phay/Tiện", "Machining"), ("P", "Sơn", "Painting / Coating"),
             ("K", "Vệ sinh + Đóng gói", "Cleaning & Packing"), ("H", "Hold / chờ xử lý", "Hold")]
    for i, (a, b, c) in enumerate(codes):
        put(18 + i, 3, a); put(18 + i, 4, b); put(18 + i, 5, c)

    put(26, 5, "ĐỊNH MỨC SẢN XUẤT")
    put(27, 3, product_pcode)
    put(27, 5, "Lưu ý: Nếu thông tin nào còn thiếu, bộ phận bổ sung thêm")

    header_row = 29
    for i, col_name in enumerate(ERP_COLUMNS, start=1):
        put(header_row, i, col_name)

    r = header_row + 2
    for row in rows:
        for i, col_name in enumerate(ERP_COLUMNS, start=1):
            put(r, i, row.get(col_name))
        r += 1

    for i in range(1, len(ERP_COLUMNS) + 1):
        ws.column_dimensions[get_column_letter(i)].width = 16

    wb.save(out_path)
    return out_path
