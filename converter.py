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

# ---------------------------------------------------------------------------
# QUY TRÌNH SẢN XUẤT CHUẨN (đọc trực tiếp từ file quy_trinh_san_xuat.xlsx)
# ---------------------------------------------------------------------------
# Đây là danh sách CHÍNH THỨC — Công đoạn trong file ERP xuất ra chỉ được lấy
# từ các mã dưới đây (CUT, BD, WD, MC, PNT, PK, FM, AS, PL), KHÔNG tự bịa thêm
# nhãn khác (vd "Mua", "Packing", "ĐÓNG PEM" như bản trước đây là SAI so với
# quy trình chuẩn của công ty).
DEFAULT_QTSX_TABLE = {
    "QTSX_01": {"mo_ta": "Quy trình tạo hình", "steps": [("CUT", "Cắt"), ("BD", "Chấn, uốn")]},
    "QTSX_02": {"mo_ta": "Quy trình hàn, mài, phay tiện", "steps": [("WD", "Hàn mài"), ("MC", "Phay, tiện")]},
    "QTSX_03": {"mo_ta": "Quy trình sơn", "steps": [("PNT", "Sơn")]},
    "QTSX_04": {"mo_ta": "Quy trình đóng gói", "steps": [("PK", "Vệ sinh, Đóng gói")]},
    "QTSX_05": {"mo_ta": "Quy trình cắt", "steps": [("CUT", "Cắt")]},
    "QTSX_06": {"mo_ta": "Quy trình Hàn mài", "steps": [("BD", "Hàn mài")]},  # SIC: file gốc ghi mã BD cho bước này
    "QTSX_07": {"mo_ta": "Quy trình phay tiện", "steps": [("MC", "Phay, tiện")]},
    "QTSX_08": {"mo_ta": "Quy trình dập tạo hình", "steps": [("FM", "Dập tạo hình")]},
    "QTSX_09": {"mo_ta": "Quy trình lắp ráp", "steps": [("AS", "Lắp ráp")]},
    "QTSX_10": {"mo_ta": "Quy trình mạ", "steps": [("PL", "Mạ")]},
}

# Ghi chú gốc trong file quy trình chuẩn (áp dụng khi gộp mã):
#   - Cắt, chấn/uốn: cùng 1 mã (cùng 1 QTSX)      -> QTSX_01, Công đoạn "CUT, BD"
#   - Hàn/mài, phay/tiện: cùng 1 mã                -> QTSX_02, Công đoạn "WD, MC"
#   - Sơn: 1 mã                                    -> QTSX_03, Công đoạn "PNT"
#   - Đóng gói: 1 mã                               -> QTSX_04, Công đoạn "PK"

DEFAULT_RULES = {
    # Shortcode khách hàng dùng để ghép mã: "SG - <KH> - <code>"
    "customer_shortcode": "STV",
    # Các tiền tố mã vẽ nội bộ (drawing code) sẽ được gắn "<Loại> - KH - <code>"
    # Mã KHÔNG nằm trong danh sách này (vd FA, D, PM, PR, STX, số thuần...)
    # được xem là mã hàng mua/thương mại -> giữ nguyên, không thêm tiền tố.
    "drawing_prefixes": ["GZ", "ME", "CA", "CE", "RA"],

    # Quy tắc xác định QTSX + Công đoạn theo TỪ KHOÁ trong mô tả / vật liệu.
    # Mỗi luật: [từ khoá (không dấu, để trống nếu áp dụng theo is_sheet/is_round...),
    #            mã_qtsx, công_đoạn_hiển_thị]
    # Khớp từ TRÊN XUỐNG, luật đầu tiên khớp sẽ được dùng.
    "process_keyword_rules": [
        # hàng gia công đúc: cần xác nhận tay -> để trống, không đoán bừa
        ["duc", None, None],
        # hàng tròn đặc / trục -> phay tiện
        ["tron dac", "QTSX_07", "MC"],
        ["truc", "QTSX_07", "MC"],
        # hardware lắp ráp (ép/bắt vít) vào cụm
        ["tu giu", "QTSX_09", "AS"],
        ["bulong", "QTSX_09", "AS"],
        ["vit", "QTSX_09", "AS"],
        ["dai oc", "QTSX_09", "AS"],
        ["lo xo", "QTSX_09", "AS"],
        ["chot", "QTSX_09", "AS"],
        ["long den", "QTSX_09", "AS"],
        # bao bì / nhãn / hướng dẫn / đóng gói
        ["nhan dan", "QTSX_04", "PK"],
        ["tui ", "QTSX_04", "PK"],
        ["huong dan", "QTSX_04", "PK"],
        ["bao hanh", "QTSX_04", "PK"],
        ["catalogue", "QTSX_04", "PK"],
        ["gang tay", "QTSX_04", "PK"],
        ["container", "QTSX_04", "PK"],
        ["dong goi", "QTSX_04", "PK"],
        # sơn
        ["son ", "QTSX_03", "PNT"],
        # hàng mua thuần tuý, không gia công tại xưởng (kính, gioăng, dây sợi thuỷ tinh, keo)
        ["kinh", None, None],
        ["gioang", None, None],
        ["day soi thuy tinh", None, None],
        ["keo dan", None, None],
        ["bang keo", None, None],
    ],

    # Nhóm/Section (cột A của BOM kỹ thuật) coi là đóng gói -> luôn QTSX_04 / PK
    "packaging_sections": ["VẬT TƯ ĐÓNG GÓI", "VẬT TƯ CONTAINER"],

    # Mặc định cho tấm kim loại (Steel/Inox/Aluminium... có đủ dài x rộng x dày)
    "default_qtsx_sheet_metal": "QTSX_01",
    "default_process_sheet_metal": "CUT, BD",
    # Mặc định cho cụm lắp ráp (SG có nhiều chi tiết con)
    "default_qtsx_assembly": "QTSX_09",
    "default_process_assembly": "AS",

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


def load_qtsx_table(path: str) -> dict:
    """Đọc sheet 'QTSX' trong file quy_trinh_san_xuat.xlsx để lấy danh sách
    quy trình + công đoạn CHÍNH THỨC của công ty. Nếu không đọc được, dùng
    DEFAULT_QTSX_TABLE làm dự phòng."""
    try:
        wb = openpyxl.load_workbook(path, data_only=True)
        if "QTSX" not in wb.sheetnames:
            return DEFAULT_QTSX_TABLE
        ws = wb["QTSX"]
        table = {}
        cur_ma = None
        for r in range(6, ws.max_row + 1):
            ma = _norm(ws.cell(row=r, column=2).value)
            mo_ta = _norm(ws.cell(row=r, column=3).value)
            step = _norm(ws.cell(row=r, column=4).value)
            code = _norm(ws.cell(row=r, column=5).value)
            ten = _norm(ws.cell(row=r, column=6).value)
            if ma and isinstance(ma, str) and ma.upper().startswith("QTSX"):
                cur_ma = ma.strip()
                table[cur_ma] = {"mo_ta": mo_ta, "steps": []}
            if cur_ma and code:
                table[cur_ma]["steps"].append((str(code).strip(), ten))
        return table or DEFAULT_QTSX_TABLE
    except Exception:
        return DEFAULT_QTSX_TABLE


def _guess_process(mo_ta: str, material: str, rules: dict, is_sheet_metal: bool,
                    is_assembly: bool = False, section_name: str = None):
    """Trả về tuple (qtsx_code, cong_doan_display) dựa theo quy trình CHUẨN.
    Nếu không xác định chắc chắn được -> trả (None, None) để người dùng tự điền,
    thay vì đoán bừa."""
    text = _strip_accents(mo_ta or "")

    if section_name:
        for pkg in rules.get("packaging_sections", []):
            if _strip_accents(pkg) in _strip_accents(section_name):
                return "QTSX_04", "PK"

    for rule in rules.get("process_keyword_rules", []):
        kw, qtsx, proc = rule[0], rule[1], rule[2]
        if kw and _strip_accents(kw) in text:
            return qtsx, proc

    if is_sheet_metal:
        return rules.get("default_qtsx_sheet_metal", "QTSX_01"), rules.get("default_process_sheet_metal", "CUT, BD")

    if is_assembly:
        return rules.get("default_qtsx_assembly", "QTSX_09"), rules.get("default_process_assembly", "AS")

    # Không xác định chắc chắn -> để trống, tránh gán sai công đoạn
    return None, None


def _calc_hao_hut_percent(raw_weight, finished_weight):
    """Công thức CHÍNH THỨC (theo quy_trinh_san_xuat.xlsx):
    % hao hụt = ((KL NVL đầu vào - KL thành phẩm đầu ra) / KL thành phẩm đầu ra) x 100
    Trả về None nếu thiếu dữ liệu hoặc không áp dụng được (vd hàng mua không gia công)."""
    try:
        raw_weight = float(raw_weight)
        finished_weight = float(finished_weight)
    except (TypeError, ValueError):
        return None
    if finished_weight == 0:
        return None
    return round(((raw_weight - finished_weight) / finished_weight) * 100, 2)


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
    """Trả về list[dict] theo đúng cột ERP_COLUMNS (dict key = tên cột).

    Quy tắc số liệu (theo công thức CHÍNH THỨC trong quy_trinh_san_xuat.xlsx):
      - "Số lượng" (NVL) của vật liệu cân theo Kg = TỔNG khối lượng nguyên liệu
        cần cho SL sản phẩm (raw_weight_total = raw_weight_per_unit x SL), KHÔNG
        phải khối lượng của 1 đơn vị — vì công thức chuẩn là:
            "Khối lượng NVL cần = KL SP đầu ra x (1 + % hao hụt) x Số lượng cần sản xuất"
      - "% Tỉ lệ hao hụt" = ((KL NVL đầu vào - KL thành phẩm đầu ra) / KL thành
        phẩm đầu ra) x 100, tính trực tiếp từ 2 cột khối lượng có sẵn trong BOM
        kỹ thuật (không dùng cột "TỈ LỆ HAO HỤT" mơ hồ trong file kỹ thuật).
      - "Công đoạn"/"QTSX" chỉ lấy trong bộ mã chính thức (CUT, BD, WD, MC, PNT,
        PK, FM, AS, PL). Trường hợp không đủ căn cứ để xác định chắc chắn ->
        để TRỐNG, không đoán bừa, để kỹ sư tự điền.
    """
    rules = rules or DEFAULT_RULES
    rows = []

    def new_row(**kw):
        r = {c: None for c in ERP_COLUMNS}
        r.update(kw)
        rows.append(r)
        return r

    def material_fields(node: TechNode, is_sheet: bool):
        """Trả về dict các trường Tên NVL/ĐVT (NVL)/Số lượng/% hao hụt cho 1 node vật liệu tấm."""
        if is_sheet:
            qty_total = None
            if node.raw_weight is not None and node.qty is not None:
                try:
                    qty_total = round(float(node.raw_weight) * float(node.qty), 6)
                except (TypeError, ValueError):
                    qty_total = node.raw_weight
            else:
                qty_total = node.raw_weight
            return {
                "Tên NVL": f"{node.material} {node.thickness}mm" if node.material else node.mo_ta,
                "ĐVT (NVL)": rules.get("default_unit_weight", "Kg"),
                "Số lượng": qty_total,
                "% Tỉ lệ hao hụt": _calc_hao_hut_percent(node.raw_weight, node.finished_weight),
            }
        return {
            "Tên NVL": node.mo_ta,
            "ĐVT (NVL)": rules.get("default_unit_piece", "Cái"),
            "Số lượng": node.qty,
            "% Tỉ lệ hao hụt": None,
        }

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
                # Cụm lắp ráp chính -> dòng SG cấp cao (mặc định là công đoạn Lắp ráp - AS)
                qtsx_cum, cd_cum = _guess_process(cum.mo_ta, cum.material, rules, False, is_assembly=True)
                new_row(**{
                    "STT": stt,
                    "QTSX": qtsx_cum,
                    "Mã sản phẩm": _code_with_prefix(cum.ma, "SG", rules),
                    "Loại": "BTP",
                    "Tên sản phẩm": cum.mo_ta,
                    "SL sản phẩm": cum.qty or 1,
                    "ĐVT sp": rules.get("default_unit_piece", "Cái"),
                    "Công đoạn": cd_cum,
                })

                sub_idx = 0
                for part in cum.children:
                    sub_idx += 1
                    sub_stt = f"{stt}.{sub_idx}"

                    if not part.children:
                        # Leaf: chi tiết có sẵn dữ liệu vật liệu ngay trên dòng của nó
                        is_sheet = _is_raw_material_row(part)
                        qtsx, cd = _guess_process(part.mo_ta, part.material, rules, is_sheet)
                        row = {
                            "STT": sub_stt,
                            "QTSX": qtsx,
                            "Mã sản phẩm": _code_with_prefix(part.ma, "SG" if is_sheet else "PT", rules),
                            "Loại": "BTP",
                            "Tên sản phẩm": part.mo_ta,
                            "SL sản phẩm": part.qty or 1,
                            "ĐVT sp": rules.get("default_unit_piece", "Cái"),
                            "Công đoạn": cd,
                        }
                        row.update(material_fields(part, is_sheet))
                        new_row(**row)
                    else:
                        # Nhóm chi tiết: dòng SG cha không mang vật liệu, vật liệu ở các dòng con
                        new_row(**{
                            "STT": sub_stt,
                            "Mã sản phẩm": _code_with_prefix(part.ma, "SG", rules),
                            "Loại": "BTP",
                            "Tên sản phẩm": part.mo_ta,
                            "SL sản phẩm": part.qty or 1,
                            "ĐVT sp": rules.get("default_unit_piece", "Cái"),
                        })
                        for i, child in enumerate(part.children):
                            is_sheet = _is_raw_material_row(child)
                            if i == 0 and is_sheet:
                                # Phôi cắt đầu tiên của chi tiết -> mã .01 riêng (theo mẫu ERP)
                                code = _code_with_prefix(child.ma, "PT", rules)
                                code = f"{code}.01"
                                qtsx, cd = _guess_process(child.mo_ta, child.material, rules, True)
                                row = {
                                    "QTSX": qtsx,
                                    "Mã sản phẩm": code,
                                    "SL sản phẩm": child.qty or 1,
                                    "ĐVT sp": rules.get("default_unit_piece", "Cái"),
                                    "Công đoạn": cd,
                                }
                                row.update(material_fields(child, True))
                                new_row(**row)
                            else:
                                # Hardware lắp thêm vào chi tiết (vd trụ ren tự giữ, bulong...) -> công đoạn Lắp ráp
                                qtsx, cd = _guess_process(child.mo_ta, child.material, rules, False)
                                new_row(**{
                                    "QTSX": qtsx,
                                    "Mã sản phẩm": _code_with_prefix(part.ma, "PT", rules),
                                    "Loại": "BTP",
                                    "Tên sản phẩm": child.mo_ta,
                                    "SL sản phẩm": child.qty or 1,
                                    "ĐVT sp": rules.get("default_unit_piece", "Cái"),
                                    "Mã NVL": child.ma,
                                    "Tên NVL": child.mo_ta,
                                    "ĐVT (NVL)": rules.get("default_unit_piece", "Cái"),
                                    "Số lượng": child.qty,
                                    "Công đoạn": cd,
                                })
            else:
                # Nhóm vật tư phụ (sơn, phụ kiện, đóng gói, container) -> mỗi item 1 dòng STT riêng
                is_sheet = _is_raw_material_row(cum)
                qtsx, cd = _guess_process(cum.mo_ta, cum.material, rules, is_sheet, section_name=section.get("name"))
                row = {
                    "STT": stt,
                    "QTSX": qtsx,
                    "Mã sản phẩm": _code_with_prefix(cum.ma, "PT", rules) if cum.ma else None,
                    "Loại": "BTP",
                    "Tên sản phẩm": cum.mo_ta,
                    "SL sản phẩm": cum.qty or 1,
                    "ĐVT sp": rules.get("default_unit_piece", "Cái"),
                    "Công đoạn": cd,
                }
                if is_sheet:
                    row.update(material_fields(cum, True))
                new_row(**row)
                # nếu có children (hiếm ở nhóm phụ) thì liệt kê tiếp không STT
                for child in cum.children:
                    qtsx_c, cd_c = _guess_process(child.mo_ta, child.material, rules, False, section_name=section.get("name"))
                    new_row(**{
                        "QTSX": qtsx_c,
                        "Mã sản phẩm": _code_with_prefix(cum.ma, "PT", rules) if cum.ma else None,
                        "Loại": "BTP",
                        "Tên sản phẩm": child.mo_ta,
                        "SL sản phẩm": child.qty or 1,
                        "ĐVT sp": rules.get("default_unit_piece", "Cái"),
                        "Mã NVL": child.ma,
                        "Tên NVL": child.mo_ta,
                        "ĐVT (NVL)": rules.get("default_unit_piece", "Cái"),
                        "Số lượng": child.qty,
                        "Công đoạn": cd_c,
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
