# -*- coding: utf-8 -*-
import io
import json
import copy

import streamlit as st
import pandas as pd

from converter import (
    load_technical_bom, convert_to_erp, write_erp_workbook,
    DEFAULT_RULES, ERP_COLUMNS,
)

st.set_page_config(page_title="Chuyển đổi BOM Kỹ thuật -> ERP", layout="wide")

st.title("🔧 Công cụ chuyển đổi BOM Kỹ thuật ➜ BOM ERP")
st.caption(
    "Đọc file BOM kỹ thuật (sheet 'Form', cấu trúc cây Cụm > Chi tiết > Vật tư) "
    "và tự động dựng bảng BOM theo đúng cột dùng cho ERP. "
    "Kết quả là bản NHÁP — vẫn cần kỹ sư rà soát trước khi nạp vào hệ thống thật."
)

with st.expander("ℹ️ Công cụ hoạt động như thế nào? (đọc trước khi dùng)"):
    st.markdown(
        """
- File BOM **kỹ thuật** lưu thông tin **thiết kế**: kích thước, vật liệu, khối lượng.
- File BOM **ERP** cần thêm thông tin **sản xuất**: QTSX (quy trình sản xuất), Công đoạn,
  và đôi khi 1 chi tiết được tách thành nhiều bước (cắt → đóng PEM → hàn...).
- Thông tin sản xuất này **không có sẵn** trong BOM kỹ thuật, nên công cụ chỉ **đoán**
  công đoạn theo từ khoá trong mô tả (có thể chỉnh sửa danh sách từ khoá ở thanh bên).
- Tên chi tiết (Tên sản phẩm) ở một số dòng có thể bị thiếu vì file kỹ thuật gốc
  không luôn ghi tên tiếng Anh của chi tiết — bạn có thể bổ sung trực tiếp trong
  bảng xem trước bên dưới trước khi tải về.
        """
    )

# ---------------------------------------------------------------
# Sidebar: cấu hình / quy tắc chuyển đổi
# ---------------------------------------------------------------
st.sidebar.header("⚙️ Quy tắc chuyển đổi")

rules = copy.deepcopy(DEFAULT_RULES)

rules["customer_shortcode"] = st.sidebar.text_input(
    "Shortcode khách hàng (KKK)", value=DEFAULT_RULES["customer_shortcode"]
)

drawing_prefixes_str = st.sidebar.text_input(
    "Tiền tố mã vẽ nội bộ (cách nhau bởi dấu phẩy) — các mã này sẽ được gắn 'FG/SG/PT - KH -'",
    value=", ".join(DEFAULT_RULES["drawing_prefixes"]),
)
rules["drawing_prefixes"] = [x.strip().upper() for x in drawing_prefixes_str.split(",") if x.strip()]

with st.sidebar.expander("Từ khoá đoán Công đoạn"):
    kw_df = pd.DataFrame(DEFAULT_RULES["process_keywords"], columns=["Từ khoá (không dấu)", "Công đoạn"])
    kw_df = st.data_editor(kw_df, num_rows="dynamic", use_container_width=True, key="kw_editor")
    rules["process_keywords"] = [[str(r["Từ khoá (không dấu)"]), str(r["Công đoạn"])]
                                  for _, r in kw_df.iterrows() if r["Từ khoá (không dấu)"]]

st.sidebar.markdown("---")
st.sidebar.caption("Đổi quy tắc sẽ áp dụng lại ngay khi bạn bấm 'Chuyển đổi'.")

# ---------------------------------------------------------------
# Upload file
# ---------------------------------------------------------------
uploaded = st.file_uploader("📤 Tải lên file BOM kỹ thuật (.xlsx)", type=["xlsx"])

if uploaded is not None:
    try:
        header, sections = load_technical_bom(uploaded)
    except Exception as e:
        st.error(f"Không đọc được file. Lỗi: {e}")
        st.stop()

    st.subheader("1️⃣ Thông tin nhận diện từ file")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Khách hàng", header.get("khach_hang") or "?")
    c2.metric("Dự án", header.get("du_an") or "?")
    c3.metric("Revision", str(header.get("revision") or "?"))
    c4.metric("Người lập BOM", header.get("nguoi_lap") or "?")

    du_an = header.get("du_an") or ""
    default_pcode = du_an.split("_")[0].strip() if du_an else ""
    default_pname = du_an.split("_", 1)[1].strip() if "_" in du_an else du_an

    colA, colB = st.columns(2)
    product_code = colA.text_input("Mã sản phẩm (PCODE) dùng cho FG", value=default_pcode)
    product_name = colB.text_input("Tên thành phẩm (FG)", value=default_pname)

    st.write(
        f"Đã phát hiện **{sum(len(s['cums']) for s in sections)}** cụm/chi tiết/vật tư "
        f"trong **{len(sections)}** nhóm: " +
        ", ".join(f"{s['name']} ({len(s['cums'])})" for s in sections)
    )

    if st.button("🚀 Chuyển đổi sang BOM ERP", type="primary"):
        rows = convert_to_erp(
            header, sections, rules,
            product_code=product_code or None,
            product_name=product_name or None,
        )
        st.session_state["erp_rows"] = rows

    if "erp_rows" in st.session_state:
        st.subheader("2️⃣ Xem trước & chỉnh sửa BOM ERP")
        st.caption(
            "Bạn có thể sửa trực tiếp trong bảng (đổi tên, mã, công đoạn, gộp dòng...) "
            "trước khi tải về. Ô trống có thể bổ sung thủ công."
        )
        df = pd.DataFrame(st.session_state["erp_rows"], columns=ERP_COLUMNS)
        edited_df = st.data_editor(
            df, num_rows="dynamic", use_container_width=True, height=600, key="erp_editor"
        )

        st.subheader("3️⃣ Tải file BOM ERP (.xlsx)")
        out_buf = io.BytesIO()
        rows_out = edited_df.where(pd.notnull(edited_df), None).to_dict("records")
        write_erp_workbook(
            rows_out, "___tmp_erp_output.xlsx",
            product_pcode=product_code or "PCODE",
            customer_shortcode=rules["customer_shortcode"],
        )
        with open("___tmp_erp_output.xlsx", "rb") as f:
            out_buf.write(f.read())
        out_buf.seek(0)

        st.download_button(
            "⬇️ Tải file BOM ERP",
            data=out_buf,
            file_name=f"BOM_ERP_{product_code or 'output'}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
else:
    st.info("Vui lòng tải lên file BOM kỹ thuật (.xlsx) để bắt đầu.")
