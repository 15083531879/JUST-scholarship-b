import io
import re
import tempfile
import subprocess
from pathlib import Path

import pandas as pd
import pdfplumber
import streamlit as st
from docx import Document

st.set_page_config(
    page_title="江苏科技大学 · 研究生奖学金成绩B计算器",
    page_icon="🎓",
    layout="centered",
    initial_sidebar_state="collapsed",
)

GRADE_MAP = {
    "优": 95, "优秀": 95,
    "良": 85, "良好": 85,
    "中": 75,
    "及格": 65,
    "不及格": 0,
}

LOGO = "static/校徽.png"

st.markdown("""
<style>
.block-container {max-width: 900px; padding-top: 1.2rem; padding-bottom: 3rem;}
.header {display:flex; align-items:center; gap:18px; padding:8px 0 18px;}
.header img {width:78px; height:78px; object-fit:contain;}
.school {font-size:15px; color:#666; margin-bottom:3px;}
.title {font-size:27px; font-weight:800; line-height:1.25;}
.subtitle {font-size:15px; color:#666; margin-top:4px;}
.upload-card {padding:20px; border:1px solid #e6eaf0; border-radius:16px; background:#fff;}
.result-card {padding:24px; border-radius:18px; background:linear-gradient(135deg,#f5fbff,#f8fbff); border:1px solid #d8eaf5; text-align:center; margin:18px 0;}
.result-label {font-size:15px; color:#5f6b76;}
.result-number {font-size:56px; font-weight:850; letter-spacing:-1px;}
.footer {font-size:12px; color:#8a8f95; text-align:center; margin-top:30px;}
div[data-testid="stMetric"] {background:#fafbfd; padding:12px; border-radius:12px; border:1px solid #edf0f3;}
</style>
""", unsafe_allow_html=True)

if Path(LOGO).exists():
    st.markdown(
        f"""
        <div class="header">
          <img src="{LOGO}">
          <div>
            <div class="school">江苏科技大学 · 环境与化学工程学院</div>
            <div class="title">研究生奖学金成绩 B 计算器</div>
            <div class="subtitle">上传成绩单，自动识别第一学年课程并计算成绩 B</div>
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

with st.expander("📌 计算规则", expanded=False):
    st.markdown("""
**成绩 B：**

> B = Σ（课程成绩 × 课程学分）÷ Σ课程学分

- 第一学年培养方案中的**学位课和非学位课**均纳入。
- 等级成绩换算：**优秀/优 = 95；良好/良 = 85；中 = 75；及格 = 65；不及格 = 0**。
- 系统优先按成绩单中的**第1、2学期**识别第一学年课程。
- 建议使用研究生管理系统导出的 PDF 或 DOCX 成绩单。
""")

st.markdown('<div class="upload-card">', unsafe_allow_html=True)
uploaded = st.file_uploader(
    "📄 上传成绩单",
    type=["pdf", "docx", "doc"],
    help="支持 PDF、DOCX；DOC 格式需要部署环境安装 LibreOffice。",
)
st.markdown('</div>', unsafe_allow_html=True)

def norm(x):
    return re.sub(r"\s+", "", str(x or "")).strip()

def grade_value(x):
    s = norm(x)
    if not s:
        return None
    if s in GRADE_MAP:
        return float(GRADE_MAP[s])
    m = re.search(r"-?\d+(?:\.\d+)?", s)
    return float(m.group()) if m else None

def credit_value(x):
    m = re.search(r"\d+(?:\.\d+)?", norm(x))
    return float(m.group()) if m else None

def semester_value(x):
    s = norm(x)
    if "一" in s or re.search(r"(^|[^0-9])1([^0-9]|$)", s):
        return 1
    if "二" in s or re.search(r"(^|[^0-9])2([^0-9]|$)", s):
        return 2
    return None

def find_col(headers, names):
    for i, h in enumerate(headers):
        h = norm(h)
        if any(name in h for name in names):
            return i
    return None

def parse_rows(rows):
    cleaned = []
    for row in rows:
        vals = [norm(v) for v in row]
        if any(vals):
            cleaned.append(vals)

    header = None
    for hi, row in enumerate(cleaned[:20]):
        ci = find_col(row, ["课程名称", "课程名", "课程"])
        cr = find_col(row, ["学分"])
        se = find_col(row, ["选修学期", "开课学期", "学期"])
        gr = find_col(row, ["成绩"])
        if ci is not None and cr is not None and gr is not None:
            header = (hi, ci, cr, se, gr)
            break

    if header is None:
        raise ValueError("未能识别成绩单表头。请上传研究生管理系统导出的成绩单。")

    hi, ci, cr, se, gr = header
    records = []

    for row in cleaned[hi + 1:]:
        need = max(ci, cr, gr, se if se is not None else 0) + 1
        row = row + [""] * max(0, need - len(row))

        course = row[ci]
        credit = credit_value(row[cr])
        raw = row[gr]
        semester = semester_value(row[se]) if se is not None else None
        score = grade_value(raw)

        if not course or credit is None or score is None:
            continue

        records.append({
            "课程名称": course,
            "学分": credit,
            "原始成绩": raw,
            "换算成绩": score,
            "学期": semester,
        })

    if not records:
        raise ValueError("未识别到有效课程成绩。")
    return records

def parse_docx(data):
    doc = Document(io.BytesIO(data))
    rows = []
    for table in doc.tables:
        for row in table.rows:
            rows.append([cell.text for cell in row.cells])
    return parse_rows(rows)

def parse_pdf(data):
    rows = []
    with pdfplumber.open(io.BytesIO(data)) as pdf:
        for page in pdf.pages:
            for table in page.extract_tables() or []:
                rows.extend(table)
    return parse_rows(rows)

def convert_doc(data):
    with tempfile.TemporaryDirectory() as td:
        src = Path(td) / "input.doc"
        src.write_bytes(data)
        p = subprocess.run(
            ["libreoffice", "--headless", "--convert-to", "docx",
             "--outdir", td, str(src)],
            capture_output=True, text=True
        )
        out = Path(td) / "input.docx"
        if p.returncode != 0 or not out.exists():
            raise RuntimeError("DOC 转换失败：部署环境未安装 LibreOffice。")
        return out.read_bytes()

if uploaded:
    try:
        data = uploaded.getvalue()
        suffix = Path(uploaded.name).suffix.lower()

        with st.spinner("正在识别成绩单并计算……"):
            if suffix == ".pdf":
                records = parse_pdf(data)
            elif suffix == ".docx":
                records = parse_docx(data)
            elif suffix == ".doc":
                records = parse_docx(convert_doc(data))
            else:
                raise ValueError("不支持的文件类型。")

            all_df = pd.DataFrame(records)

            if all_df["学期"].notna().any():
                calc_df = all_df[all_df["学期"].isin([1, 2])].copy()

                if len(calc_df) == 0:
                    calc_df = all_df.copy()
                    st.warning(
                        "未能可靠识别第一学年，当前按全部识别课程计算，请核对。"
                    )
                else:
                    st.success(
                        "已按第 1、2 学期识别第一学年课程。"
                    )
            else:
                calc_df = all_df.copy()
                st.warning(
                    "成绩单缺少可识别的学期信息，当前按全部课程计算，请核对。"
                )

            calc_df["成绩×学分"] = (
                calc_df["换算成绩"] * calc_df["学分"]
            )

            total_credit = float(calc_df["学分"].sum())
            weighted = float(calc_df["成绩×学分"].sum())

            B = (
                weighted / total_credit
                if total_credit
                else 0
            )

            st.markdown(
                f"""
                <div class="result-card">
                  <div class="result-label">本次计算结果 · 成绩 B</div>
                  <div class="result-number">{B:.2f}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )

            a, b, c = st.columns(3)

            a.metric(
                "课程数",
                len(calc_df),
            )

            b.metric(
                "总学分",
                f"{total_credit:.2f}",
            )

            c.metric(
                "加权总分",
                f"{weighted:.2f}",
            )

            st.subheader("📋 识别明细")

            show = calc_df[
                [
                    "课程名称",
                    "学分",
                    "原始成绩",
                    "换算成绩",
                    "成绩×学分",
                ]
            ].copy()

            st.dataframe(
                show,
                use_container_width=True,
                hide_index=True,
            )

            csv = show.to_csv(
                index=False
            ).encode("utf-8-sig")

            st.download_button(
                "⬇️ 下载计算明细",
                data=csv,
                file_name="研究生奖学金成绩B计算明细.csv",
                mime="text/csv",
                use_container_width=True,
            )

            st.info(
                "请在提交奖学金材料前核对课程、学分及成绩识别结果。"
            )

    except Exception as e:
        st.error(
            f"处理成绩单时发生错误：{e}"
        )


st.markdown(
    '<div class="footer">江苏科技大学环境与化学工程学院 · 研究生奖学金成绩 B 计算器</div>',
    unsafe_allow_html=True,
)
