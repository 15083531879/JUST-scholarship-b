import io
import re
import tempfile
import subprocess
from pathlib import Path

import pandas as pd
import pdfplumber
import streamlit as st
from docx import Document


# ============================================================
# 页面基本设置
# ============================================================

st.set_page_config(
    page_title="江苏科技大学 · 研究生奖学金成绩B计算器",
    page_icon="🎓",
    layout="centered",
    initial_sidebar_state="collapsed",
)


# ============================================================
# 等级成绩换算
# ============================================================

GRADE_MAP = {
    "优": 95,
    "优秀": 95,
    "良": 85,
    "良好": 85,
    "中": 75,
    "及格": 65,
    "不及格": 0,
}


# ============================================================
# Logo
# ============================================================

LOGO = "static/校徽.png"


# ============================================================
# 页面样式
# ============================================================

st.markdown(
    """
<style>

.block-container {
    max-width: 900px;
    padding-top: 1.2rem;
    padding-bottom: 3rem;
    overflow: visible;
}

.school {
    font-size: 15px;
    color: #666;
    margin-bottom: 3px;
}

.title {
    font-size: 27px;
    font-weight: 800;
    line-height: 1.25;
}

.subtitle {
    font-size: 15px;
    color: #666;
    margin-top: 4px;
}

.footer {
    font-size: 12px;
    color: #8a8f95;
    text-align: center;
    margin-top: 30px;
}

div[data-testid="stMetric"] {
    background: #fafbfd;
    padding: 12px;
    border-radius: 12px;
    border: 1px solid #edf0f3;
}

</style>
""",
    unsafe_allow_html=True,
)


# ============================================================
# 页面标题
# ============================================================

if Path(LOGO).exists():
    st.image(LOGO, width=78)


st.markdown(
    """
    <div class="school">
        江苏科技大学 · 环境与化学工程学院
    </div>

    <div class="title">
        研究生奖学金成绩 B 计算器
    </div>

    <div class="subtitle">
        上传成绩单，自动识别第一学年课程并计算成绩 B
    </div>
    """,
    unsafe_allow_html=True,
)


st.markdown("<div style='height: 12px;'></div>", unsafe_allow_html=True)


# ============================================================
# 计算规则说明
# ============================================================

with st.expander("📌 计算规则", expanded=False):
    st.markdown(
        """
**成绩 B：**

> B = Σ（课程成绩 × 课程学分）÷ Σ课程学分

- 第一学年培养方案中的**学位课和非学位课**均纳入。
- 等级成绩换算：
  **优秀/优 = 95；良好/良 = 85；中 = 75；及格 = 65；不及格 = 0**。
- 系统优先按成绩单中的**第1、2学期**识别第一学年课程。
- “第一外国语（硕士英语II）”两条成绩记录分别按 **1.5 学分**计算。
- 绩点计算公式：
  **绩点 = B / 10 - 5**
- 建议使用研究生管理系统导出的 PDF 或 DOCX 成绩单。
"""
    )


# ============================================================
# 文件上传
# ============================================================

with st.container(border=True):
    uploaded = st.file_uploader(
        "📄 上传成绩单",
        type=["pdf", "docx", "doc"],
        help="支持 PDF、DOCX；DOC 格式需要部署环境安装 LibreOffice。",
    )


# ============================================================
# 基础工具函数
# ============================================================

def norm(x):
    return re.sub(r"\s+", "", str(x or "")).strip()


def grade_value(x):
    """
    将成绩转换为数值（安全范围拦截，避免将学分当作成绩）。
    """
    s = norm(x)
    if not s:
        return None

    # 1. 匹配等级成绩
    for k, v in GRADE_MAP.items():
        if k in s:
            return float(v)

    # 2. 匹配百分制数字（成绩一般 >= 30 分，或 0 分为不及格；避免 1.0~12.0 的学分被误认）
    nums = re.findall(r"\d+(?:\.\d+)?", s)
    for n in nums:
        val = float(n)
        if val == 0 or 30.0 <= val <= 100.0:
            return val

    return None


def credit_value(x):
    """
    从学分字段中提取数字（常见范围 0.5~12.0）。
    """
    s = norm(x)
    nums = re.findall(r"\d+(?:\.\d+)?", s)
    for n in nums:
        val = float(n)
        if 0.5 <= val <= 12.0:
            return val
    return None


def semester_value(x):
    """
    识别第1、2学期。
    """
    s = norm(x)
    if not s:
        return None

    if any(kw in s for kw in ["第一学期", "第1学期", "秋季"]):
        return 1
    if any(kw in s for kw in ["第二学期", "第2学期", "春季"]):
        return 2

    if "一" in s or re.search(r"(?:^|[^0-9])1(?:[^0-9]|$)", s):
        return 1
    if "二" in s or re.search(r"(?:^|[^0-9])2(?:[^0-9]|$)", s):
        return 2

    return None


def find_header_groups(row):
    """
    在一行中寻找表头标志，判断单栏/多栏。
    """
    norm_row = [norm(cell) for cell in row]
    course_cols = []
    for i, cell in enumerate(norm_row):
        if "成绩单" in cell or len(cell) > 12:
            continue
        if cell in ["课程名称", "课程名", "课程", "课程代码/名称", "课程名称(中文)"]:
            course_cols.append(i)
        elif "课程" in cell and not any(ex in cell for ex in ["代码", "类别", "性质", "类型", "属性"]):
            course_cols.append(i)

    return course_cols


def parse_sub_row(sub_row):
    """
    智能解析一行/半行单元格中的课程名、学分、成绩、学期。
    彻底防御由于列错位导致的读取失败。
    """
    cells = [norm(c) for c in sub_row if norm(c)]
    if not cells:
        return None

    # 1. 寻找成绩
    score = None
    score_idx = -1
    for idx, cell in enumerate(cells):
        if any(kw in cell for kw in ["课程", "代码", "名称", "学分", "学期", "类型", "属性"]):
            continue
        gv = grade_value(cell)
        if gv is not None:
            score = gv
            score_idx = idx
            break

    # 2. 寻找学分
    credit = None
    credit_idx = -1
    for idx, cell in enumerate(cells):
        if idx == score_idx:
            continue
        if any(kw in cell for kw in ["课程", "代码", "名称", "成绩", "学期"]):
            continue
        cv = credit_value(cell)
        if cv is not None:
            credit = cv
            credit_idx = idx
            break

    # 3. 寻找学期
    semester = None
    for idx, cell in enumerate(cells):
        if idx in (score_idx, credit_idx):
            continue
        sv = semester_value(cell)
        if sv is not None:
            semester = sv
            break

    # 4. 抽取课程名称
    possible_courses = []
    for idx, cell in enumerate(cells):
        if idx in (score_idx, credit_idx):
            continue
        if cell in ["必修", "选修", "学位课", "非学位课", "限选", "任选", "公选"]:
            continue
        if any(kw in cell for kw in ["课程名称", "课程代码", "研究生", "成绩单", "学分", "成绩", "学期", "考核"]):
            continue
        if re.match(r"^\d+$", cell):
            continue
        possible_courses.append(cell)

    if possible_courses and credit is not None and score is not None:
        course_name = max(possible_courses, key=len)
        return {
            "课程名称": course_name,
            "学分": credit,
            "原始成绩": str(cells[score_idx]),
            "换算成绩": score,
            "学期": semester,
        }

    return None


# ============================================================
# 成绩单表格解析
# ============================================================

def parse_rows(rows):
    cleaned = []
    for row in rows:
        vals = [norm(v) for v in row]
        if any(vals):
            cleaned.append(vals)

    records = []
    has_header = False
    num_groups = 1

    for row in cleaned:
        header_cols = find_header_groups(row)
        if header_cols:
            has_header = True
            num_groups = max(1, len(header_cols))
            continue

        if not has_header:
            # 未发现表头前尝试直接尝试解析数据行
            res = parse_sub_row(row)
            if res:
                records.append(res)
            continue

        num_cols = len(row)
        if num_groups > 1 and num_cols >= 6:
            # 双栏/多栏表格分段智能抽取
            mid = num_cols // num_groups
            for g_idx in range(num_groups):
                sub = row[g_idx * mid : (g_idx + 1) * mid]
                res = parse_sub_row(sub)
                if res:
                    records.append(res)
        else:
            # 单栏表格抽取
            res = parse_sub_row(row)
            if res:
                records.append(res)

    # 特殊处理第一外国语（硕士英语II）
    for record in records:
        course_name = norm(str(record["课程名称"]))
        if "第一外国语" in course_name and (
            "硕士英语II" in course_name or "硕士英语Ⅱ" in course_name or "英语II" in course_name
        ):
            record["学分"] = 1.5

    return records


# ============================================================
# DOCX 解析
# ============================================================

def parse_docx(data):
    doc = Document(io.BytesIO(data))
    rows = []
    for table in doc.tables:
        for row in table.rows:
            rows.append([cell.text for cell in row.cells])

    records = parse_rows(rows)
    if not records:
        raise ValueError("未能识别到有效课程成绩，请检查 Word 成绩单格式。")
    return records


# ============================================================
# PDF 解析与文本兜底
# ============================================================

def parse_pdf_text_fallback(data):
    records = []
    with pdfplumber.open(io.BytesIO(data)) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            for line in text.split("\n"):
                tokens = line.strip().split()
                if len(tokens) >= 3:
                    res = parse_sub_row(tokens)
                    if res:
                        records.append(res)

    for record in records:
        course_name = norm(str(record["课程名称"]))
        if "第一外国语" in course_name and (
            "硕士英语II" in course_name or "硕士英语Ⅱ" in course_name or "英语II" in course_name
        ):
            record["学分"] = 1.5

    return records


def parse_pdf(data):
    rows = []
    with pdfplumber.open(io.BytesIO(data)) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables() or []
            if not tables:
                tables = page.extract_tables(
                    table_settings={
                        "vertical_strategy": "text",
                        "horizontal_strategy": "text",
                        "snap_tolerance": 3,
                    }
                ) or []

            for table in tables:
                if table:
                    rows.extend(table)

    records = []
    if rows:
        try:
            records = parse_rows(rows)
        except Exception:
            records = []

    if not records:
        records = parse_pdf_text_fallback(data)

    if not records:
        raise ValueError("未识别到有效课程成绩，请确认上传的是清晰的研究生成绩单 PDF。")

    return records


# ============================================================
# DOC 转 DOCX
# ============================================================

def convert_doc(data):
    with tempfile.TemporaryDirectory() as td:
        src = Path(td) / "input.doc"
        src.write_bytes(data)

        p = subprocess.run(
            [
                "libreoffice",
                "--headless",
                "--convert-to",
                "docx",
                "--outdir",
                td,
                str(src),
            ],
            capture_output=True,
            text=True,
        )

        out = Path(td) / "input.docx"

        if p.returncode != 0 or not out.exists():
            raise RuntimeError("DOC 转换失败：部署环境未安装 LibreOffice。")

        return out.read_bytes()


# ============================================================
# 上传文件后的计算流程
# ============================================================

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

            # 按第 1、2 学期筛选
            if all_df["学期"].notna().any():
                calc_df = all_df[all_df["学期"].isin([1, 2])].copy()
                if len(calc_df) == 0:
                    calc_df = all_df.copy()
                    st.warning("未能可靠识别第一学年学期标识，当前按全部识别课程计算，请核对。")
                else:
                    st.success("已按第 1、2 学期识别第一学年课程。")
            else:
                calc_df = all_df.copy()
                st.warning("成绩单缺少可识别的学期信息，当前按全部课程计算，请核对。")

            calc_df["成绩×学分"] = calc_df["换算成绩"] * calc_df["学分"]
            total_credit = float(calc_df["学分"].sum())
            weighted = float(calc_df["成绩×学分"].sum())
            B = weighted / total_credit if total_credit else 0
            grade_point = B / 10 - 5

            st.markdown(
                f"### 加权成绩 B：`{B:.2f}` &nbsp;&nbsp; 绩点：`{grade_point:.3f}`"
            )

            a, b, c = st.columns(3)
            a.metric("课程数", len(calc_df))
            b.metric("总学分", f"{total_credit:.2f}")
            c.metric("加权总分", f"{weighted:.2f}")

            st.subheader("📋 识别明细")

            show = calc_df[
                ["课程名称", "学分", "原始成绩", "换算成绩", "成绩×学分"]
            ].copy()

            show["学分"] = show["学分"].map(lambda x: f"{x:.2f}")
            show["换算成绩"] = show["换算成绩"].map(lambda x: f"{x:.2f}")
            show["成绩×学分"] = show["成绩×学分"].map(lambda x: f"{x:.2f}")

            st.dataframe(show, use_container_width=True, hide_index=True)

            csv = show.to_csv(index=False).encode("utf-8-sig")
            st.download_button(
                "⬇️ 下载计算明细 CSV",
                data=csv,
                file_name="研究生奖学金成绩B计算明细.csv",
                mime="text/csv",
                use_container_width=True,
            )

            st.info("请在提交奖学金材料前对照上方列表核对各项课程学分与成绩。")

    except Exception as e:
        st.error(f"处理成绩单时发生错误：{e}")


# ============================================================
# 页脚
# ============================================================

st.markdown(
    """
    <div class="footer">
        江苏科技大学环境与化学工程学院 · 研究生奖学金成绩 B 计算器
    </div>
    """,
    unsafe_allow_html=True,
)
