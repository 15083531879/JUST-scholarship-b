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
# 等级成绩换算映射
# ============================================================

GRADE_MAP = {
    "优秀": 95, "优": 95,
    "良好": 85, "良": 85,
    "中等": 75, "中": 75,
    "及格": 65,
    "不及格": 0, "不合格": 0,
}


# ============================================================
# Logo 与 样式
# ============================================================

LOGO = "static/校徽.png"

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
# 页面标题与说明
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

with st.expander("📌 计算规则", expanded=False):
    st.markdown(
        """
**成绩 B：**

> B = Σ（课程成绩 × 课程学分）÷ Σ课程学分

- 第一学年培养方案中的**学位课和非学位课**均纳入。
- 等级成绩换算：
  **优秀/优 = 95；良好/良 = 85；中 = 75；及格 = 65；不及格 = 0**。
- 系统优先按成绩单中的**第1、2学期**识别第一学年课程。
- “第一外国语（硕士英语I/II）”两条成绩记录分别按 **1.5 学分**计算。
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


def credit_value(x):
    """提取学分（通常在 0.5 ~ 12.0 之间）"""
    s = norm(x)
    nums = re.findall(r"\d+(?:\.\d+)?", s)
    for n in nums:
        val = float(n)
        if 0.5 <= val <= 12.0:
            return val
    return None


def semester_value(x):
    """识别学期"""
    s = norm(x)
    if not s:
        return None

    if any(kw in s for kw in ["第一学期", "第1学期", "秋季", "2023-2024-1", "2024-2025-1"]):
        return 1
    if any(kw in s for kw in ["第二学期", "第2学期", "春季", "2023-2024-2", "2024-2025-2"]):
        return 2

    if "一" in s or re.search(r"(?:^|[^0-9])1(?:[^0-9]|$)", s):
        return 1
    if "二" in s or re.search(r"(?:^|[^0-9])2(?:[^0-9]|$)", s):
        return 2

    return None


def extract_single_line_record(line_str):
    """
    智能解析单行合并文本（专门攻克无框 PDF 表格）。
    利用物理顺序：学分 -> 学时 -> 成绩 -> 学期。
    成绩永远位于学时之后，取最后一个 >=30 的数字作为成绩，彻底解决学时干扰！
    """
    s = str(line_str).strip()
    if not s:
        return None

    tokens = s.split()
    if len(tokens) < 3:
        return None

    # 1. 寻找等级成绩（优/良/中/及格）
    grade_val = None
    for token in tokens:
        for k, v in GRADE_MAP.items():
            if k in token:
                grade_val = float(v)
                break
        if grade_val is not None:
            break

    # 2. 抽取学分与数值成绩
    credit = None
    valid_scores = []

    for token in tokens:
        cv = credit_value(token)
        if cv and credit is None:
            credit = cv
            continue

        nums = re.findall(r"\d+(?:\.\d+)?", token)
        for n in nums:
            v = float(n)
            if v == 0 or 30.0 <= v <= 100.0:
                valid_scores.append(v)

    # 如果有等级成绩，优先使用；没有则取最后一个有效数字（学时在前，成绩在后！）
    final_score = grade_val
    if final_score is None and valid_scores:
        final_score = valid_scores[-1]

    # 3. 抽取学期
    sem = semester_value(s)

    # 4. 提取课程名称
    course_words = []
    for t in tokens:
        if any(kw in t for kw in ["必修", "选修", "学位", "考核", "考试", "查考", "研究生", "成绩单"]):
            continue
        if re.match(r"^\d+(?:\.\d+)?$", t):
            continue
        if any(k in t for k in GRADE_MAP.keys()):
            continue
        course_words.append(t)

    course_name = norm("".join(course_words))

    if course_name and credit is not None and final_score is not None:
        return {
            "课程名称": course_name,
            "学分": credit,
            "原始成绩": str(final_score),
            "换算成绩": final_score,
            "学期": sem,
        }

    return None


def find_column_indices(rows, max_header_rows=3):
    """
    智能检测表头的列索引（支持跨行合并表头与左右多栏）。
    """
    for row_idx in range(min(len(rows), max_header_rows)):
        row = [norm(c) for c in rows[row_idx]]
        if len(row) <= 1:
            continue

        course_cols = []
        for i, cell in enumerate(row):
            if "成绩单" in cell or len(cell) > 15:
                continue
            if any(k in cell for k in ["课程名称", "课程名", "课程代码/名称", "课程中文名称"]):
                course_cols.append(i)
            elif "课程" in cell and not any(ex in cell for ex in ["代码", "类别", "性质", "类型", "属性"]):
                course_cols.append(i)

        if not course_cols:
            continue

        groups = []
        num_cols = len(row)

        for idx, ci in enumerate(course_cols):
            start = max(0, ci - 1)
            end = course_cols[idx + 1] if idx + 1 < len(course_cols) else num_cols

            cr_idx, sc_idx, se_idx = None, None, None

            # 找学分列
            for col in range(start, end):
                cell = row[col]
                if cr_idx is None and "学分" in cell and "学时" not in cell:
                    cr_idx = col

            # 优先寻找总评/综合/最终成绩列
            for col in range(start, end):
                cell = row[col]
                if any(kw in cell for kw in ["总评", "综合", "考核结果", "最终成绩", "成绩"]):
                    if not any(ex in cell for ex in ["平时", "期末", "学时", "学分", "代码", "类别"]):
                        sc_idx = col
                        break

            # 找学期列
            for col in range(start, end):
                cell = row[col]
                if se_idx is None and any(kw in cell for kw in ["学期", "开课学期", "修读学期"]):
                    se_idx = col

            if ci is not None and cr_idx is not None and sc_idx is not None:
                groups.append({
                    "course": ci,
                    "credit": cr_idx,
                    "score": sc_idx,
                    "semester": se_idx
                })

        if groups:
            return groups

    return []


# ============================================================
# 表格解析主逻辑
# ============================================================

def parse_rows(rows):
    records = []
    groups = find_column_indices(rows)

    # 1. 如果通过表头精准定位到了列索引
    if groups:
        for row in rows:
            cleaned_row = [str(c or "").replace("\n", " ").strip() for c in row]
            if not any(cleaned_row):
                continue

            for g in groups:
                c_idx, cr_idx, sc_idx, se_idx = g["course"], g["credit"], g["score"], g["semester"]
                max_req = max(c_idx, cr_idx, sc_idx, se_idx if se_idx is not None else 0)

                if len(cleaned_row) <= max_req:
                    continue

                raw_course = norm(cleaned_row[c_idx])
                raw_credit = norm(cleaned_row[cr_idx])
                raw_score = norm(cleaned_row[sc_idx])
                raw_sem = norm(cleaned_row[se_idx]) if se_idx is not None else ""

                if not raw_course or any(kw in raw_course for kw in ["课程名称", "课程代码", "研究生成绩单", "学分", "成绩"]):
                    continue

                credit = credit_value(raw_credit)
                
                # 转换成绩
                score = None
                for k, v in GRADE_MAP.items():
                    if k in raw_score:
                        score = float(v)
                        break
                if score is None:
                    nums = re.findall(r"\d+(?:\.\d+)?", raw_score)
                    for n in nums:
                        val = float(n)
                        if val == 0 or 30.0 <= val <= 100.0:
                            score = val
                            break

                semester = semester_value(raw_sem)

                if credit is not None and score is not None:
                    records.append({
                        "课程名称": raw_course,
                        "学分": credit,
                        "原始成绩": raw_score,
                        "换算成绩": score,
                        "学期": semester,
                    })

    # 2. 如果表头定位失败（如无边框 PDF 表格），采用逐行单文本智能解析
    if not records:
        for row in rows:
            line_text = " ".join([str(c or "").strip() for c in row if str(c or "").strip()])
            res = extract_single_line_record(line_text)
            if res:
                records.append(res)

    # 特殊规整：第一外国语（硕士英语I/II）按 1.5 学分计算
    for record in records:
        course_name = norm(str(record["课程名称"]))
        if "第一外国语" in course_name and any(kw in course_name for kw in ["英语I", "英语Ⅰ", "英语II", "英语Ⅱ", "硕士英语"]):
            record["学分"] = 1.5

    return records


# ============================================================
# 文件读取入口（DOCX / PDF）
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

            # 补充整页文本作为兜底
            if not tables:
                text = page.extract_text() or ""
                for line in text.split("\n"):
                    if line.strip():
                        rows.append([line.strip()])

    records = parse_rows(rows)
    if not records:
        raise ValueError("未识别到有效课程成绩，请确认上传的是清晰的研究生成绩单 PDF。")

    return records


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
# 主界面逻辑
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
