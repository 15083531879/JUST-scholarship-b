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

    .result-card {
        padding: 24px;
        border-radius: 18px;
        background: linear-gradient(
            135deg,
            #f5fbff,
            #f8fbff
        );
        border: 1px solid #d8eaf5;
        text-align: center;
        margin: 18px 0;
    }

    .result-label {
        font-size: 15px;
        color: #5f6b76;
        margin-bottom: 8px;
    }

    .result-number {
        font-size: 36px;
        font-weight: 800;
        letter-spacing: -0.5px;
    }

    .result-point {
        font-size: 20px;
        color: #5f6b76;
        margin-top: 10px;
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
    st.image(
        LOGO,
        width=78,
    )

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

st.markdown(
    "<div style='height: 12px;'></div>",
    unsafe_allow_html=True,
)


# ============================================================
# 计算规则说明
# ============================================================

with st.expander(
    "📌 计算规则",
    expanded=False,
):
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
        type=[
            "pdf",
            "docx",
            "doc",
        ],
        help="支持 PDF、DOCX；DOC 格式需要部署环境安装 LibreOffice。",
    )


# ============================================================
# 基础工具函数
# ============================================================

def norm(x):
    """
    清理字符串中的空白字符。
    """
    if x is None:
        return ""

    return re.sub(
        r"\s+",
        "",
        str(x),
    ).strip()


def grade_value(x):
    """
    将成绩转换为数值。

    例如：
    优秀 -> 95
    良好 -> 85
    90 -> 90
    89.5 -> 89.5
    """

    s = norm(x)

    if not s:
        return None

    if s in GRADE_MAP:
        return float(GRADE_MAP[s])

    m = re.search(
        r"-?\d+(?:\.\d+)?",
        s,
    )

    if m:
        return float(m.group())

    return None


def credit_value(x):
    """
    从学分字段中提取数字。
    """

    s = norm(x)

    if not s:
        return None

    m = re.search(
        r"\d+(?:\.\d+)?",
        s,
    )

    if m:
        return float(m.group())

    return None


def semester_value(x):
    """
    识别第1、2学期。
    """

    s = norm(x)

    if not s:
        return None

    if (
        "一" in s
        or re.search(
            r"(^|[^0-9])1([^0-9]|$)",
            s,
        )
    ):
        return 1

    if (
        "二" in s
        or re.search(
            r"(^|[^0-9])2([^0-9]|$)",
            s,
        )
    ):
        return 2

    return None


def find_col(headers, names):
    """
    根据表头名称寻找列位置。
    """

    for i, h in enumerate(headers):

        h = norm(h)

        if any(
            name in h
            for name in names
        ):
            return i

    return None


def is_valid_course_name(name):
    """
    判断是否可能是课程名称。
    """

    name = norm(name)

    if not name:
        return False

    invalid_words = [
        "类别",
        "课程名称",
        "课程名",
        "学时",
        "学分",
        "选修学期",
        "成绩",
        "备注",
        "培养环节",
        "总学分",
        "平均分",
        "打印时间",
        "成绩管理部门",
    ]

    if name in invalid_words:
        return False

    return len(name) >= 2


# ============================================================
# 第一外国语特殊处理
# ============================================================

def fix_special_courses(records):
    """
    特殊课程学分修正。

    江苏科技大学该类成绩单中：
    第一外国语(硕士英语II)

    可能出现：
        第一条：0学分
        第二条：3学分

    实际计算按照：
        第一条：1.5学分
        第二条：1.5学分
    """

    english_records = []

    for record in records:

        course_name = norm(
            record.get("课程名称", "")
        )

        if (
            "第一外国语" in course_name
            and (
                "硕士英语II" in course_name
                or "硕士英语Ⅱ" in course_name
            )
        ):
            english_records.append(record)

    if len(english_records) >= 2:

        for record in english_records:
            record["学分"] = 1.5

    return records


# ============================================================
# 表格解析
# ============================================================

def parse_rows(rows):
    """
    通用表格解析。

    DOCX 主要使用此方法。
    PDF 也作为第一层解析方式使用。
    """

    cleaned = []

    for row in rows:

        if row is None:
            continue

        vals = [
            norm(v)
            for v in row
        ]

        if any(vals):
            cleaned.append(vals)

    if not cleaned:
        raise ValueError(
            "未读取到成绩单表格内容。"
        )

    # --------------------------------------------------------
    # 找表头
    # --------------------------------------------------------

    header = None

    for hi, row in enumerate(
        cleaned[:30]
    ):

        ci = find_col(
            row,
            [
                "课程名称",
                "课程名",
                "课程",
            ],
        )

        cr = find_col(
            row,
            [
                "学分",
            ],
        )

        se = find_col(
            row,
            [
                "选修学期",
                "开课学期",
                "学期",
            ],
        )

        gr = find_col(
            row,
            [
                "成绩",
            ],
        )

        if (
            ci is not None
            and cr is not None
            and gr is not None
        ):
            header = (
                hi,
                ci,
                cr,
                se,
                gr,
            )
            break

    if header is None:
        raise ValueError(
            "未能识别成绩单表头。"
            "请上传研究生管理系统导出的成绩单。"
        )

    hi, ci, cr, se, gr = header

    records = []

    # --------------------------------------------------------
    # 逐行解析
    # --------------------------------------------------------

    for row in cleaned[
        hi + 1:
    ]:

        need = max(
            ci,
            cr,
            gr,
            se if se is not None else 0,
        ) + 1

        row = row + [
            ""
        ] * max(
            0,
            need - len(row),
        )

        course = row[ci]

        credit = credit_value(
            row[cr]
        )

        raw = row[gr]

        if se is not None:
            semester = semester_value(
                row[se]
            )
        else:
            semester = None

        score = grade_value(
            raw
        )

        if (
            not is_valid_course_name(course)
            or credit is None
            or score is None
        ):
            continue

        records.append(
            {
                "课程名称": course,
                "学分": credit,
                "原始成绩": raw,
                "换算成绩": score,
                "学期": semester,
            }
        )

    records = fix_special_courses(
        records
    )

    if not records:
        raise ValueError(
            "未识别到有效课程成绩。"
        )

    return records


# ============================================================
# DOCX 解析
# ============================================================

def parse_docx(data):

    doc = Document(
        io.BytesIO(data)
    )

    rows = []

    for table in doc.tables:

        for row in table.rows:

            rows.append(
                [
                    cell.text
                    for cell in row.cells
                ]
            )

    return parse_rows(rows)


# ============================================================
# PDF：第一种方法——表格解析
# ============================================================

def parse_pdf_tables(data):
    """
    使用 pdfplumber 的 extract_tables() 解析 PDF。

    这是第一层方法。
    """

    rows = []

    with pdfplumber.open(
        io.BytesIO(data)
    ) as pdf:

        for page in pdf.pages:

            tables = (
                page.extract_tables()
                or []
            )

            for table in tables:

                if not table:
                    continue

                rows.extend(
                    table
                )

    if not rows:
        return []

    try:
        return parse_rows(
            rows
        )
    except Exception:
        return []


# ============================================================
# PDF：第二种方法——文本行解析
# ============================================================

def parse_pdf_text(data):
    """
    针对江苏科技大学成绩单 PDF 的文本结构进行解析。

    由于 PDF 表格列位置可能发生错位，
    这里不依赖 extract_tables() 的列索引，
    而是从每一行文本中识别：

        课程名称
        学时
        学分
        学期
        成绩
        备注

    例如：

        精细化工工艺 32 2.0 1 93.00 正常

    """

    records = []

    with pdfplumber.open(
        io.BytesIO(data)
    ) as pdf:

        for page in pdf.pages:

            text = page.extract_text(
                x_tolerance=2,
                y_tolerance=3,
            )

            if not text:
                continue

            lines = text.splitlines()

            for line in lines:

                line = line.strip()

                if not line:
                    continue

                # ------------------------------------------------
                # 排除表头、统计信息
                # ------------------------------------------------

                if (
                    "课程名称" in line
                    or "课程名" in line
                    or "学时" in line
                    or "学分" in line
                    or "成绩" in line
                    or "培养环节名称" in line
                    or "成绩管理部门" in line
                    or "打印时间" in line
                ):
                    continue

                # ------------------------------------------------
                # 排除学生基本信息
                # ------------------------------------------------

                if (
                    "学号" in line
                    or "姓名" in line
                    or "性别" in line
                    or "学制" in line
                    or "学院" in line
                    or "专业" in line
                    or "入学日期" in line
                    or "导师姓名" in line
                ):
                    continue

                # ------------------------------------------------
                # 清理 PDF 中可能存在的多余空格
                # ------------------------------------------------

                line = re.sub(
                    r"\s+",
                    " ",
                    line,
                ).strip()

                # ------------------------------------------------
                # 关键模式
                #
                # 课程名称 + 学时 + 学分 + 学期 + 成绩
                #
                # 示例：
                #
                # 精细化工工艺 32 2.0 1 93.00 正常
                #
                # 催化理论与研究方法 32 2.0 1 良好 正常
                # ------------------------------------------------

                pattern = re.compile(
                    r"^(.*?)\s+"
                    r"(\d+(?:\.\d+)?)\s+"
                    r"(\d+(?:\.\d+)?)\s+"
                    r"([12])\s+"
                    r"(优秀|优|良好|良|中|及格|不及格|-?\d+(?:\.\d+)?)"
                    r"(?:\s+正常)?$"
                )

                m = pattern.match(
                    line
                )

                if not m:
                    continue

                course = m.group(1).strip()
                hours = m.group(2)
                credit = m.group(3)
                semester = m.group(4)
                raw_score = m.group(5)

                if not is_valid_course_name(
                    course
                ):
                    continue

                credit_num = credit_value(
                    credit
                )

                score_num = grade_value(
                    raw_score
                )

                if (
                    credit_num is None
                    or score_num is None
                ):
                    continue

                records.append(
                    {
                        "课程名称": course,
                        "学分": credit_num,
                        "原始成绩": raw_score,
                        "换算成绩": score_num,
                        "学期": int(
                            semester
                        ),
                    }
                )

    return fix_special_courses(
        records
    )


# ============================================================
# PDF：第三种方法——基于文字坐标的解析
# ============================================================

def parse_pdf_words(data):
    """
    PDF 最后备用解析方案。

    使用 extract_words() 获取文字坐标，
    根据每一行的 y 坐标重新组合文字。

    主要解决某些 PDF 中：
    extract_tables() 列错位，
    extract_text() 行结构异常的问题。
    """

    records = []

    with pdfplumber.open(
        io.BytesIO(data)
    ) as pdf:

        for page in pdf.pages:

            words = page.extract_words(
                x_tolerance=2,
                y_tolerance=3,
                keep_blank_chars=False,
            )

            if not words:
                continue

            # ------------------------------------------------
            # 按 y 坐标聚合成行
            # ------------------------------------------------

            lines = []

            for word in words:

                top = float(
                    word.get(
                        "top",
                        0,
                    )
                )

                placed = False

                for line in lines:

                    if abs(
                        line["top"] - top
                    ) <= 3:

                        line["words"].append(
                            word
                        )

                        placed = True
                        break

                if not placed:

                    lines.append(
                        {
                            "top": top,
                            "words": [
                                word
                            ],
                        }
                    )

            lines.sort(
                key=lambda x: x["top"]
            )

            # ------------------------------------------------
            # 每一行重新按照 x 排序
            # ------------------------------------------------

            for line in lines:

                ws = sorted(
                    line["words"],
                    key=lambda x: float(
                        x.get(
                            "x0",
                            0,
                        )
                    ),
                )

                text = " ".join(
                    w["text"]
                    for w in ws
                )

                text = re.sub(
                    r"\s+",
                    " ",
                    text,
                ).strip()

                if not text:
                    continue

                # 同 parse_pdf_text 的课程模式
                pattern = re.compile(
                    r"^(.*?)\s+"
                    r"(\d+(?:\.\d+)?)\s+"
                    r"(\d+(?:\.\d+)?)\s+"
                    r"([12])\s+"
                    r"(优秀|优|良好|良|中|及格|不及格|-?\d+(?:\.\d+)?)"
                    r"(?:\s+正常)?$"
                )

                m = pattern.match(
                    text
                )

                if not m:
                    continue

                course = m.group(1).strip()
                credit = m.group(3)
                semester = m.group(4)
                raw_score = m.group(5)

                if not is_valid_course_name(
                    course
                ):
                    continue

                credit_num = credit_value(
                    credit
                )

                score_num = grade_value(
                    raw_score
                )

                if (
                    credit_num is None
                    or score_num is None
                ):
                    continue

                records.append(
                    {
                        "课程名称": course,
                        "学分": credit_num,
                        "原始成绩": raw_score,
                        "换算成绩": score_num,
                        "学期": int(
                            semester
                        ),
                    }
                )

    return fix_special_courses(
        records
    )


# ============================================================
# PDF 总解析器
# ============================================================

def parse_pdf(data):
    """
    PDF 多策略解析。

    优先级：

    1. 表格解析
    2. 文本行解析
    3. 坐标文字解析

    只要成功识别出合理课程，就返回。
    """

    # --------------------------------------------------------
    # 方法 1：extract_tables
    # --------------------------------------------------------

    records = parse_pdf_tables(
        data
    )

    if records:

        # 基本质量检查
        valid_count = sum(
            1
            for r in records
            if (
                r.get("学分") is not None
                and r.get("换算成绩") is not None
                and r.get("学期") in [1, 2]
            )
        )

        if valid_count >= 3:
            return records

    # --------------------------------------------------------
    # 方法 2：extract_text
    # --------------------------------------------------------

    records = parse_pdf_text(
        data
    )

    if records:

        if len(records) >= 3:
            return records

    # --------------------------------------------------------
    # 方法 3：extract_words
    # --------------------------------------------------------

    records = parse_pdf_words(
        data
    )

    if records:
        return records

    raise ValueError(
        "PDF 成绩单解析失败。"
        "请确认 PDF 是江苏科技大学研究生管理系统导出的成绩单。"
    )


# ============================================================
# DOC 转 DOCX
# ============================================================

def convert_doc(data):

    with tempfile.TemporaryDirectory() as td:

        src = (
            Path(td)
            / "input.doc"
        )

        src.write_bytes(
            data
        )

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

        out = (
            Path(td)
            / "input.docx"
        )

        if (
            p.returncode != 0
            or not out.exists()
        ):

            raise RuntimeError(
                "DOC 转换失败："
                "部署环境未安装 LibreOffice。"
            )

        return out.read_bytes()


# ============================================================
# 检查成绩单数据
# ============================================================

def validate_records(records):
    """
    对解析结果做基本质量检查。
    """

    if not records:
        return False

    valid = 0

    for record in records:

        credit = record.get(
            "学分"
        )

        score = record.get(
            "换算成绩"
        )

        semester = record.get(
            "学期"
        )

        if (
            credit is not None
            and 0 <= credit <= 10
            and score is not None
            and 0 <= score <= 100
            and semester in [1, 2]
        ):
            valid += 1

    return valid >= 3


# ============================================================
# 上传文件后的主要计算流程
# ============================================================

if uploaded:

    try:

        data = uploaded.getvalue()

        suffix = (
            Path(uploaded.name)
            .suffix
            .lower()
        )

        # ----------------------------------------------------
        # 文件识别
        # ----------------------------------------------------

        with st.spinner(
            "正在识别成绩单并计算……"
        ):

            if suffix == ".pdf":

                records = parse_pdf(
                    data
                )

            elif suffix == ".docx":

                records = parse_docx(
                    data
                )

            elif suffix == ".doc":

                records = parse_docx(
                    convert_doc(data)
                )

            else:

                raise ValueError(
                    "不支持的文件类型。"
                )

            # ------------------------------------------------
            # 数据质量检查
            # ------------------------------------------------

            if not validate_records(
                records
            ):
                raise ValueError(
                    "成绩单解析结果异常，"
                    "请检查 PDF/DOCX 格式。"
                )

            # ------------------------------------------------
            # DataFrame
            # ------------------------------------------------

            all_df = pd.DataFrame(
                records
            )

            # ------------------------------------------------
            # 第一学年：
            # 第1、2学期
            # ------------------------------------------------

            if all_df[
                "学期"
            ].notna().any():

                calc_df = all_df[
                    all_df[
                        "学期"
                    ].isin(
                        [
                            1,
                            2,
                        ]
                    )
                ].copy()

                if len(
                    calc_df
                ) == 0:

                    calc_df = (
                        all_df.copy()
                    )

                    st.warning(
                        "未能可靠识别第一学年，"
                        "当前按全部识别课程计算，请核对。"
                    )

                else:

                    st.success(
                        "已按第 1、2 学期识别第一学年课程。"
                    )

            else:

                calc_df = (
                    all_df.copy()
                )

                st.warning(
                    "成绩单缺少可识别的学期信息，"
                    "当前按全部课程计算，请核对。"
                )

            # ------------------------------------------------
            # 成绩 × 学分
            # ------------------------------------------------

            calc_df[
                "成绩×学分"
            ] = (
                calc_df[
                    "换算成绩"
                ]
                * calc_df[
                    "学分"
                ]
            )

            # ------------------------------------------------
            # 总学分
            # ------------------------------------------------

            total_credit = float(
                calc_df[
                    "学分"
                ].sum()
            )

            # ------------------------------------------------
            # 加权总分
            # ------------------------------------------------

            weighted = float(
                calc_df[
                    "成绩×学分"
                ].sum()
            )

            # ------------------------------------------------
            # 成绩 B
            # ------------------------------------------------

            if total_credit:

                B = (
                    weighted
                    / total_credit
                )

            else:

                B = 0

            # ------------------------------------------------
            # 绩点
            # ------------------------------------------------

            grade_point = (
                B / 10 - 5
            )

            # =================================================
            # 结果显示
            # =================================================

            st.markdown(
                f"""
                <div class="result-card">

                    <div class="result-label">
                        加权成绩 B
                    </div>

                    <div class="result-number">
                        {B:.2f}
                    </div>

                    <div class="result-point">
                        绩点：{grade_point:.3f}
                    </div>

                </div>
                """,
                unsafe_allow_html=True,
            )

            # =================================================
            # 三项统计指标
            # =================================================

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

            # =================================================
            # 识别结果检查
            # =================================================

            st.subheader(
                "📋 识别明细"
            )

            show = calc_df[
                [
                    "课程名称",
                    "学期",
                    "学分",
                    "原始成绩",
                    "换算成绩",
                    "成绩×学分",
                ]
            ].copy()

            # ------------------------------------------------
            # 格式化
            # ------------------------------------------------

            show["学分"] = show[
                "学分"
            ].map(
                lambda x:
                f"{x:.2f}"
            )

            show["换算成绩"] = show[
                "换算成绩"
            ].map(
                lambda x:
                f"{x:.2f}"
            )

            show["成绩×学分"] = show[
                "成绩×学分"
            ].map(
                lambda x:
                f"{x:.2f}"
            )

            st.dataframe(
                show,
                use_container_width=True,
                hide_index=True,
            )

            # =================================================
            # PDF 解析诊断
            # =================================================

            if suffix == ".pdf":

                with st.expander(
                    "🔍 PDF 解析检查",
                    expanded=False,
                ):

                    st.write(
                        f"成功识别课程：{len(all_df)} 门"
                    )

                    st.write(
                        f"识别总学分："
                        f"{all_df['学分'].sum():.2f}"
                    )

                    st.write(
                        "如果这里的课程、学分、成绩均正确，"
                        "则计算结果可以直接使用。"
                    )

                    st.dataframe(
                        all_df[
                            [
                                "课程名称",
                                "学分",
                                "原始成绩",
                                "换算成绩",
                                "学期",
                            ]
                        ],
                        use_container_width=True,
                        hide_index=True,
                    )

            # =================================================
            # 下载计算明细
            # =================================================

            csv = (
                show
                .to_csv(
                    index=False
                )
                .encode(
                    "utf-8-sig"
                )
            )

            st.download_button(
                "⬇️ 下载计算明细",
                data=csv,
                file_name=(
                    "研究生奖学金成绩B计算明细.csv"
                ),
                mime="text/csv",
                use_container_width=True,
            )

            # =================================================
            # 提醒
            # =================================================

            st.info(
                "请在提交奖学金材料前核对课程、"
                "学分及成绩识别结果。"
            )

    except Exception as e:

        st.error(
            f"处理成绩单时发生错误：{e}"
        )


# ============================================================
# 页脚
# ============================================================

st.markdown(
    """
    <div class="footer">
        江苏科技大学环境与化学工程学院 ·
        研究生奖学金成绩 B 计算器
    </div>
    """,
    unsafe_allow_html=True,
)
