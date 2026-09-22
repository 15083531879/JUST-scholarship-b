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

    s = str(x)

    # PDF 中可能存在换行、不可见空格
    s = s.replace("\xa0", " ")
    s = s.replace("\u3000", " ")

    return re.sub(
        r"\s+",
        "",
        s,
    ).strip()


def grade_value(x):
    """
    将成绩转换为数值。

    支持：

    优秀 -> 95
    良好 -> 85
    90 -> 90
    89.5 -> 89.5
    85/良好 -> 85
    85.0(优) -> 85
    """

    s = norm(x)

    if not s:
        return None

    # 先匹配等级
    for grade, value in GRADE_MAP.items():

        if grade in s:

            # 如果同时存在数字成绩，
            # 优先使用数字成绩
            m = re.search(
                r"-?\d+(?:\.\d+)?",
                s,
            )

            if m:
                return float(
                    m.group()
                )

            return float(value)

    # 再匹配数字
    m = re.search(
        r"-?\d+(?:\.\d+)?",
        s,
    )

    if m:

        return float(
            m.group()
        )

    return None


def credit_value(x):
    """
    从学分字段中提取数字。

    支持：

    3
    3.0
    3.0学分
    学分3
    """

    s = norm(x)

    if not s:
        return None

    m = re.search(
        r"\d+(?:\.\d+)?",
        s,
    )

    if m:

        return float(
            m.group()
        )

    return None


def semester_value(x):
    """
    识别第1、2学期。
    """

    s = norm(x)

    if not s:
        return None

    # 中文学期
    if "第一学期" in s or "第一" in s:
        return 1

    if "第二学期" in s or "第二" in s:
        return 2

    # 数字学期
    if re.search(
        r"(^|[^0-9])1([^0-9]|$)",
        s,
    ):
        return 1

    if re.search(
        r"(^|[^0-9])2([^0-9]|$)",
        s,
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


# ============================================================
# 特殊课程处理
# ============================================================

def fix_special_courses(records):
    """
    修正特殊课程：

    第一外国语（硕士英语II）
    第一外国语（硕士英语Ⅱ）
    第一外国语（硕士英语2）

    两条记录分别按 1.5 学分计算。
    """

    for record in records:

        course_name = norm(
            record.get(
                "课程名称",
                "",
            )
        )

        if (
            "第一外国语" in course_name
            and (
                "硕士英语II" in course_name
                or "硕士英语Ⅱ" in course_name
                or "硕士英语2" in course_name
            )
        ):

            record["学分"] = 1.5

    return records


# ============================================================
# 通用表格解析
# ============================================================

def parse_rows(rows):
    """
    将 Word / PDF 提取出来的普通二维表格
    统一解析为课程记录。
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

            cleaned.append(
                vals
            )

    if not cleaned:

        raise ValueError(
            "未读取到成绩单表格内容。"
        )

    # --------------------------------------------------------
    # 寻找表头
    # --------------------------------------------------------

    header = None

    # 不仅搜索前20行。
    # PDF 多页时可能在后面才出现有效表头。
    search_limit = min(
        len(cleaned),
        100,
    )

    for hi in range(
        search_limit
    ):

        row = cleaned[hi]

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
                "上课学期",
                "学期",
            ],
        )

        gr = find_col(
            row,
            [
                "成绩",
                "总评成绩",
                "课程成绩",
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

        if len(row) < need:

            row = row + [
                ""
            ] * (
                need - len(row)
            )

        course = row[ci]

        credit = credit_value(
            row[cr]
        )

        raw = row[gr]

        semester = None

        if se is not None:

            semester = semester_value(
                row[se]
            )

        score = grade_value(
            raw
        )

        # ----------------------------------------------------
        # 无效记录跳过
        # ----------------------------------------------------

        if (
            not course
            or credit is None
            or score is None
        ):

            continue

        # 排除明显的表头、合计行
        if (
            "课程名称" in course
            or "课程名" == course
            or "合计" in course
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

    return parse_rows(
        rows
    )


# ============================================================
# PDF：寻找表头
# ============================================================

def pdf_header_info(table):
    """
    在 PDF 表格中寻找表头。

    返回：

    {
        "header_row": 行号,
        "course_cols": [...],
        "credit_cols": [...],
        "grade_cols": [...],
        "semester_cols": [...]
    }

    PDF 某些情况下会出现重复表头，
    因此这里不只寻找第一个列。
    """

    if not table:
        return None

    for row_index, row in enumerate(
        table[:30]
    ):

        if not row:
            continue

        cells = [
            norm(v)
            for v in row
        ]

        course_cols = []
        credit_cols = []
        grade_cols = []
        semester_cols = []

        for i, cell in enumerate(
            cells
        ):

            if not cell:
                continue

            if (
                "课程名称" in cell
                or "课程名" in cell
                or cell == "课程"
            ):
                course_cols.append(i)

            if "学分" in cell:
                credit_cols.append(i)

            if (
                "成绩" in cell
                or "总评成绩" in cell
                or "课程成绩" in cell
            ):
                grade_cols.append(i)

            if (
                "选修学期" in cell
                or "开课学期" in cell
                or "上课学期" in cell
                or cell == "学期"
            ):
                semester_cols.append(i)

        if (
            course_cols
            and credit_cols
            and grade_cols
        ):

            return {
                "header_row": row_index,
                "course_cols": course_cols,
                "credit_cols": credit_cols,
                "grade_cols": grade_cols,
                "semester_cols": semester_cols,
            }

    return None


# ============================================================
# PDF：单个表格解析
# ============================================================

def parse_pdf_table(table):
    """
    解析 pdfplumber 返回的单个表格。

    重点处理 PDF 中可能出现的：
    - 空白列
    - 重复表头
    - 左右两组课程列
    - 多页重复表头
    """

    info = pdf_header_info(
        table
    )

    if info is None:
        return []

    header_row = info[
        "header_row"
    ]

    course_cols = info[
        "course_cols"
    ]

    credit_cols = info[
        "credit_cols"
    ]

    grade_cols = info[
        "grade_cols"
    ]

    semester_cols = info[
        "semester_cols"
    ]

    records = []

    # --------------------------------------------------------
    # 正常情况下只有一组列
    #
    # 如果 PDF 出现：
    #
    # 课程名称 ... 学分 ... 成绩 ... 课程名称 ... 学分 ... 成绩
    #
    # 那么 course_cols / credit_cols / grade_cols
    # 会出现多个位置。
    #
    # 这里按课程名称分别处理。
    # --------------------------------------------------------

    for course_index, course_col in enumerate(
        course_cols
    ):

        # ----------------------------------------------------
        # 找距离当前课程名称最近的“学分”列
        # ----------------------------------------------------

        credit_candidates = [
            x
            for x in credit_cols
            if x >= course_col
        ]

        if credit_candidates:

            credit_col = min(
                credit_candidates,
                key=lambda x: abs(
                    x - course_col
                ),
            )

        else:

            credit_col = min(
                credit_cols,
                key=lambda x: abs(
                    x - course_col
                ),
            )

        # ----------------------------------------------------
        # 找距离当前课程名称最近的“成绩”列
        # ----------------------------------------------------

        grade_candidates = [
            x
            for x in grade_cols
            if x >= course_col
        ]

        if grade_candidates:

            grade_col = min(
                grade_candidates,
                key=lambda x: abs(
                    x - course_col
                ),
            )

        else:

            grade_col = min(
                grade_cols,
                key=lambda x: abs(
                    x - course_col
                ),
            )

        # ----------------------------------------------------
        # 找当前课程组对应的学期列
        # ----------------------------------------------------

        semester_col = None

        if semester_cols:

            semester_candidates = [
                x
                for x in semester_cols
                if (
                    min(
                        course_col,
                        grade_col,
                    )
                    <= x
                    <= max(
                        course_col,
                        grade_col,
                    )
                )
            ]

            if semester_candidates:

                semester_col = min(
                    semester_candidates,
                    key=lambda x: abs(
                        x - course_col
                    ),
                )

            else:

                semester_col = min(
                    semester_cols,
                    key=lambda x: abs(
                        x - course_col
                    ),
                )

        # ----------------------------------------------------
        # 逐行读取
        # ----------------------------------------------------

        for row in table[
            header_row + 1:
        ]:

            if not row:
                continue

            row = list(row)

            max_index = max(
                course_col,
                credit_col,
                grade_col,
                semester_col
                if semester_col is not None
                else 0,
            )

            if len(row) <= max_index:

                row += [
                    ""
                ] * (
                    max_index
                    + 1
                    - len(row)
                )

            course = norm(
                row[course_col]
            )

            credit = credit_value(
                row[credit_col]
            )

            raw = norm(
                row[grade_col]
            )

            semester = None

            if semester_col is not None:

                semester = semester_value(
                    row[semester_col]
                )

            score = grade_value(
                raw
            )

            if (
                not course
                or credit is None
                or score is None
            ):
                continue

            # 排除重复表头
            if (
                "课程名称" in course
                or "课程名" == course
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

    return records


# ============================================================
# PDF：普通表格解析
# ============================================================

def parse_pdf_by_tables(data):
    """
    优先使用 pdfplumber 的表格识别。

    不直接把所有 table 拼起来，
    而是逐个表格识别自己的表头和列。
    """

    records = []

    with pdfplumber.open(
        io.BytesIO(data)
    ) as pdf:

        for page in pdf.pages:

            tables = (
                page.extract_tables(
                    table_settings={
                        "vertical_strategy": "lines",
                        "horizontal_strategy": "lines",
                        "intersection_tolerance": 5,
                        "snap_tolerance": 3,
                        "join_tolerance": 3,
                    }
                )
                or []
            )

            for table in tables:

                table_records = parse_pdf_table(
                    table
                )

                records.extend(
                    table_records
                )

    return fix_special_courses(
        records
    )


# ============================================================
# PDF：备用文本解析
# ============================================================

def parse_pdf_by_text(data):
    """
    当 PDF 表格边框识别失败时，
    尝试直接读取 PDF 文本。

    适用于“肉眼看起来是表格，
    但 PDF 本身没有真正的表格边框”的情况。
    """

    all_lines = []

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

            for line in text.splitlines():

                line = line.strip()

                if line:
                    all_lines.append(
                        line
                    )

    if not all_lines:
        return []

    records = []

    # --------------------------------------------------------
    # 尝试寻找包含课程、学分、成绩的表头
    # --------------------------------------------------------

    header_index = None

    for i, line in enumerate(
        all_lines
    ):

        s = norm(line)

        if (
            "课程名称" in s
            and "学分" in s
            and "成绩" in s
        ):

            header_index = i
            break

    # --------------------------------------------------------
    # 如果找到了表头，解析后续文本
    # --------------------------------------------------------

    start = (
        header_index + 1
        if header_index is not None
        else 0
    )

    for line in all_lines[
        start:
    ]:

        clean = line.strip()

        if not clean:
            continue

        # 排除表头
        if (
            "课程名称" in clean
            or "课程名" in clean
        ):
            continue

        # 排除明显非课程内容
        if (
            "学号" in clean
            or "姓名" in clean
            or "学院" in clean
            or "专业" in clean
        ):
            continue

        # ----------------------------------------------------
        # 查找成绩
        #
        # 常见形式：
        # 85
        # 85.0
        # 优
        # 良好
        # ----------------------------------------------------

        score_match = re.search(
            r"(?<![\d.])"
            r"(?:100(?:\.0+)?|"
            r"(?:[0-9]{1,2})(?:\.[0-9]+)?)"
            r"(?![\d.])",
            clean,
        )

        grade_match = None

        if score_match is None:

            for grade in GRADE_MAP:

                if grade in clean:

                    grade_match = grade
                    break

        if (
            score_match is None
            and grade_match is None
        ):
            continue

        # ----------------------------------------------------
        # 查找学分
        #
        # 由于课程名称本身可能包含数字，
        # 因此优先查找“学分”附近数字。
        # ----------------------------------------------------

        credit_match = re.search(
            r"(\d+(?:\.\d+)?)\s*学分",
            clean,
        )

        if credit_match is None:

            # 尝试找常见的 0.5 / 1 / 1.5 / 2 / 3 / 4
            credit_matches = list(
                re.finditer(
                    r"(?<![\d.])"
                    r"\d+(?:\.\d+)?"
                    r"(?![\d.])",
                    clean,
                )
            )

            if not credit_matches:
                continue

            # 一般最后面的数字更可能是学分
            credit_match = credit_matches[-1]

        credit = float(
            credit_match.group(1)
            if credit_match.lastindex
            else credit_match.group()
        )

        # ----------------------------------------------------
        # 成绩
        # ----------------------------------------------------

        if score_match is not None:

            score = float(
                score_match.group()
            )

            raw_score = score_match.group()

        else:

            score = float(
                GRADE_MAP[
                    grade_match
                ]
            )

            raw_score = grade_match

        # ----------------------------------------------------
        # 课程名称
        #
        # 简单策略：
        # 去掉成绩、学分及明显字段后，
        # 保留前面的文字。
        # ----------------------------------------------------

        course = clean

        if credit_match is not None:

            course = course[
                :credit_match.start()
            ]

        if score_match is not None:

            course = course.replace(
                score_match.group(),
                "",
            )

        for grade in GRADE_MAP:

            course = course.replace(
                grade,
                "",
            )

        course = re.sub(
            r"\s+",
            "",
            course,
        ).strip()

        # 去除开头可能出现的序号
        course = re.sub(
            r"^\d+[、.)．]?",
            "",
            course,
        )

        if not course:
            continue

        # ----------------------------------------------------
        # 学期
        # ----------------------------------------------------

        semester = semester_value(
            clean
        )

        records.append(
            {
                "课程名称": course,
                "学分": credit,
                "原始成绩": raw_score,
                "换算成绩": score,
                "学期": semester,
            }
        )

    return fix_special_courses(
        records
    )


# ============================================================
# PDF 总解析函数
# ============================================================

def parse_pdf(data):
    """
    PDF 解析采用两级策略：

    第一优先：
        表格解析

    第二备用：
        文本行解析

    最终统一返回标准课程记录。
    """

    # --------------------------------------------------------
    # 第一种：PDF 表格
    # --------------------------------------------------------

    try:

        records = parse_pdf_by_tables(
            data
        )

        if records:

            return records

    except Exception:
        pass

    # --------------------------------------------------------
    # 第二种：PDF 文本
    # --------------------------------------------------------

    try:

        records = parse_pdf_by_text(
            data
        )

        if records:

            return records

    except Exception:
        pass

    raise ValueError(
        "PDF 成绩单解析失败。"
        "请确认该 PDF 是研究生管理系统导出的成绩单，"
        "而不是扫描图片版 PDF。"
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
# 第一学年筛选
# ============================================================

def select_first_year(all_df):

    # --------------------------------------------------------
    # 如果有学期信息
    # --------------------------------------------------------

    if all_df[
        "学期"
    ].notna().any():

        calc_df = all_df[
            all_df["学期"].isin(
                [
                    1,
                    2,
                ]
            )
        ].copy()

        if len(calc_df) > 0:

            return (
                calc_df,
                "已按第 1、2 学期识别第一学年课程。",
                False,
            )

    # --------------------------------------------------------
    # 没有可靠学期信息
    # --------------------------------------------------------

    return (
        all_df.copy(),
        "未能可靠识别第一学年，当前按全部识别课程计算，请核对。",
        True,
    )


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
            # 转为 DataFrame
            # ------------------------------------------------

            all_df = pd.DataFrame(
                records
            )

            # ------------------------------------------------
            # 第一学年
            # ------------------------------------------------

            calc_df, message, warning = (
                select_first_year(
                    all_df
                )
            )

            if warning:

                st.warning(
                    message
                )

            else:

                st.success(
                    message
                )

            # ------------------------------------------------
            # 计算：成绩 × 学分
            # ------------------------------------------------

            calc_df["成绩×学分"] = (
                calc_df["换算成绩"]
                * calc_df["学分"]
            )

            # ------------------------------------------------
            # 总学分
            # ------------------------------------------------

            total_credit = float(
                calc_df["学分"].sum()
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
            #
            # B / 10 - 5
            # ------------------------------------------------

            grade_point = (
                B / 10 - 5
            )

            # =================================================
            # 结果显示
            # =================================================

            st.write(
                f"加权成绩 B：{B:.2f}；绩点：{grade_point:.3f}"
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
            # 识别明细
            # =================================================

            st.subheader(
                "📋 识别明细"
            )

            show = calc_df[
                [
                    "课程名称",
                    "学分",
                    "原始成绩",
                    "换算成绩",
                    "成绩×学分",
                ]
            ].copy()

            # ------------------------------------------------
            # 保留合理的小数位
            # ------------------------------------------------

            show["学分"] = show[
                "学分"
            ].map(
                lambda x: f"{x:.2f}"
            )

            show["换算成绩"] = show[
                "换算成绩"
            ].map(
                lambda x: f"{x:.2f}"
            )

            show["成绩×学分"] = show[
                "成绩×学分"
            ].map(
                lambda x: f"{x:.2f}"
            )

            st.dataframe(
                show,
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
