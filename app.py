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

/* 学校名称 */

.school {
    font-size: 15px;
    color: #666;
    margin-bottom: 3px;
}

/* 页面标题 */

.title {
    font-size: 27px;
    font-weight: 800;
    line-height: 1.25;
}

/* 副标题 */

.subtitle {
    font-size: 15px;
    color: #666;
    margin-top: 4px;
}


/* ============================================================
   成绩结果卡片
   ============================================================ */

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


/* ============================================================
   页脚
   ============================================================ */

.footer {
    font-size: 12px;
    color: #8a8f95;
    text-align: center;
    margin-top: 30px;
}


/* ============================================================
   Metric
   ============================================================ */

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

    return re.sub(
        r"\s+",
        "",
        str(x or ""),
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

        return float(
            GRADE_MAP[s]
        )

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
    """

    m = re.search(
        r"\d+(?:\.\d+)?",
        norm(x),
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


# ============================================================
# 成绩单表格解析
# ============================================================

def parse_rows(rows):

    # --------------------------------------------------------
    # 第一步：清理空行
    # --------------------------------------------------------

    cleaned = []

    for row in rows:

        vals = [
            norm(v)
            for v in row
        ]

        if any(vals):

            cleaned.append(
                vals
            )

    # --------------------------------------------------------
    # 第二步：寻找成绩单表头
    # --------------------------------------------------------

    header = None

    for hi, row in enumerate(
        cleaned[:20]
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
    # 第三步：逐行解析课程
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

        # 如果某一行列数不足，用空字符串补齐
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

        # ----------------------------------------------------
        # 无效记录跳过
        # ----------------------------------------------------

        if (
            not course
            or credit is None
            or score is None
        ):

            continue

        # ----------------------------------------------------
        # 保存有效课程
        # ----------------------------------------------------

        records.append(
            {
                "课程名称": course,
                "学分": credit,
                "原始成绩": raw,
                "换算成绩": score,
                "学期": semester,
            }
        )

    # --------------------------------------------------------
    # 第四步：特殊处理第一外国语（硕士英语II）
    # --------------------------------------------------------
    #
    # 实际成绩单中：
    #
    # 第一外国语（硕士英语II）
    # 记录1 -> 0学分
    # 记录2 -> 3学分
    #
    # 实际应该：
    #
    # 记录1 -> 1.5学分
    # 记录2 -> 1.5学分
    #
    # 因此两条记录统一修正为1.5学分。
    # --------------------------------------------------------

    for record in records:

        course_name = norm(
            str(
                record["课程名称"]
            )
        )

        if (
            "第一外国语" in course_name
            and (
                "硕士英语II" in course_name
                or "硕士英语Ⅱ" in course_name
            )
        ):

            record["学分"] = 1.5

    # --------------------------------------------------------
    # 第五步：检查是否成功识别课程
    # --------------------------------------------------------

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
# PDF 解析
# ============================================================

def parse_pdf(data):

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

                rows.extend(
                    table
                )

    return parse_rows(rows)


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
            # 第1、2学期 = 第一学年
            # ------------------------------------------------

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

                if len(calc_df) == 0:

                    calc_df = all_df.copy()

                    st.warning(
                        "未能可靠识别第一学年，"
                        "当前按全部识别课程计算，请核对。"
                    )

                else:

                    st.success(
                        "已按第 1、2 学期识别第一学年课程。"
                    )

            else:

                calc_df = all_df.copy()

                st.warning(
                    "成绩单缺少可识别的学期信息，"
                    "当前按全部课程计算，请核对。"
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
                calc_df["成绩×学分"].sum()
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
            # 公式：
            # 绩点 = B / 10 - 5
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
