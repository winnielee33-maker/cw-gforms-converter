import io
import os
import re
import json
from pathlib import Path
from collections import Counter

import streamlit as st
import fitz  # PyMuPDF
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from google import genai
from google.genai import types

APP_NAME = "Coach Winnie – Forms Converter"
APP_VERSION = "V4 Compatibility Mode"
DEFAULT_MODEL = os.getenv("GEMINI_MODEL", "gemini-3.7-flash")

st.set_page_config(
    page_title=APP_NAME,
    page_icon="📝",
    layout="centered",
)

st.markdown("""
<style>
.block-container {max-width: 960px; padding-top: 2rem; padding-bottom: 3rem;}
.hero {
  padding: 1.4rem 1.6rem;
  border-radius: 22px;
  background: linear-gradient(135deg, rgba(98,72,255,.12), rgba(40,180,160,.10));
  border: 1px solid rgba(120,120,120,.18);
  margin-bottom: 1.1rem;
}
.hero h1 {margin:0; font-size:2rem;}
.hero p {margin:.45rem 0 0 0; opacity:.78;}
.small {font-size:.9rem; opacity:.72;}
.report-card {
  border:1px solid rgba(120,120,120,.18);
  border-radius:16px;
  padding:1rem 1.1rem;
  margin:.6rem 0;
}
</style>
""", unsafe_allow_html=True)

st.markdown(f"""
<div class="hero">
  <h1>📝 {APP_NAME}</h1>
  <p>Google Forms PDF → Microsoft Forms Quick Import Word (.docx)</p>
  <p><strong>{APP_VERSION}</strong></p>
</div>
""", unsafe_allow_html=True)

st.info(
    "V4 会优先生成 Microsoft Forms Quick Import 较容易识别的题型。"
    "Google Forms 的 Grid / Likert / Matching 会自动转换成兼容的独立 Choice 题。"
)

st.subheader("🔑 使用自己的 Gemini API Key")
st.markdown(
    "第一次使用？可到 **Google AI Studio** 建立 Gemini API Key："
    "[Get API Key](https://aistudio.google.com/app/apikey)"
)
api_key_input = st.text_input(
    "Gemini API Key",
    type="password",
    placeholder="AIza...",
    help="每次使用时输入。此 App 不会把你的 API Key 写入 GitHub、文件或数据库。"
)
st.caption(
    "🔒 API Key 只用于本次页面会话的转换请求。关闭/刷新页面后请重新输入。"
    "同一个有效的 Gemini API Key 可以重复使用，不需要每次重新建立。"
)

with st.expander("V4 Compatibility Mode 转换规则", expanded=False):
    st.markdown("""
- **Multiple choice / Dropdown** → Choice
- **Checkboxes** → Choice（Multiple Answers）
- **Short answer / Paragraph** → Text
- **Linear scale** → Choice，并保留原来的刻度选项
- **Multiple choice grid / Likert-style grid** → 每一行拆成一个独立 Choice，保留原列选项
- **Checkbox grid** → 每一行拆成一个独立 Multiple Answers Choice
- **Matching** → 尽量拆成独立 Choice；无法可靠判断时标记人工检查
- **Date** → Text，并在转换报告中提示导入后可手动改成 Date
- **Ranking** → Text/Choice，并在转换报告中提示导入后手动改成 Ranking
- **Image / Map / Chart** → 保留题干，并提示导入后手动补图片
- 不推测 PDF 中没有显示的正确答案
- 不自行改写、总结或补充原题内容
""")

pdf_file = st.file_uploader("① 上传 Google Forms PDF", type=["pdf"])

answer_file = st.file_uploader(
    "② 答案文件（可选）",
    type=["pdf", "txt", "docx"],
    help="如果答案不在原 PDF 中，可另上传答案版。没有答案文件也可以直接转换。"
)

col1, col2 = st.columns(2)
with col1:
    output_mode = st.radio(
        "输出类型",
        ["Form（无答案）", "Quiz（有答案时加入答案）"],
        index=0
    )
with col2:
    language_note = st.selectbox(
        "界面提示语言",
        ["保持原文件语言", "中文提示", "English prompts"],
        index=0
    )


def extract_pdf_text(data: bytes) -> str:
    doc = fitz.open(stream=data, filetype="pdf")
    chunks = []
    for i, page in enumerate(doc):
        txt = page.get_text("text")
        chunks.append(f"\n===== PAGE {i+1} =====\n{txt}")
    return "\n".join(chunks)


def extract_docx_text(data: bytes) -> str:
    tmp = io.BytesIO(data)
    doc = Document(tmp)
    return "\n".join(p.text for p in doc.paragraphs)


def extract_answer_text(uploaded) -> str:
    if uploaded is None:
        return ""
    data = uploaded.getvalue()
    suffix = Path(uploaded.name).suffix.lower()
    if suffix == ".pdf":
        return extract_pdf_text(data)
    if suffix == ".docx":
        return extract_docx_text(data)
    return data.decode("utf-8", errors="ignore")


def clean_json_text(s: str) -> str:
    s = s.strip()
    s = re.sub(r"^```(?:json)?\s*", "", s, flags=re.I)
    s = re.sub(r"\s*```$", "", s)
    start = s.find("{")
    end = s.rfind("}")
    if start >= 0 and end > start:
        s = s[start:end+1]
    return s


def build_prompt(form_text: str, answer_text: str, quiz_mode: bool) -> str:
    answer_instruction = (
        "An answer source is provided. Add answers ONLY when explicitly supported by that answer source. "
        "Never infer missing answers."
        if answer_text else
        "No answer source is provided. Do NOT infer or invent any answer."
    )

    return f"""
You are the conversion engine for \"Coach Winnie – Forms Converter V4 Compatibility Mode\".

GOAL
Convert a Google Forms print/PDF into a structure optimized for Microsoft Forms Quick Import Word (.docx).
The import document should use only broadly compatible structures: choice, multiple-answer choice, and open text.

NON-NEGOTIABLE CONTENT RULES
1. Use only information present in the supplied form text and optional answer source.
2. Preserve original title, description, section order, question order, wording and options as faithfully as possible.
3. Do not add, delete, summarize, rewrite, correct, explain or supplement the user's question content.
4. Do not infer answers. {answer_instruction}
5. Do not include Google printing UI text such as timestamps, URLs, \"Mark only one oval\", \"Tick all that apply\", page footers, navigation labels, or form editing controls as question content.
6. If parsing is uncertain, preserve visible text, choose the safest compatible output, and set review_required=true.

V4 COMPATIBILITY MAPPING
A. Google Multiple choice -> output_type=single_choice.
B. Google Dropdown -> output_type=single_choice.
C. Google Checkboxes -> output_type=multiple_answers.
D. Google Short answer -> output_type=short_answer.
E. Google Paragraph -> output_type=paragraph.
F. Google Linear scale -> output_type=single_choice using the original visible scale values/options.
G. Google Multiple choice grid / Likert-style grid:
   - Create ONE independent single_choice question for EACH ROW.
   - Use the original row label as the question text.
   - If the grid has an overall title, preserve context by combining only original text in this form: "<grid title> — <row label>".
   - Use the original column labels as options for every expanded row.
   - original_type must be "multiple_choice_grid" and conversion_action must be "expanded_grid_row_to_choice".
H. Google Checkbox grid:
   - Create ONE independent multiple_answers question for EACH ROW.
   - Use the original row label as the question text, optionally prefixed by the original grid title as above.
   - Use original column labels as options.
   - original_type="checkbox_grid" and conversion_action="expanded_grid_row_to_multiple_answers".
I. Matching questions:
   - If clearly expressed as rows + shared choices, expand each row to a single_choice question.
   - Otherwise use short_answer and set review_required=true.
J. Date:
   - output_type=short_answer.
   - original_type="date" and conversion_action="date_to_text_manual_change_recommended".
K. Ranking:
   - Preserve the visible prompt and items. Use short_answer unless a reliable choice representation is obvious.
   - original_type="ranking" and conversion_action="ranking_manual_change_recommended".
L. Image / map / chart dependent question:
   - Preserve the question text.
   - image_required=true.
   - Do not hallucinate missing image content.

QUIZ MODE
quiz_mode = {str(quiz_mode).lower()}
If quiz_mode is false, leave answers empty even if an answer source exists.
If quiz_mode is true, include answer labels/text only when explicitly supported by the supplied answer source.

RETURN STRICT JSON ONLY with this schema:
{{
  "title": "string",
  "description": "string",
  "sections": [
    {{
      "title": "string",
      "description": "string",
      "questions": [
        {{
          "number": "string",
          "question": "string",
          "output_type": "single_choice|multiple_answers|short_answer|paragraph",
          "original_type": "multiple_choice|dropdown|checkboxes|short_answer|paragraph|linear_scale|multiple_choice_grid|checkbox_grid|matching|date|ranking|image_question|unknown",
          "conversion_action": "kept|expanded_grid_row_to_choice|expanded_grid_row_to_multiple_answers|matching_expanded_to_choice|date_to_text_manual_change_recommended|ranking_manual_change_recommended|image_manual_insert_required|fallback_review_required",
          "options": ["string"],
          "answer": ["string"],
          "required": true,
          "image_required": false,
          "review_required": false
        }}
      ]
    }}
  ]
}}

FORM TEXT
----------------
{form_text}

OPTIONAL ANSWER SOURCE
----------------
{answer_text if answer_text else "(none)"}
"""


def make_docx(structured: dict, quiz_mode: bool) -> bytes:
    doc = Document()

    normal = doc.styles["Normal"]
    normal.font.name = "Microsoft YaHei"
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), "Microsoft YaHei")
    normal.font.size = Pt(11)

    title = structured.get("title") or "Microsoft Forms Import"
    p = doc.add_paragraph()
    r = p.add_run(title)
    r.bold = True
    r.font.size = Pt(18)

    desc = (structured.get("description") or "").strip()
    if desc:
        doc.add_paragraph(desc)

    for sec in structured.get("sections", []):
        sec_title = (sec.get("title") or "").strip()
        sec_desc = (sec.get("description") or "").strip()
        if sec_title:
            p = doc.add_paragraph()
            r = p.add_run(sec_title)
            r.bold = True
            r.font.size = Pt(14)
        if sec_desc:
            doc.add_paragraph(sec_desc)

        for q in sec.get("questions", []):
            num = (q.get("number") or "").strip()
            text = (q.get("question") or "").strip()
            qtype = q.get("output_type", "short_answer")
            required = bool(q.get("required", False))
            image_required = bool(q.get("image_required", False))

            prefix = f"{num} " if num else ""
            suffix = ""
            if qtype == "multiple_answers":
                suffix += " (Multiple Answers)"
            if required:
                suffix += " *"

            p = doc.add_paragraph()
            rr = p.add_run(f"{prefix}{text}{suffix}".strip())
            rr.bold = True

            if image_required:
                doc.add_paragraph("[IMAGE REQUIRED – Add original image manually after import]")

            opts = q.get("options") or []
            for opt in opts:
                doc.add_paragraph(str(opt))

            if qtype in ("short_answer", "paragraph") and not opts:
                doc.add_paragraph("")

            if quiz_mode:
                ans = q.get("answer") or []
                if ans:
                    label = "Answers: " if len(ans) > 1 else "Answer: "
                    pa = doc.add_paragraph()
                    ra = pa.add_run(label + ", ".join(map(str, ans)))
                    ra.bold = True

            doc.add_paragraph("")

    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()


def safe_filename(name: str) -> str:
    stem = Path(name).stem
    stem = re.sub(r'[\\/:*?"<>|]+', "_", stem)
    return f"{stem}_Microsoft_Forms_Import_V4.docx"


def build_conversion_report(structured: dict):
    questions = [
        q
        for sec in structured.get("sections", [])
        for q in sec.get("questions", [])
    ]
    output_types = Counter(q.get("output_type", "unknown") for q in questions)
    original_types = Counter(q.get("original_type", "unknown") for q in questions)
    actions = Counter(q.get("conversion_action", "kept") for q in questions)

    review_count = sum(1 for q in questions if q.get("review_required"))
    image_count = sum(1 for q in questions if q.get("image_required"))
    manual_date = actions.get("date_to_text_manual_change_recommended", 0)
    manual_ranking = actions.get("ranking_manual_change_recommended", 0)
    expanded_choice = actions.get("expanded_grid_row_to_choice", 0)
    expanded_multi = actions.get("expanded_grid_row_to_multiple_answers", 0)
    matching_expanded = actions.get("matching_expanded_to_choice", 0)

    return {
        "total": len(questions),
        "output_types": output_types,
        "original_types": original_types,
        "review_count": review_count,
        "image_count": image_count,
        "manual_date": manual_date,
        "manual_ranking": manual_ranking,
        "expanded_choice": expanded_choice,
        "expanded_multi": expanded_multi,
        "matching_expanded": matching_expanded,
    }


if st.button("✨ Convert to Microsoft Forms Word", type="primary", use_container_width=True):
    if not pdf_file:
        st.error("请先上传 Google Forms PDF。")
        st.stop()

    api_key = api_key_input.strip()
    if not api_key:
        st.error("请先输入你自己的 Gemini API Key。")
        st.stop()

    try:
        with st.status("正在读取并转换表单…", expanded=True) as status:
            st.write("读取 PDF 内容")
            form_text = extract_pdf_text(pdf_file.getvalue())

            st.write("读取答案文件" if answer_file else "没有答案文件，将不会推测答案")
            answer_text = extract_answer_text(answer_file)

            quiz_mode = output_mode.startswith("Quiz")
            prompt = build_prompt(form_text, answer_text, quiz_mode)

            st.write("Gemini 正在识别题型并执行 Quick Import Compatibility 转换")
            client = genai.Client(api_key=api_key)
            response = client.models.generate_content(
                model=DEFAULT_MODEL,
                contents=prompt,
                config=types.GenerateContentConfig(
                    response_mime_type="application/json",
                    temperature=0.1,
                ),
            )
            raw = response.text or ""
            structured = json.loads(clean_json_text(raw))

            st.write("正在生成 Microsoft Forms Quick Import Word")
            docx_bytes = make_docx(structured, quiz_mode)
            report = build_conversion_report(structured)
            status.update(label="转换完成 ✅", state="complete")

        st.success(f"完成：共生成 {report['total']} 个 Microsoft Forms Quick Import 兼容题目。")

        st.subheader("📊 转换报告")
        c1, c2, c3 = st.columns(3)
        c1.metric("Choice", report["output_types"].get("single_choice", 0))
        c2.metric("Multiple Answers", report["output_types"].get("multiple_answers", 0))
        c3.metric(
            "Text",
            report["output_types"].get("short_answer", 0) + report["output_types"].get("paragraph", 0),
        )

        notes = []
        if report["expanded_choice"]:
            notes.append(f"Grid / Likert：已拆成 {report['expanded_choice']} 个独立 Choice 题。")
        if report["expanded_multi"]:
            notes.append(f"Checkbox Grid：已拆成 {report['expanded_multi']} 个 Multiple Answers 题。")
        if report["matching_expanded"]:
            notes.append(f"Matching：已拆成 {report['matching_expanded']} 个 Choice 题。")
        if report["manual_date"]:
            notes.append(f"Date：{report['manual_date']} 题先转为 Text；导入后可手动改成 Date。")
        if report["manual_ranking"]:
            notes.append(f"Ranking：{report['manual_ranking']} 题需要导入后手动调整为 Ranking。")
        if report["image_count"]:
            notes.append(f"图片 / 地图 / 图表：{report['image_count']} 题需要导入后手动补图。")
        if report["review_count"]:
            notes.append(f"人工检查：{report['review_count']} 题解析不确定，建议导入后核对。")

        if notes:
            st.warning("\n\n".join(notes))
        else:
            st.info("没有检测到需要额外人工转换的特殊题型。")

        st.download_button(
            "⬇️ Download Microsoft Forms Word",
            data=docx_bytes,
            file_name=safe_filename(pdf_file.name),
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )

        st.caption(
            "建议：Microsoft Forms → Quick Import 后，再检查 Multiple Answers、Date、Ranking、图片及复杂 Grid 题。"
        )

    except json.JSONDecodeError:
        st.error("Gemini 输出无法解析为结构化数据，请重试。若 PDF 很复杂，可分成较小部分转换。")
    except Exception as e:
        msg = str(e)
        if "API_KEY_INVALID" in msg or "API key not valid" in msg or "INVALID_ARGUMENT" in msg:
            st.error("Gemini API Key 无效或请求设置不正确。请到 Google AI Studio 检查 API Key 后再试。")
        elif "RESOURCE_EXHAUSTED" in msg or "429" in msg:
            st.error("Gemini API 当前额度或速率限制已用完。请稍后再试，或检查 Google AI Studio / Google Cloud 的 API 配额。")
        else:
            st.error(f"转换失败：{e}")

st.divider()
st.markdown("""
<div class="small">
<strong>Microsoft Forms 导入：</strong>
Microsoft Forms → Quick Import → Upload from this device → 选择生成的 Word → Form / Quiz。<br>
<strong>V4 Compatibility Mode：</strong>
复杂题型会先转换成 Quick Import 更容易接受的 Choice / Multiple Answers / Text，然后在转换报告中提示需要手动调整的项目。<br>
<strong>Privacy：</strong>
上传内容会使用你本次输入的 Gemini API Key 发送到 Google Gemini API 进行转换；本 App 不会将 API Key 写入 GitHub、文件或数据库。请勿上传不应交由该服务处理的敏感资料。
</div>
""", unsafe_allow_html=True)
