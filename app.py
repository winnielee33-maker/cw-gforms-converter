
import io
import os
import re
import json
import tempfile
from pathlib import Path

import streamlit as st
import fitz  # PyMuPDF
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from openai import OpenAI

APP_NAME = "Coach Winnie – Forms Converter"
DEFAULT_MODEL = os.getenv("OPENAI_MODEL", "gpt-5.6-terra")

st.set_page_config(
    page_title=APP_NAME,
    page_icon="📝",
    layout="centered",
)

st.markdown("""
<style>
.block-container {max-width: 920px; padding-top: 2rem; padding-bottom: 3rem;}
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
</style>
""", unsafe_allow_html=True)

st.markdown(f"""
<div class="hero">
  <h1>📝 {APP_NAME}</h1>
  <p>Google Forms PDF → Microsoft Forms Quick Import Word (.docx)</p>
</div>
""", unsafe_allow_html=True)

st.info(
    "上传从 Google Forms 打印/另存的 PDF。系统会保留题目顺序和选项，"
    "并把 Grid / Matching 题拆成 Microsoft Forms 较容易导入的独立题目。"
)

st.subheader("🔑 使用自己的 OpenAI API Key")
api_key_input = st.text_input(
    "OpenAI API Key",
    type="password",
    placeholder="sk-...",
    help="每次使用时输入。此 App 不会把你的 API Key 写入 GitHub、文件或数据库。"
)
st.caption("🔒 API Key 只用于本次页面会话的转换请求。关闭/刷新页面后请重新输入。ChatGPT Free/Plus 订阅与 API 额度分开。")

with st.expander("转换规则", expanded=False):
    st.markdown("""
- 不添加、删除、总结或改写原题。
- 不推测 PDF 中没有显示的正确答案。
- 单选题保留为单选格式。
- Checkbox / Multiple Answers 保留并标示为 **(Multiple Answers)**。
- Short Answer / Paragraph 保留为开放题。
- Multiple Choice Grid / Matching 会拆成多个独立选择题。
- 图片 / 地图 / 图表题保留题干，并标示 **[IMAGE REQUIRED – Add original image manually after import]**。
- 若上传答案文件，只使用答案文件明确显示的答案，不自行推测。
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
You are the conversion engine for "Coach Winnie – Forms Converter".

TASK
Convert a Google Forms print/PDF into a structure suitable for Microsoft Forms Quick Import Word (.docx).

NON-NEGOTIABLE RULES
1. Use only information present in the supplied form text and optional answer source.
2. Preserve original title, description, section order, question order, numbering, wording and options as faithfully as possible.
3. Do not add, delete, summarize, rewrite, correct, explain or supplement question content.
4. Do not infer answers. {answer_instruction}
5. Preserve whether a question is single choice, multiple answers, short answer/paragraph, or grid/matching when identifiable.
6. Convert each grid/matching row into a separate single-choice question using the original row label and original column choices.
7. For any image/map/chart-dependent question, preserve the question text and add exactly:
   [IMAGE REQUIRED – Add original image manually after import]
8. If parsing is uncertain, preserve the visible text and set "review_required": true instead of guessing.
9. Do not include Google footer text, timestamps, page URLs, "Mark only one oval", "Tick all that apply", or similar printing UI labels as question content.

QUIZ MODE
quiz_mode = {str(quiz_mode).lower()}
If quiz_mode is false, leave answers empty even if an answer source exists.
If quiz_mode is true, include answer labels/text only when explicitly supported by the supplied answer source.

RETURN STRICT JSON ONLY, with this schema:
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
          "type": "single_choice|multiple_answers|short_answer|paragraph|image_question",
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
            qtype = q.get("type", "short_answer")
            required = bool(q.get("required", False))
            image_required = bool(q.get("image_required", False))
            review_required = bool(q.get("review_required", False))

            prefix = f"{num} " if num else ""
            suffix = ""
            if qtype == "multiple_answers":
                suffix += " (Multiple Answers)"
            if required:
                suffix += " *"

            p = doc.add_paragraph()
            rr = p.add_run(f"{prefix}{text}{suffix}".strip())
            rr.bold = True

            if image_required or qtype == "image_question":
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

            if review_required:
                pr = doc.add_paragraph()
                rr = pr.add_run("[REVIEW REQUIRED]")
                rr.italic = True

            doc.add_paragraph("")

    final = doc.add_paragraph()
    r = final.add_run("Import note: ")
    r.bold = True
    final.add_run(
        "After Quick Import, check Multiple Answers settings and manually add any required images, maps or charts."
    )

    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

def safe_filename(name: str) -> str:
    stem = Path(name).stem
    stem = re.sub(r'[\\/:*?"<>|]+', "_", stem)
    return f"{stem}_Microsoft_Forms_Import.docx"

if st.button("✨ Convert to Microsoft Forms Word", type="primary", use_container_width=True):
    if not pdf_file:
        st.error("请先上传 Google Forms PDF。")
        st.stop()

    api_key = api_key_input.strip()
    if not api_key:
        st.error("请先输入你自己的 OpenAI API Key。")
        st.stop()

    try:
        with st.status("正在读取并转换表单…", expanded=True) as status:
            st.write("读取 PDF 内容")
            form_text = extract_pdf_text(pdf_file.getvalue())

            st.write("读取答案文件" if answer_file else "没有答案文件，将不会推测答案")
            answer_text = extract_answer_text(answer_file)

            quiz_mode = output_mode.startswith("Quiz")
            prompt = build_prompt(form_text, answer_text, quiz_mode)

            st.write("AI 正在识别题目、选项与 Grid / Matching 结构")
            client = OpenAI(api_key=api_key)
            response = client.responses.create(
                model=DEFAULT_MODEL,
                input=prompt,
            )
            raw = response.output_text
            structured = json.loads(clean_json_text(raw))

            st.write("正在生成 Microsoft Forms Quick Import Word")
            docx_bytes = make_docx(structured, quiz_mode)
            status.update(label="转换完成 ✅", state="complete")

        q_count = sum(len(s.get("questions", [])) for s in structured.get("sections", []))
        review_count = sum(
            1 for s in structured.get("sections", [])
            for q in s.get("questions", [])
            if q.get("review_required")
        )
        img_count = sum(
            1 for s in structured.get("sections", [])
            for q in s.get("questions", [])
            if q.get("image_required") or q.get("type") == "image_question"
        )

        st.success(f"完成：共识别 {q_count} 个导入题目。")
        if review_count or img_count:
            st.warning(
                f"需要人工检查：{review_count} 题；需要手动补图片/地图/图表：{img_count} 题。"
            )

        st.download_button(
            "⬇️ Download Microsoft Forms Word",
            data=docx_bytes,
            file_name=safe_filename(pdf_file.name),
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )

    except json.JSONDecodeError:
        st.error("AI 输出无法解析为结构化数据，请重试。若 PDF 很复杂，可分成较小部分转换。")
    except Exception as e:
        st.error(f"转换失败：{e}")

st.divider()
st.markdown("""
<div class="small">
<strong>Microsoft Forms 导入：</strong>
Microsoft Forms → Quick Import → Upload from this device → 选择生成的 Word → Form / Quiz。<br>
<strong>Privacy：</strong>
上传内容会使用你本次输入的 OpenAI API Key 发送到 OpenAI API 进行转换；本 App 不会将 API Key 写入 GitHub、文件或数据库。请勿上传不应交由该服务处理的敏感资料。
</div>
""", unsafe_allow_html=True)
