import io
import os
import time
import re
import json
import hashlib
from pathlib import Path

import streamlit as st
import fitz  # PyMuPDF
from docx import Document
from docx.shared import Pt, Inches
from docx.oxml.ns import qn
from google import genai
from google.genai import types

APP_NAME = "Coach Winnie – Forms Converter"
DEFAULT_MODEL = os.getenv("GEMINI_MODEL", "gemini-3.6-flash")

APP_VERSION = "V3.3.2 Simple 3 Types + Image Extraction"

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
  <p><strong>{APP_VERSION}</strong></p>
</div>
""", unsafe_allow_html=True)

st.info(
    "上传从 Google Forms 打印/另存的 PDF。系统会保留题目顺序和选项，"
    "并把 Grid / Matching 题拆成 Microsoft Forms 较容易导入的独立题目。"
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
    "✅ V3.3 保留 V3.2 的 3 种题型分类：Choice、Multiple-answer Choice、Open text；另外独立提取 PDF 图片，不让图片处理干扰题型识别。"
    "\n\n"
    "🛡️ 仅增加一次 503 / 429 自动重试，并使用 gemini-3.6-flash。\n\n"
    "🔒 API Key 只用于本次页面会话的转换请求。关闭/刷新页面后请重新输入。"
    "同一个有效的 Gemini API Key 可以重复使用，不需要每次重新建立。"
)

with st.expander("V3.1 转换规则（恢复第一个 Gemini 版本逻辑）", expanded=False):
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



def call_gemini_stable(api_key: str, prompt: str, status_box=None):
    """
    Keep V3 conversion logic unchanged.
    Only add one short retry for temporary 429/503 server-load errors.
    """
    client = genai.Client(api_key=api_key)
    last_error = None

    for attempt in range(2):
        try:
            if status_box:
                if attempt == 0:
                    status_box.write(f"正在连接 Gemini：{DEFAULT_MODEL}")
                else:
                    status_box.write("Gemini 暂时繁忙，3 秒后再尝试一次…")

            response = client.models.generate_content(
                model=DEFAULT_MODEL,
                contents=prompt,
                config=types.GenerateContentConfig(
                    response_mime_type="application/json",
                    temperature=0.1,
                ),
            )
            return response
        except Exception as exc:
            last_error = exc
            msg = str(exc).lower()
            temporary = any(x in msg for x in [
                "503", "unavailable", "high demand",
                "429", "resource_exhausted", "rate limit"
            ])
            if not temporary or attempt == 1:
                raise
            time.sleep(3)

    raise last_error


def build_prompt(form_text: str, answer_text: str = "", ui_language: str = "保持原文件语言") -> str:
    answer_section = ""
    if answer_text.strip():
        answer_section = f"""
ANSWER SOURCE:
{answer_text}

Use the answer source only when it explicitly provides an answer.
Never infer or guess an answer.
"""

    return f"""
You are converting a Google Forms PDF into a Microsoft Forms Quick Import Word document.

IMPORTANT: Use ONLY these 3 output question categories:

1. choice
2. multiple_answers
3. open_text

Do NOT create any other output question category.

SIMPLE CLASSIFICATION RULES

A. CHOICE
Use "choice" for:
- Google Forms Multiple choice
- Dropdown
- Linear scale
- Multiple choice grid: split EACH ROW into a separate choice question
- Likert-style grid: split EACH ROW / statement into a separate choice question
- Matching: split into separate choice questions whenever the options are clear
- Ranking: convert to choice for import compatibility

B. MULTIPLE-ANSWER CHOICE
Use "multiple_answers" for:
- Google Forms Checkboxes
- Checkbox grid: split EACH ROW into a separate multiple_answers question

The Word document still uses normal choice-option formatting.
Microsoft Forms may require the teacher to manually enable "Multiple answers" after import.

C. OPEN TEXT
Use "open_text" for:
- Short answer
- Paragraph
- Date
- Any question that genuinely has no selectable options
- Any question that cannot be safely converted to Choice without inventing content

CORE RULE
If the original question has selectable options, preserve those options and prefer Choice / Multiple-answer Choice.
Do NOT turn a valid option-based question into Open text merely because formatting is imperfect.

CONTENT RULES
- Interface/output language preference: {ui_language}
- Preserve original wording, numbering, order, options, language, names, dates and values.
- Do not add, rewrite, summarize or invent question content.
- For grids, preserve the original row order and column choices.
- Do not invent missing options.
- If an image is necessary to answer a question, set image_required=true and preserve any visible image reference text, but do not invent image content.
- If no answer source is supplied, do not infer answers.

Return valid JSON only in this structure:
{{
  "title": "",
  "description": "",
  "questions": [
    {{
      "number": 1,
      "question": "",
      "type": "choice",
      "options": [],
      "answer": "",
      "image_required": false,
      "note": ""
    }}
  ]
}}

Allowed values for "type" are ONLY:
- "choice"
- "multiple_answers"
- "open_text"

GOOGLE FORM CONTENT:
{form_text}

{answer_section}
"""



def extract_pdf_images(pdf_bytes: bytes):
    """Extract useful raster images without changing AI question classification."""
    images, seen = [], set()
    pdf = fitz.open(stream=pdf_bytes, filetype="pdf")
    try:
        for page_index in range(len(pdf)):
            page = pdf[page_index]
            for image_index, item in enumerate(page.get_images(full=True), start=1):
                try:
                    info = pdf.extract_image(item[0])
                    raw = info.get("image", b"")
                    ext = (info.get("ext") or "png").lower()
                    width = int(info.get("width") or 0)
                    height = int(info.get("height") or 0)
                    if len(raw) < 1500 or width < 80 or height < 80 or width * height < 10000:
                        continue
                    digest = hashlib.sha1(raw).hexdigest()
                    if digest in seen:
                        continue
                    seen.add(digest)
                    images.append({
                        "id": f"P{page_index+1:02d}_IMG{image_index:02d}",
                        "page": page_index + 1,
                        "filename": f"page_{page_index+1:02d}_image_{image_index:02d}.{ext}",
                        "bytes": raw,
                    })
                except Exception:
                    continue
    finally:
        pdf.close()
    return images


def make_reference_docx(structured: dict, quiz_mode: bool, images) -> bytes:
    """Make a reference copy: converted questions first, extracted images afterwards by PDF page."""
    base = make_docx(structured, quiz_mode)
    doc = Document(io.BytesIO(base))
    if images:
        doc.add_page_break()
        p = doc.add_paragraph()
        r = p.add_run("Extracted Images from Original PDF")
        r.bold = True
        r.font.size = Pt(16)
        doc.add_paragraph(
            "These images are saved for reference. Microsoft Forms Quick Import may require manual image insertion."
        )
        for img in images:
            doc.add_paragraph(f'{img["id"]} — PDF page {img["page"]}')
            try:
                doc.add_picture(io.BytesIO(img["bytes"]), width=Inches(4.8))
            except Exception:
                doc.add_paragraph(f'[Saved separately: {img["filename"]}]')
    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


def make_image_bundle(quick_docx: bytes, reference_docx: bytes, images, stem: str) -> bytes:
    out = io.BytesIO()
    lines = [
        "Coach Winnie – Forms Converter V3.3",
        "Images are extracted independently and do not affect question classification.",
        "",
        "Image mapping:"
    ]
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr(f"{stem}_Quick_Import.docx", quick_docx)
        z.writestr(f"{stem}_Reference_With_Images.docx", reference_docx)
        for img in images:
            z.writestr(f'images/{img["filename"]}', img["bytes"])
            lines.append(f'{img["id"]} -> PDF page {img["page"]} -> images/{img["filename"]}')
        z.writestr("image_mapping.txt", "\n".join(lines))
    return out.getvalue()


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

            if qtype in ("open_text", "short_answer", "paragraph") and not opts:
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
        st.error("请先输入你自己的 Gemini API Key。")
        st.stop()

    try:
        with st.status("正在读取并转换表单…", expanded=True) as status:
            st.write("读取 PDF 内容")
            pdf_bytes = pdf_file.getvalue()
            form_text = extract_pdf_text(pdf_bytes)
            extracted_images = extract_pdf_images(pdf_bytes)
            st.write(f"已自动提取 {len(extracted_images)} 张图片；图片处理不会影响 3 种题型分类。")

            st.write("读取答案文件" if answer_file else "没有答案文件，将不会推测答案")
            answer_text = extract_answer_text(answer_file)

            quiz_mode = output_mode.startswith("Quiz")
            prompt = build_prompt(form_text, answer_text)

            st.write("Gemini 正在识别题目、选项与 Grid / Matching 结构")
            response = call_gemini_stable(
                api_key=api_key,
                prompt=prompt,
                status_box=status,
            )
            raw = response.text or ""
            structured = json.loads(clean_json_text(raw))

            st.write("正在生成 Microsoft Forms Quick Import Word")
            docx_bytes = make_docx(structured, quiz_mode)
            reference_docx_bytes = make_reference_docx(structured, quiz_mode, extracted_images)
            stem = re.sub(r'[\\/:*?"<>|]+', "_", Path(pdf_file.name).stem)
            image_bundle_bytes = make_image_bundle(
                docx_bytes, reference_docx_bytes, extracted_images, stem
            )
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

        st.download_button(
            "🖼️ Download Reference Word with Images",
            data=reference_docx_bytes,
            file_name=f"{Path(pdf_file.name).stem}_Reference_With_Images.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )

        st.download_button(
            "📦 Download Word + Images ZIP",
            data=image_bundle_bytes,
            file_name=f"{Path(pdf_file.name).stem}_V3_3_Word_Images.zip",
            mime="application/zip",
            use_container_width=True,
        )

    except json.JSONDecodeError:
        st.error("Gemini 输出无法解析为结构化数据，请重试。若 PDF 很复杂，可分成较小部分转换。")
    except Exception as e:
        msg = str(e)
        if "API_KEY_INVALID" in msg or "API key not valid" in msg or "INVALID_ARGUMENT" in msg:
            st.error("Gemini API Key 无效或请求设置不正确。请到 Google AI Studio 检查 API Key 后再试。")
        elif "RESOURCE_EXHAUSTED" in msg or "429" in msg:
            st.error("Gemini API 当前达到免费额度或速率限制。请稍后再试。")
        elif "503" in msg or "UNAVAILABLE" in msg or "high demand" in msg.lower():
            st.error("Gemini 服务目前繁忙。系统已自动重试一次，请稍后再按 Convert。")
        else:
            st.error(f"转换失败：{e}")

st.divider()
st.markdown("""
<div class="small">
<strong>Microsoft Forms 导入：</strong>
Microsoft Forms → Quick Import → Upload from this device → 选择生成的 Word → Form / Quiz。<br>
<strong>Privacy：</strong>
上传内容会使用你本次输入的 Gemini API Key 发送到 Google Gemini API 进行转换；本 App 不会将 API Key 写入 GitHub、文件或数据库。请勿上传不应交由该服务处理的敏感资料。
</div>
""", unsafe_allow_html=True)
