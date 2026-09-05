import io
import os
import re
import json
import hashlib
import zipfile
import time
from pathlib import Path
from collections import Counter, defaultdict

import streamlit as st
import fitz  # PyMuPDF
from docx import Document
from docx.shared import Pt, Inches
from docx.oxml.ns import qn
from google import genai
from google.genai import types

APP_NAME = "Coach Winnie – Forms Converter"
APP_VERSION = "V6.7 Quiz Answer Letter Format"
DEFAULT_MODEL = os.getenv("GEMINI_MODEL", "gemini-3.6-flash")

st.set_page_config(page_title=APP_NAME, page_icon="📝", layout="centered")

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
    "V6.5 只使用 3 种题型：Choice、Multiple-answer Choice、Open text。"
    "Gemini 不再判断复杂的 question_type / original_type / conversion_action。"
    "图片提取仍独立进行，不影响题型分类。"
    "注意：Microsoft Forms Quick Import 可能不会把 Word 内图片自动带入题目。"
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
    "🛡️ V6.5：继续使用 gemini-3.6-flash；遇到 503 / high demand 只自动重试一次。题型识别仅保留 3 类，减少 Gemini 判断负担。\n\n"
    "🔒 API Key 只用于本次页面会话。关闭/刷新页面后请重新输入。"
    "同一个有效的 Gemini API Key 可以重复使用。"
)

with st.expander("V6.7 Microsoft Forms Quick Import + Quiz Answer 格式", expanded=False):
    st.markdown("""
**按照 Microsoft Forms Import Guidance：**

- 推荐题型：**Multiple choice**、**Open text**
- 每一题之间要有清楚分隔
- 内容要**垂直排列**
- Quick Import Word **不放图片 / figures**
- 复杂 Grid / Likert / Matching 会先拆成独立 Choice
- Checkboxes / Checkbox grid 会输出为 Choice；导入后请手动开启 **Multiple answers**
- Short answer / Paragraph / 无法安全转换的题型 → Open text
- 图片会另外保存到 ZIP，不影响 Quick Import Word
- Quiz 模式若提供答案文件：Choice / Multiple-answer Choice 会输出 **Answer: A** 或 **Answer: A, C**
- 答案必须来自上传的答案文件；缺少答案时不会推测
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
    st.selectbox(
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


def extract_pdf_images(data: bytes):
    """Extract useful raster images from PDF, keeping page metadata."""
    doc = fitz.open(stream=data, filetype="pdf")
    images = []
    seen = set()

    for page_no, page in enumerate(doc, start=1):
        page_imgs = page.get_images(full=True)
        page_index = 0

        for img in page_imgs:
            xref = img[0]
            try:
                info = doc.extract_image(xref)
            except Exception:
                continue

            blob = info.get("image", b"")
            width = int(info.get("width", 0) or 0)
            height = int(info.get("height", 0) or 0)
            ext = (info.get("ext") or "png").lower()

            # Filter tiny icons, bullets and low-value decoration.
            if not blob or len(blob) < 2500:
                continue
            if width < 120 or height < 80 or width * height < 25000:
                continue

            digest = hashlib.sha1(blob).hexdigest()
            if digest in seen:
                continue
            seen.add(digest)

            page_index += 1
            image_id = f"P{page_no:02d}_IMG{page_index:02d}"
            filename = f"page_{page_no:02d}_image_{page_index:02d}.{ext}"

            images.append({
                "id": image_id,
                "page": page_no,
                "index": page_index,
                "filename": filename,
                "ext": ext,
                "width": width,
                "height": height,
                "bytes": blob,
            })

    return images


def extract_docx_text(data: bytes) -> str:
    doc = Document(io.BytesIO(data))
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


def image_inventory_text(images) -> str:
    if not images:
        return "(No extractable raster images detected.)"
    rows = []
    for im in images:
        rows.append(
            f"{im['id']} | page={im['page']} | size={im['width']}x{im['height']} | file={im['filename']}"
        )
    return "\n".join(rows)


def build_prompt(form_text: str, answer_text: str, quiz_mode: bool, images) -> str:
    answer_instruction = (
        "An answer source is provided. Add answers ONLY when explicitly supported by that answer source. "
        "Never infer missing answers."
        if answer_text else
        "No answer source is provided. Do NOT infer or invent any answer."
    )

    return f"""
You are the conversion engine for "Coach Winnie – Forms Converter V6.7".

GOAL
Convert a Google Forms print/PDF into a structure optimized for Microsoft Forms Quick Import Word (.docx).

IMPORTANT: USE ONLY THESE 3 INTERNAL QUESTION TYPES
1. choice
2. multiple_answers
3. open_text

Do NOT create or return any other question type.
The Word renderer will convert both choice and multiple_answers into Microsoft Forms Multiple choice layout.
Only open_text will render as Microsoft Forms Open text.
Do NOT classify original_type.
Do NOT return conversion_action.
Do NOT create a long taxonomy of Google Forms question types.

SIMPLE CLASSIFICATION

A. choice
Use "choice" when the original question has selectable options and only one answer is expected.
Examples include:
- Multiple choice
- Dropdown
- Linear scale
- Multiple choice Grid
- Likert-style Grid
- Matching
- Other clearly option-based single-answer questions

For Grid / Likert / Matching:
- Split EACH ROW / statement into a separate choice question when the structure is clear.
- Preserve shared column/choice labels as options.
- Preserve original row order.
- If there is an overall title, combine ONLY original text as "<title> — <row label>".

B. multiple_answers
Use "multiple_answers" only when multiple selections are allowed.
Examples include:
- Checkboxes
- Checkbox Grid

For Checkbox Grid:
- Split EACH ROW into a separate multiple_answers question.
- Preserve shared column/choice labels as options.

C. open_text
Use "open_text" for:
- Short answer
- Paragraph
- Questions with no selectable options
- Questions that cannot be safely converted into Choice without inventing content

CORE RULES
1. If the original question clearly has selectable options, preserve them and prefer choice / multiple_answers.
2. Do NOT turn a valid option-based question into open_text merely because PDF formatting is imperfect.
3. If options are genuinely missing or unsafe to reconstruct, use open_text and set review_required=true.
4. Preserve original title, section order, question order, wording, numbering, language, names, dates, values and options.
5. Do not add, delete, rewrite, summarize, correct, explain or supplement question content.
6. Do not infer answers. {answer_instruction}
7. Do not include Google print/UI text as question content.
8. Image processing is separate from question-type classification.
   IMPORTANT: images must NOT be embedded inside the Quick Import Word document.
9. For an image-dependent question:
   - image_required=true
   - source_page = the PDF page number if visible from PAGE markers
   - image_refs may contain ONLY ids from IMAGE INVENTORY and only from the same source page
   - never invent an image id
10. If image matching is uncertain, leave image_refs empty and set review_required=true.

QUIZ MODE
quiz_mode = {str(quiz_mode).lower()}
If quiz_mode is false, leave answers empty.
If quiz_mode is true, include answers ONLY when explicitly supported by the answer source.

ANSWER FORMAT FOR QUIZ
- For choice and multiple_answers questions, return the answer as ENGLISH OPTION LETTERS only: A, B, C, D, E, ...
- Example: if the answer source says "(1) A", return ["A"].
- If the answer source gives the exact option text instead of a letter, map it to the corresponding English option letter only when the match is explicit.
- For multiple correct options, return multiple English letters, e.g. ["A", "C"].
- Never guess a letter.
- If the answer source has no answer for that question, return [].
- Do not add "Answer:" inside the JSON answer value; the Word renderer will add that label.

RETURN STRICT JSON ONLY:
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
          "output_type": "choice|multiple_answers|open_text",
          "options": ["string"],
          "answer": ["string"],
          "required": true,
          "image_required": false,
          "source_page": 1,
          "image_refs": ["P01_IMG01"],
          "review_required": false
        }}
      ]
    }}
  ]
}}

FORM TEXT
----------------
{form_text}

IMAGE INVENTORY
----------------
{image_inventory_text(images)}

OPTIONAL ANSWER SOURCE
----------------
{answer_text if answer_text else "(none)"}
"""

def normalize_for_quick_import(structured: dict):
    """Enforce only 3 output types: choice, multiple_answers, open_text."""
    for sec in structured.get("sections", []):
        for q in sec.get("questions", []):
            qtype = str(q.get("output_type", "open_text") or "open_text").strip().lower()
            opts = [str(x).strip() for x in (q.get("options") or []) if str(x).strip()]

            if qtype not in ("choice", "multiple_answers", "open_text"):
                qtype = "open_text"
                q["review_required"] = True

            if qtype in ("choice", "multiple_answers"):
                if len(opts) >= 2:
                    q["output_type"] = qtype
                    q["options"] = opts
                else:
                    q["output_type"] = "open_text"
                    q["options"] = []
                    q["review_required"] = True
            else:
                q["output_type"] = "open_text"
                q["options"] = []

    return structured

def classify_gemini_error(exc: Exception) -> str:
    msg = str(exc).lower()

    # Model endpoint/access changed: skip this model and try the next supported fallback.
    model_unavailable_signals = [
        "404", "not_found", "not found", "no longer available",
        "model is not available", "model not available"
    ]
    if any(s in msg for s in model_unavailable_signals):
        return "model_unavailable"

    # Temporary capacity/rate issues: retry same model, then fall back.
    temporary_signals = [
        "503", "unavailable", "high demand", "temporarily unavailable",
        "service unavailable", "resource exhausted", "429",
        "rate limit", "quota exceeded"
    ]
    if any(s in msg for s in temporary_signals):
        return "temporary"

    return "fatal"


def call_gemini_with_retry(api_key: str, prompt: str, status_box=None):
    """
    Stable Free API mode:
    - no models.list() call before conversion;
    - use gemini-3.6-flash directly;
    - retry a temporary 429/503 only once;
    - keep the error message short and teacher-friendly.
    """
    client = genai.Client(api_key=api_key)
    model_name = DEFAULT_MODEL
    last_error = None

    for attempt in range(1, 3):
        try:
            if status_box:
                if attempt == 1:
                    status_box.write(f"正在连接 Gemini：{model_name}")
                else:
                    status_box.write("Gemini 暂时繁忙，正在进行最后一次自动重试…")

            response = client.models.generate_content(
                model=model_name,
                contents=prompt,
                config=types.GenerateContentConfig(
                    response_mime_type="application/json",
                    temperature=0.1,
                ),
            )
            return response, model_name

        except Exception as exc:
            last_error = exc
            error_kind = classify_gemini_error(exc)

            # Invalid key, malformed request, or a model endpoint problem:
            # don't create a long retry loop.
            if error_kind in ("fatal", "model_unavailable"):
                raise

            if attempt == 1:
                if status_box:
                    status_box.write("服务暂时繁忙，3 秒后再试一次…")
                time.sleep(3)

    raise RuntimeError(
        "Gemini 服务目前繁忙。Mini App 已自动重试一次，请稍后再按 Convert。"
    ) from last_error


def build_image_lookup(images):
    return {im["id"]: im for im in images}


def option_label(index: int) -> str:
    """Excel-style letters: A..Z, AA..AZ..."""
    n = index + 1
    letters = ""
    while n:
        n, rem = divmod(n - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def add_image_to_docx(doc, image_bytes, max_width_inches=5.8):
    """Insert image with conservative sizing. Returns True on success."""
    try:
        stream = io.BytesIO(image_bytes)
        doc.add_picture(stream, width=Inches(max_width_inches))
        return True
    except Exception:
        try:
            stream = io.BytesIO(image_bytes)
            doc.add_picture(stream)
            return True
        except Exception:
            return False



def normalize_answer_letters(answer_values, options):
    """Return explicit answers as English option letters; never guess."""
    if answer_values is None:
        return []
    if isinstance(answer_values, str):
        answer_values = [answer_values]

    opts = [str(x).strip() for x in (options or [])]
    result = []

    for raw in answer_values:
        value = str(raw).strip()
        if not value:
            continue

        m = re.fullmatch(r"\(?\s*([A-Za-z]{1,3})\s*\)?[.\s]*", value)
        if m:
            letters = m.group(1).upper()
            if opts:
                idx = 0
                for ch in letters:
                    idx = idx * 26 + (ord(ch) - 64)
                idx -= 1
                if 0 <= idx < len(opts) and letters not in result:
                    result.append(letters)
            elif letters not in result:
                result.append(letters)
            continue

        for idx, opt in enumerate(opts):
            if value == opt:
                letter = option_label(idx)
                if letter not in result:
                    result.append(letter)
                break

    return result


def make_docx(structured: dict, quiz_mode: bool, images):
    """
    Microsoft Forms Quick Import guidance format:
    - Multiple choice: question + vertically stacked A./B./C. options
    - Open text: question only, no answer lines/placeholders
    - Clear blank paragraph between questions
    - NO images/figures embedded in Quick Import Word
    """
    doc = Document()

    normal = doc.styles["Normal"]
    normal.font.name = "Microsoft YaHei"
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), "Microsoft YaHei")
    normal.font.size = Pt(11)

    mapping_rows = []
    embedded_ids = set()  # Intentionally empty for Quick Import Word.

    title = (structured.get("title") or "Microsoft Forms Import").strip()
    if title:
        p = doc.add_paragraph()
        r = p.add_run(title)
        r.bold = True
        r.font.size = Pt(16)

    desc = (structured.get("description") or "").strip()
    if desc:
        doc.add_paragraph(desc)

    global_q_no = 1

    for sec in structured.get("sections", []):
        sec_title = (sec.get("title") or "").strip()
        sec_desc = (sec.get("description") or "").strip()

        # Keep section text simple and vertical.
        if sec_title:
            p = doc.add_paragraph()
            r = p.add_run(sec_title)
            r.bold = True

        if sec_desc:
            doc.add_paragraph(sec_desc)

        for q in sec.get("questions", []):
            original_num = str(q.get("number") or "").strip()
            text = str(q.get("question") or "").strip()
            qtype = str(q.get("output_type") or "open_text").strip()
            image_required = bool(q.get("image_required", False))
            source_page = q.get("source_page")
            image_refs = list(q.get("image_refs") or [])

            # Prefer original number if present; otherwise use sequential numbering.
            number_text = original_num if original_num else str(global_q_no)
            if number_text.endswith("."):
                question_line = f"{number_text} {text}".strip()
            else:
                question_line = f"{number_text}. {text}".strip()

            # Plain vertical question line.
            doc.add_paragraph(question_line)

            opts = [str(x).strip() for x in (q.get("options") or []) if str(x).strip()]

            # Microsoft Forms recommended Multiple choice layout.
            if qtype in ("choice", "multiple_answers"):
                for idx, opt in enumerate(opts):
                    doc.add_paragraph(f"{option_label(idx)}. {opt}")

            # Open text: no answer line, no placeholder, no technical marker.
            elif qtype == "open_text":
                pass

            # Do not embed images in Quick Import Word.
            if image_required:
                mapping_rows.append({
                    "question": question_line,
                    "page": source_page,
                    "image_refs": image_refs,
                    "status": "saved_separately_for_manual_insert",
                })

            # Quiz mode: write explicit answers using English option letters.
            # Example: Answer: A  /  Answer: A, C
            if quiz_mode and qtype in ("choice", "multiple_answers"):
                answer_letters = normalize_answer_letters(q.get("answer") or [], opts)
                if answer_letters:
                    p_ans = doc.add_paragraph()
                    r_ans = p_ans.add_run("Answer: " + ", ".join(answer_letters))
                    r_ans.bold = True

            # Clear separation between questions.
            doc.add_paragraph("")
            global_q_no += 1

    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue(), mapping_rows, embedded_ids


def safe_stem(name: str) -> str:
    stem = Path(name).stem
    return re.sub(r'[\\/:*?"<>|]+', "_", stem)


def safe_docx_filename(name: str) -> str:
    return f"{safe_stem(name)}_Microsoft_Forms_Import_V6_7.docx"


def make_mapping_text(mapping_rows, images):
    lines = [
        "Coach Winnie – Forms Converter V6.7",
        "Image Mapping Report",
        "",
        "Important: Quick Import Word intentionally contains NO images, following Microsoft Forms import guidance.",
        "Use this report and the images folder to add images manually after Quick Import.",
        "",
    ]

    if mapping_rows:
        for i, row in enumerate(mapping_rows, start=1):
            refs = ", ".join(row["image_refs"]) if row["image_refs"] else "(none)"
            lines.append(f"{i}. Question: {row['question']}")
            lines.append(f"   PDF page: {row.get('page') or '(unknown)'}")
            lines.append(f"   Image refs: {refs}")
            lines.append(f"   Status: {row['status']}")
            lines.append("")
    else:
        lines.append("No image-dependent questions were identified.")
        lines.append("")

    if images:
        lines.append("Extracted image inventory:")
        for im in images:
            lines.append(
                f"- {im['id']} -> images/{im['filename']} "
                f"(page {im['page']}, {im['width']}x{im['height']})"
            )
    else:
        lines.append("No extractable raster images were detected in the PDF.")

    return "\n".join(lines)


def make_bundle_zip(pdf_name: str, docx_bytes: bytes, images, mapping_text: str) -> bytes:
    bio = io.BytesIO()
    with zipfile.ZipFile(bio, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr(safe_docx_filename(pdf_name), docx_bytes)
        zf.writestr("image_mapping.txt", mapping_text)
        for im in images:
            zf.writestr(f"images/{im['filename']}", im["bytes"])
    return bio.getvalue()


def build_conversion_report(structured: dict, images, embedded_ids):
    questions = [
        q for sec in structured.get("sections", [])
        for q in sec.get("questions", [])
    ]
    output_types = Counter(q.get("output_type", "open_text") for q in questions)

    return {
        "total": len(questions),
        "output_types": output_types,
        "review_count": sum(1 for q in questions if q.get("review_required")),
        "image_questions": sum(1 for q in questions if q.get("image_required")),
        "extracted_images": len(images),
        "embedded_images": len(embedded_ids),
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
            st.write("读取 PDF 文字")
            pdf_bytes = pdf_file.getvalue()
            form_text = extract_pdf_text(pdf_bytes)

            st.write("提取 PDF 图片")
            images = extract_pdf_images(pdf_bytes)
            st.write(f"检测到 {len(images)} 个可提取候选图片")

            st.write("读取答案文件" if answer_file else "没有答案文件，将不会推测答案")
            answer_text = extract_answer_text(answer_file)

            quiz_mode = output_mode.startswith("Quiz")
            prompt = build_prompt(form_text, answer_text, quiz_mode, images)

            st.write("Gemini 正在按 3 种题型识别：Choice / Multiple-answer Choice / Open text")
            response, model_used = call_gemini_with_retry(
                api_key, prompt, status
            )
            st.write(f"Gemini 连接成功：{model_used}")
            raw = response.text or ""
            structured = json.loads(clean_json_text(raw))
            structured = normalize_for_quick_import(structured)

            st.write("生成 Microsoft Forms Quick Import Word（纯文字 / 选项垂直排列，不嵌入图片）")
            docx_bytes, mapping_rows, embedded_ids = make_docx(
                structured, quiz_mode, images
            )
            mapping_text = make_mapping_text(mapping_rows, images)
            bundle_bytes = make_bundle_zip(
                pdf_file.name, docx_bytes, images, mapping_text
            )
            report = build_conversion_report(structured, images, embedded_ids)
            status.update(label="转换完成 ✅", state="complete")

        st.success(f"完成：共生成 {report['total']} 个 Quick Import 兼容题目。")

        st.subheader("📊 转换报告")
        c1, c2, c3 = st.columns(3)
        c1.metric("Choice", report["output_types"].get("choice", 0))
        c2.metric("Multiple-answer Choice", report["output_types"].get("multiple_answers", 0))
        c3.metric("Open text", report["output_types"].get("open_text", 0))

        c4, c5, c6 = st.columns(3)
        c4.metric("PDF 图片", report["extracted_images"])
        c5.metric("图片题", report["image_questions"])
        c6.metric("Quick Import 内图片", report["embedded_images"])

        notes = []
        multi_count = report["output_types"].get("multiple_answers", 0)
        if multi_count:
            notes.append(
                f"Multiple-answer Choice：{multi_count} 题。"
                "Microsoft Forms Quick Import 后，请为这些题开启 Multiple answers。"
            )
        if report["image_questions"]:
            notes.append(
                "图片题：为了符合 Microsoft Forms Import Guidance，Quick Import Word 不放图片。"
                "原图会保存在 ZIP 的 images 文件夹，请按 image_mapping.txt 在导入后补回。"
            )
        if report["review_count"]:
            notes.append(
                f"人工检查：{report['review_count']} 题因选项不足、结构不清或图片配对不确定，"
                "已尽量安全转换，建议导入后核对。"
            )

        if notes:
            st.warning("\n\n".join(notes))
        else:
            st.info("转换完成，没有检测到需要额外人工检查的题目。")


        st.caption("✅ Quick Import Word 已按 Microsoft Forms Import Guidance 排版；Quiz 模式有答案文件时，选择题会加入英文格式 Answer: A。")

        st.download_button(
            "⬇️ Download Microsoft Forms Word",
            data=docx_bytes,
            file_name=safe_docx_filename(pdf_file.name),
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )

        st.download_button(
            "📦 Download Word + Extracted Images ZIP",
            data=bundle_bytes,
            file_name=f"{safe_stem(pdf_file.name)}_Microsoft_Forms_V6_7_Bundle.zip",
            mime="application/zip",
            use_container_width=True,
        )

        with st.expander("查看 Image Mapping Report"):
            st.code(mapping_text, language=None)

        st.caption(
            "建议流程：Microsoft Forms → Quick Import → 导入 Word → "
            "检查 Choice / Multiple-answer Choice / Open text → "
            "如果图片未自动导入，打开 ZIP 的 images 文件夹，根据 image_mapping.txt 加回对应题目。"
        )

    except json.JSONDecodeError:
        st.error("Gemini 输出无法解析为结构化数据，请重试。若 PDF 很复杂，可分成较小部分转换。")
    except Exception as e:
        msg = str(e)
        if "API_KEY_INVALID" in msg or "API key not valid" in msg or "INVALID_ARGUMENT" in msg:
            st.error("Gemini API Key 无效或请求设置不正确。请到 Google AI Studio 检查 API Key 后再试。")
        elif "404" in msg or "NOT_FOUND" in msg or "no longer available" in msg:
            st.error(
                "当前 Gemini 模型对此 API Key / project 不可用。"
                "请检查 Google AI Studio 中该 project 可使用的模型。"
            )
        elif (
            "RESOURCE_EXHAUSTED" in msg
            or "429" in msg
            or "503" in msg
            or "UNAVAILABLE" in msg
            or "服务目前繁忙" in msg
        ):
            st.error(
                "Gemini 服务目前繁忙或暂时达到 Free API 使用限制。"
                "Mini App 已自动重试一次，请稍后再按 Convert。"
            )
        else:
            st.error(f"转换失败：{e}")

st.divider()
st.markdown("""
<div class="small">
<strong>Microsoft Forms 导入：</strong>
Microsoft Forms → Quick Import → Upload from this device → 选择生成的 Word → Form / Quiz。<br>
<strong>V6.7 Quiz Answer Letter Format：</strong>
Quick Import Word 只保留垂直排列的题目与选项，不嵌入图片；图片会另外保存到 ZIP 与 image_mapping.txt。<br>Quiz 模式若上传答案文件，选择题答案以英文选项字母输出，例如 <strong>Answer: A</strong>；没有答案的题不会推测。<br>
<strong>重要限制：</strong>
Microsoft Forms Quick Import 不保证把 Word 内图片自动转换成题目图片；若图片没有带入，请使用 ZIP 内原图手动补回。<br>
<strong>Privacy：</strong>
上传内容会使用你本次输入的 Gemini API Key 发送到 Google Gemini API 进行转换；
本 App 不会将 API Key 写入 GitHub、文件或数据库。
</div>
""", unsafe_allow_html=True)
