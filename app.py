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
APP_VERSION = "V6.1 Strict Quick Import + Auto Retry"
DEFAULT_MODEL = os.getenv("GEMINI_MODEL", "gemini-3.7-flash")

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
    "V6 使用 Strict Quick Import 格式：Choice 选项统一输出为 A. / B. / C. / D.，Text 不附选项；"
    "同时从 PDF 自动提取可识别的图片，并提供图片 ZIP。"
    "注意：Microsoft Forms Quick Import 可能不会把 Word 内图片自动带入题目；"
    "因此 V5 也会输出独立图片，方便导入后快速补图。"
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
    "🛡️ 若 Gemini 暂时出现 503 / high demand，V6.1 会自动等待重试，并在需要时尝试备用 Flash 模型。\n\n"
    "🔒 API Key 只用于本次页面会话。关闭/刷新页面后请重新输入。"
    "同一个有效的 Gemini API Key 可以重复使用。"
)

with st.expander("V6 转换规则", expanded=False):
    st.markdown("""
- Multiple choice / Dropdown → Choice
- Checkboxes → Choice（Multiple Answers）
- Short answer / Paragraph → Text
- Linear scale → Choice
- Multiple choice grid / Likert-style grid → 每一行拆成独立 Choice
- Checkbox grid → 每一行拆成独立 Multiple Answers Choice
- Matching → 尽量拆成独立 Choice
- Date / Ranking → 先转为 Quick Import 兼容格式，并提示导入后手动调整
- **Image / Map / Chart → 自动从 PDF 提取候选图片，嵌入 Word，并另存图片文件**
- 不推测 PDF 中没有显示的答案
- 不改写、总结或补充原题内容
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
You are the conversion engine for "Coach Winnie – Forms Converter V6.1 Image Extraction Mode".

GOAL
Convert a Google Forms print/PDF into a structure optimized for Microsoft Forms Quick Import Word (.docx).
Use only broadly compatible structures: choice, multiple-answer choice, and open text.
The final DOCX renderer will use STRICT QUICK IMPORT formatting:
- Every choice question MUST have two or more options.
- Choice options will be rendered as A. / B. / C. / D. / E. ...
- Open-text questions MUST have no choice options.
- Grid/Likert rows MUST become independent choice questions.
- Do not put conversion notes, type labels, or technical markers into the visible question text.

NON-NEGOTIABLE CONTENT RULES
1. Use only information present in the supplied form text and optional answer source.
2. Preserve original title, description, section order, question order, wording and options as faithfully as possible.
3. Do not add, delete, summarize, rewrite, correct, explain or supplement the user's question content.
4. Do not infer answers. {answer_instruction}
5. Do not include Google printing UI text as question content.
6. If parsing is uncertain, preserve visible text, choose the safest compatible output, and set review_required=true.

COMPATIBILITY MAPPING
A. Multiple choice / Dropdown -> single_choice.
B. Checkboxes -> multiple_answers.
C. Short answer -> short_answer.
D. Paragraph -> paragraph.
E. Linear scale -> single_choice using original scale values.
F. Multiple choice grid / Likert-style grid:
   - ONE single_choice for EACH ROW.
   - original row label becomes question text.
   - shared column labels become options.
   - If there is an overall grid title, combine ONLY original text as "<grid title> — <row label>".
G. Checkbox grid:
   - ONE multiple_answers question for EACH ROW.
H. Matching:
   - If rows + shared choices are clear, expand each row to single_choice.
   - Otherwise short_answer + review_required=true.
I. Date -> short_answer + manual change recommended.
J. Ranking -> short_answer unless a reliable choice representation is obvious.
K. Image / map / chart dependent questions:
   - Preserve question text.
   - image_required=true.
   - source_page MUST be the PDF page number where the question appears, if visible from PAGE markers.
   - Choose image_refs ONLY from the supplied IMAGE INVENTORY and ONLY when the image is on source_page.
   - If exactly one plausible image is listed on that page, use it.
   - If more than one plausible image is on that page and you cannot know which one belongs to the question, include all plausible same-page image ids and set review_required=true.
   - Never invent an image id.

QUIZ MODE
quiz_mode = {str(quiz_mode).lower()}
If quiz_mode is false, leave answers empty.
If quiz_mode is true, include answers only when explicitly supported by the answer source.

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
          "output_type": "single_choice|multiple_answers|short_answer|paragraph",
          "original_type": "multiple_choice|dropdown|checkboxes|short_answer|paragraph|linear_scale|multiple_choice_grid|checkbox_grid|matching|date|ranking|image_question|unknown",
          "conversion_action": "kept|expanded_grid_row_to_choice|expanded_grid_row_to_multiple_answers|matching_expanded_to_choice|date_to_text_manual_change_recommended|ranking_manual_change_recommended|image_extracted_and_embedded|image_manual_insert_required|fallback_review_required",
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
    """Enforce unambiguous Word structure before rendering."""
    for sec in structured.get("sections", []):
        for q in sec.get("questions", []):
            qtype = q.get("output_type", "short_answer")
            opts = [str(x).strip() for x in (q.get("options") or []) if str(x).strip()]
            if qtype in ("single_choice", "multiple_answers"):
                if len(opts) >= 2:
                    q["options"] = opts
                else:
                    q["output_type"] = "short_answer"
                    q["options"] = []
                    q["review_required"] = True
                    q["conversion_action"] = "fallback_review_required"
            else:
                q["options"] = []
    return structured


def is_temporary_gemini_error(exc: Exception) -> bool:
    msg = str(exc).lower()
    temporary_signals = [
        "503", "unavailable", "high demand", "temporarily unavailable",
        "service unavailable", "resource exhausted", "429",
        "rate limit", "quota exceeded"
    ]
    return any(s in msg for s in temporary_signals)


def call_gemini_with_retry(api_key: str, prompt: str, status_box=None):
    """
    Retry the primary model on temporary 429/503 errors.
    If it is still busy, try conservative Flash fallbacks.
    """
    client = genai.Client(api_key=api_key)

    # Primary first; fallbacks are only used for temporary availability errors.
    models = []
    for m in [DEFAULT_MODEL, "gemini-2.5-flash", "gemini-2.0-flash"]:
        if m and m not in models:
            models.append(m)

    last_error = None
    for model_index, model_name in enumerate(models):
        attempts = 3 if model_index == 0 else 2

        for attempt in range(1, attempts + 1):
            try:
                if status_box:
                    if model_index == 0 and attempt == 1:
                        status_box.write(f"正在连接 Gemini：{model_name}")
                    elif model_index == 0:
                        status_box.write(
                            f"Gemini 当前繁忙，正在自动重试 {attempt-1}/{attempts-1}…"
                        )
                    else:
                        status_box.write(
                            f"主要模型仍繁忙，正在尝试备用模型：{model_name} "
                            f"({attempt}/{attempts})"
                        )

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

                # Invalid key / malformed request should fail immediately.
                if not is_temporary_gemini_error(exc):
                    raise

                # Exponential-ish short backoff: 3s, 6s; fallback gets 3s.
                if attempt < attempts:
                    wait_seconds = 3 * attempt
                    if status_box:
                        status_box.write(
                            f"服务暂时繁忙，{wait_seconds} 秒后再次尝试…"
                        )
                    time.sleep(wait_seconds)
                else:
                    break

    raise RuntimeError(
        "Gemini 服务目前繁忙，自动重试及备用模型均未成功。"
        "请稍后再试。"
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


def make_docx(structured: dict, quiz_mode: bool, images):
    doc = Document()

    normal = doc.styles["Normal"]
    normal.font.name = "Microsoft YaHei"
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), "Microsoft YaHei")
    normal.font.size = Pt(11)

    lookup = build_image_lookup(images)
    embedded_ids = set()
    mapping_rows = []

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
            source_page = q.get("source_page")
            image_refs = [x for x in (q.get("image_refs") or []) if x in lookup]

            prefix = f"{num}. " if num and not str(num).rstrip().endswith(".") else (f"{num} " if num else "")

            # STRICT QUICK IMPORT: visible question text contains only original content.
            # No "(Multiple Answers)", required asterisk, type marker, or conversion note.
            p = doc.add_paragraph()
            rr = p.add_run(f"{prefix}{text}".strip())
            rr.bold = True

            if image_required:
                if image_refs:
                    if len(image_refs) > 1:
                        note = doc.add_paragraph()
                        nr = note.add_run(
                            f"[Original image candidates extracted from PDF page {source_page}; verify the correct image after import]"
                        )
                        nr.italic = True

                    inserted = []
                    for ref in image_refs:
                        im = lookup[ref]
                        if add_image_to_docx(doc, im["bytes"]):
                            embedded_ids.add(ref)
                            inserted.append(ref)
                            cap = doc.add_paragraph(f"[{im['filename']}]")
                            cap.runs[0].italic = True

                    mapping_rows.append({
                        "question": f"{prefix}{text}".strip(),
                        "page": source_page,
                        "image_refs": inserted,
                        "status": "embedded" if inserted else "extract_failed",
                    })
                else:
                    # Keep technical image instructions OUT of the import DOCX.
                    # They belong only in image_mapping.txt / conversion report.
                    mapping_rows.append({
                        "question": f"{prefix}{text}".strip(),
                        "page": source_page,
                        "image_refs": [],
                        "status": "manual",
                    })

            opts = q.get("options") or []

            # STRICT QUICK IMPORT:
            # Choice = visibly lettered options. Text = no options at all.
            if qtype in ("single_choice", "multiple_answers"):
                for idx, opt in enumerate(opts):
                    doc.add_paragraph(f"{option_label(idx)}. {str(opt).strip()}")
            elif qtype in ("short_answer", "paragraph"):
                # Intentionally no answer lines/placeholders; Quick Import should see open text.
                pass

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
    return bio.getvalue(), mapping_rows, embedded_ids


def safe_stem(name: str) -> str:
    stem = Path(name).stem
    return re.sub(r'[\\/:*?"<>|]+', "_", stem)


def safe_docx_filename(name: str) -> str:
    return f"{safe_stem(name)}_Microsoft_Forms_Import_V6_1.docx"


def make_mapping_text(mapping_rows, images):
    lines = [
        "Coach Winnie – Forms Converter V6.1",
        "Image Mapping Report",
        "",
        "Important: Microsoft Forms Quick Import may not automatically import Word images.",
        "Use this report and the images folder to add images manually after Quick Import when needed.",
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
    output_types = Counter(q.get("output_type", "unknown") for q in questions)
    actions = Counter(q.get("conversion_action", "kept") for q in questions)

    return {
        "total": len(questions),
        "output_types": output_types,
        "review_count": sum(1 for q in questions if q.get("review_required")),
        "image_questions": sum(1 for q in questions if q.get("image_required")),
        "extracted_images": len(images),
        "embedded_images": len(embedded_ids),
        "manual_date": actions.get("date_to_text_manual_change_recommended", 0),
        "manual_ranking": actions.get("ranking_manual_change_recommended", 0),
        "expanded_choice": actions.get("expanded_grid_row_to_choice", 0),
        "expanded_multi": actions.get("expanded_grid_row_to_multiple_answers", 0),
        "matching_expanded": actions.get("matching_expanded_to_choice", 0),
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

            st.write("Gemini 正在识别题型、页码与图片题")
            response, model_used = call_gemini_with_retry(
                api_key, prompt, status
            )
            st.write(f"Gemini 连接成功：{model_used}")
            raw = response.text or ""
            structured = json.loads(clean_json_text(raw))
            structured = normalize_for_quick_import(structured)

            st.write("生成 Strict Quick Import Word 并嵌入图片")
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
        c1.metric("Choice", report["output_types"].get("single_choice", 0))
        c2.metric("Multiple Answers", report["output_types"].get("multiple_answers", 0))
        c3.metric(
            "Text",
            report["output_types"].get("short_answer", 0)
            + report["output_types"].get("paragraph", 0),
        )

        c4, c5, c6 = st.columns(3)
        c4.metric("PDF 图片", report["extracted_images"])
        c5.metric("图片题", report["image_questions"])
        c6.metric("已嵌入 Word", report["embedded_images"])

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
            notes.append(f"Ranking：{report['manual_ranking']} 题需要导入后手动调整。")
        if report["image_questions"]:
            notes.append(
                "图片题：已尽量把 PDF 原图嵌入 Word，并同时放入 ZIP 的 images 文件夹。"
                "Microsoft Forms Quick Import 可能不会自动带入这些图片，导入后请按 image_mapping.txt 补图。"
            )
        if report["review_count"]:
            notes.append(f"人工检查：{report['review_count']} 题解析不确定，建议导入后核对。")

        if notes:
            st.warning("\n\n".join(notes))
        else:
            st.info("没有检测到需要额外人工转换的特殊题型。")

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
            file_name=f"{safe_stem(pdf_file.name)}_Microsoft_Forms_V6_1_Bundle.zip",
            mime="application/zip",
            use_container_width=True,
        )

        with st.expander("查看 Image Mapping Report"):
            st.code(mapping_text, language=None)

        st.caption(
            "建议流程：Microsoft Forms → Quick Import → 导入 Word → "
            "检查 Choice / Multiple Answers / Date / Ranking → "
            "如果图片未自动导入，打开 ZIP 的 images 文件夹，根据 image_mapping.txt 加回对应题目。"
        )

    except json.JSONDecodeError:
        st.error("Gemini 输出无法解析为结构化数据，请重试。若 PDF 很复杂，可分成较小部分转换。")
    except Exception as e:
        msg = str(e)
        if "API_KEY_INVALID" in msg or "API key not valid" in msg or "INVALID_ARGUMENT" in msg:
            st.error("Gemini API Key 无效或请求设置不正确。请到 Google AI Studio 检查 API Key 后再试。")
        elif (
            "RESOURCE_EXHAUSTED" in msg
            or "429" in msg
            or "503" in msg
            or "UNAVAILABLE" in msg
            or "服务目前繁忙" in msg
        ):
            st.error(
                "Gemini 服务目前繁忙或暂时达到使用限制。"
                "Mini App 已自动重试并尝试备用模型，但暂时仍无法完成。"
                "请稍后再按 Convert。"
            )
        else:
            st.error(f"转换失败：{e}")

st.divider()
st.markdown("""
<div class="small">
<strong>Microsoft Forms 导入：</strong>
Microsoft Forms → Quick Import → Upload from this device → 选择生成的 Word → Form / Quiz。<br>
<strong>V6.1 Strict Quick Import + Auto Retry：</strong>
会从 PDF 提取可识别的原图并嵌入 Word，同时输出独立图片 ZIP 与 image_mapping.txt。<br>
<strong>重要限制：</strong>
Microsoft Forms Quick Import 不保证把 Word 内图片自动转换成题目图片；若图片没有带入，请使用 ZIP 内原图手动补回。<br>
<strong>Privacy：</strong>
上传内容会使用你本次输入的 Gemini API Key 发送到 Google Gemini API 进行转换；
本 App 不会将 API Key 写入 GitHub、文件或数据库。
</div>
""", unsafe_allow_html=True)
