import json
import logging
import re

from docx import Document
from docx.text.paragraph import Paragraph
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from google import genai
from google.genai import types

from prompts import SYSTEM_PROMPT

logger = logging.getLogger(__name__)

GEMINI_MODEL = "gemini-2.5-flash"
BATCH_SIZE = 20   # paragraphs per Gemini call
MAX_CHARS_PER_BATCH = 12000  # max total text chars per Gemini call


def iter_paragraphs(doc: Document):
    """Yield all paragraphs from document body and table cells, in order."""
    for block in doc.element.body:
        if block.tag == qn("w:p"):
            yield Paragraph(block, doc)
        elif block.tag == qn("w:tbl"):
            for p_elem in block.iter(qn("w:p")):
                yield Paragraph(p_elem, doc)


def split_para_by_linebreaks(para: Paragraph) -> list[str]:
    """Split paragraph text by <w:br/> line breaks, return non-empty chunks."""
    chunks = []
    current = []
    for elem in para._element.iter():
        tag = elem.tag.split("}")[-1] if "}" in elem.tag else elem.tag
        if tag == "t" and elem.text:
            current.append(elem.text)
        elif tag == "br":
            text = "".join(current).strip()
            if text:
                chunks.append(text)
            current = []
    text = "".join(current).strip()
    if text:
        chunks.append(text)
    return chunks if chunks else ([para.text.strip()] if para.text.strip() else [])


CHUNK_SIZE = 5000  # max chars per virtual chunk when splitting giant paragraphs


def split_text_into_chunks(text: str, max_chars: int = CHUNK_SIZE) -> list[str]:
    """Split a large text into chunks of up to max_chars, breaking at sentence/comma boundaries."""
    if len(text) <= max_chars:
        return [text]
    chunks = []
    while text:
        if len(text) <= max_chars:
            chunks.append(text)
            break
        # Find best break point (comma, period, semicolon) near max_chars
        candidate = text[:max_chars]
        for sep in (",", ";", ".", " "):
            idx = candidate.rfind(sep)
            if idx > max_chars // 2:
                candidate = text[:idx + 1]
                break
        chunks.append(candidate.strip())
        text = text[len(candidate):].strip()
    return [c for c in chunks if c]


def collect_paragraphs(doc: Document) -> list[dict]:
    """Return list of {id, para, text} for all non-empty paragraphs.
    Giant single-paragraph documents are split into logical chunks."""
    raw = []
    for para in iter_paragraphs(doc):
        text = para.text.strip()
        if not text:
            continue
        if len(text) > CHUNK_SIZE:
            # Try line-break split first
            lines = split_para_by_linebreaks(para)
            if len(lines) > 1:
                for line in lines:
                    raw.append({"para": para, "text": line, "virtual": True})
                continue
            # Fall back to character-based chunking
            chunks = split_text_into_chunks(text)
            if len(chunks) > 1:
                for chunk in chunks:
                    raw.append({"para": para, "text": chunk, "virtual": True})
                continue
        raw.append({"para": para, "text": text, "virtual": False})

    return [{"id": i, **r} for i, r in enumerate(raw)]


def set_para_text(para: Paragraph, new_text: str):
    """Replace paragraph text while preserving paragraph-level style."""
    runs = para.runs
    if not runs:
        para.add_run(new_text)
        return
    runs[0].text = new_text
    for run in runs[1:]:
        run.text = ""


def make_paragraph_element(text: str):
    """Create a minimal w:p XML element with given text."""
    new_p = OxmlElement("w:p")
    new_r = OxmlElement("w:r")
    new_t = OxmlElement("w:t")
    new_t.text = text
    new_r.append(new_t)
    new_p.append(new_r)
    return new_p


def insert_paragraphs_before(para: Paragraph, texts: list[str]):
    """Insert multiple paragraphs immediately before `para`, in order."""
    ref_elem = para._element
    for text in texts:
        new_p = make_paragraph_element(text)
        ref_elem.addprevious(new_p)
        # Each subsequent insert also goes right before the original para,
        # which pushes the previous inserts further back — maintaining order.


def fix_json_escapes(text: str) -> str:
    """Fix invalid backslash escapes inside JSON string values."""
    # Valid JSON escape chars: " \ / b f n r t u
    valid = set('"\\\/bfnrtu')
    result = []
    i = 0
    while i < len(text):
        ch = text[i]
        if ch == '\\' and i + 1 < len(text):
            next_ch = text[i + 1]
            if next_ch in valid:
                result.append(ch)
                result.append(next_ch)
                i += 2
            else:
                # Invalid escape — double the backslash
                result.append('\\\\')
                i += 1
        else:
            result.append(ch)
            i += 1
    return "".join(result)


def parse_gemini_response(raw: str) -> list[dict]:
    """Extract JSON array from Gemini response, stripping markdown fences."""
    text = raw.strip()
    text = re.sub(r"^```(?:json)?\s*", "", text)
    text = re.sub(r"\s*```$", "", text)
    text = text.strip()
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        text = fix_json_escapes(text)
        return json.loads(text)


def call_gemini(client: genai.Client, batch: list[dict]) -> dict[int, list[str]]:
    """Send a batch to Gemini, return {id: [list_of_paragraphs]} mapping."""
    input_data = [{"id": p["id"], "text": p["text"]} for p in batch]
    prompt = (
        SYSTEM_PROMPT
        + "\n\nВходные данные:\n"
        + json.dumps(input_data, ensure_ascii=False, indent=2)
    )
    response = client.models.generate_content(
        model=GEMINI_MODEL,
        contents=prompt,
        config=types.GenerateContentConfig(
            response_mime_type="application/json",
        ),
    )
    cleaned = parse_gemini_response(response.text)

    result = {}
    for item in cleaned:
        paragraphs = item.get("paragraphs")
        if paragraphs is None:
            # Fallback: old format with "text" key
            paragraphs = [item.get("text", "")]
        result[item["id"]] = paragraphs
    return result


def process_docx(input_path: str, output_path: str, gemini_api_key: str):
    """
    Read DOCX, clean and restructure all paragraphs with Gemini 2.5 Flash,
    save result to output_path.

    For item paragraphs (numbered goods), inserts 5 structured header lines
    (Наименование, Страна, Завод, Год, Гарантийный срок) before the spec text.
    """
    client = genai.Client(api_key=gemini_api_key)

    doc = Document(input_path)
    paragraphs = collect_paragraphs(doc)

    if not paragraphs:
        raise ValueError("Документ пустой или не содержит текстовых абзацев.")

    logger.info(f"Total paragraphs to process: {len(paragraphs)}")

    # Process in batches, respecting both paragraph count and char limits
    cleaned_map: dict[int, list[str]] = {}
    batch: list[dict] = []
    batch_chars = 0
    batch_num = 0

    def flush_batch(b):
        nonlocal batch_num
        if not b:
            return
        batch_num += 1
        result = call_gemini(client, b)
        cleaned_map.update(result)
        logger.info(f"Batch {batch_num} done ({len(b)} paragraphs)")

    for p in paragraphs:
        p_len = len(p["text"])
        # If single paragraph exceeds limit — send it alone
        if p_len > MAX_CHARS_PER_BATCH:
            flush_batch(batch)
            batch, batch_chars = [], 0
            flush_batch([p])
            continue
        # If adding this paragraph would exceed limits — flush first
        if batch and (len(batch) >= BATCH_SIZE or batch_chars + p_len > MAX_CHARS_PER_BATCH):
            flush_batch(batch)
            batch, batch_chars = [], 0
        batch.append(p)
        batch_chars += p_len

    flush_batch(batch)

    has_virtual = any(p.get("virtual") for p in paragraphs)

    if has_virtual:
        # Document uses line-breaks instead of paragraphs — build new doc from scratch
        out_doc = Document()
        # Copy core properties (margins etc) from original if possible
        for p in paragraphs:
            lines = cleaned_map.get(p["id"], [p["text"]])
            for line in lines:
                out_doc.add_paragraph(line)
        out_doc.save(output_path)
        logger.info(f"Virtual-paragraph doc: wrote {len(paragraphs)} lines as real paragraphs")
        return

    # Normal doc: Apply changes in-place
    changed = 0
    inserted = 0
    for p in paragraphs:
        new_paragraphs = cleaned_map.get(p["id"])
        if not new_paragraphs:
            continue

        if len(new_paragraphs) == 1:
            new_text = new_paragraphs[0]
            if new_text != p["text"]:
                set_para_text(p["para"], new_text)
                changed += 1
        else:
            headers = new_paragraphs[:-1]
            spec_text = new_paragraphs[-1]
            insert_paragraphs_before(p["para"], headers)
            set_para_text(p["para"], spec_text)
            inserted += len(headers)
            changed += 1

    logger.info(
        f"Changed {changed} paragraphs, inserted {inserted} new header paragraphs"
    )
    doc.save(output_path)
