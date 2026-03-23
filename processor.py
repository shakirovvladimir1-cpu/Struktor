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
BATCH_SIZE = 80  # paragraphs per Gemini call


def iter_paragraphs(doc: Document):
    """Yield all paragraphs from document body and table cells, in order."""
    for block in doc.element.body:
        if block.tag == qn("w:p"):
            yield Paragraph(block, doc)
        elif block.tag == qn("w:tbl"):
            for p_elem in block.iter(qn("w:p")):
                yield Paragraph(p_elem, doc)


def collect_paragraphs(doc: Document) -> list[dict]:
    """Return list of {id, para, text} for all non-empty paragraphs."""
    result = []
    for idx, para in enumerate(iter_paragraphs(doc)):
        text = para.text.strip()
        if text:
            result.append({"id": idx, "para": para, "text": text})
    return result


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


def parse_gemini_response(raw: str) -> list[dict]:
    """Extract JSON array from Gemini response, stripping markdown fences."""
    text = raw.strip()
    text = re.sub(r"^```(?:json)?\s*", "", text)
    text = re.sub(r"\s*```$", "", text)
    return json.loads(text.strip())


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

    # Process in batches
    cleaned_map: dict[int, list[str]] = {}
    for i in range(0, len(paragraphs), BATCH_SIZE):
        batch = paragraphs[i: i + BATCH_SIZE]
        batch_result = call_gemini(client, batch)
        cleaned_map.update(batch_result)
        logger.info(
            f"Batch {i // BATCH_SIZE + 1}/{(len(paragraphs) - 1) // BATCH_SIZE + 1} done"
        )

    # Apply changes: replace text and insert new paragraphs where needed
    changed = 0
    inserted = 0
    for p in paragraphs:
        new_paragraphs = cleaned_map.get(p["id"])
        if not new_paragraphs:
            continue

        if len(new_paragraphs) == 1:
            # Simple text replacement
            new_text = new_paragraphs[0]
            if new_text != p["text"]:
                set_para_text(p["para"], new_text)
                changed += 1
        else:
            # Multiple paragraphs: insert headers before the original, then update original
            headers = new_paragraphs[:-1]   # first N-1 lines are headers
            spec_text = new_paragraphs[-1]  # last line is the cleaned spec

            insert_paragraphs_before(p["para"], headers)
            set_para_text(p["para"], spec_text)

            inserted += len(headers)
            changed += 1

    logger.info(
        f"Changed {changed} paragraphs, inserted {inserted} new header paragraphs"
    )
    doc.save(output_path)
