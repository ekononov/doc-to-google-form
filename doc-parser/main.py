"""
FastAPI: POST /parse (multipart, field "file" = .docx) -> JSON with tests.

Supported question types: single (one correct answer) and multiple (several).
Languages of text and markers: English, Russian, Ukrainian (can be mixed).

How the structure is detected (heuristic, not tied to one template):
  Question: "1. ...", "1) ...", "Question 1. ...", a top-level Word list item,
            or simply a paragraph without a marker.
  Option:   "a) ...", "A. ...", "(b) ...", "- ...", "• ...", or a nested Word list item.
  Correct answer (in order of priority):
    1) a line after the options: "Answer: a, c" / "Correct answer: 2,3" / option text;
    2) a mark inside the option: "*", "+", "✓" at the start/end, "=" before the label
       (e.g. "=c) text"), or "(correct)" at the end;
    3) formatting: bold, underline, highlight or font color on some of the options
       (if all options are formatted the same way, formatting is ignored).
  Points: "(2 points)", "[3 pts]" in the question text or a separate line "Points: 2".
          If absent, the default_points parameter is used (default 1).
  Service lines (en/ru/uk) are removed together with surrounding blank lines and decoration:
    - a "question text" header before the question (the question itself follows, or is inline);
    - an "options" header before the options (unmarked lines after it are treated as options);
    - a "comment" block after the options (skipped until the next question starts).
  If there are no markers at all: blocks separated by an empty line
  (first line = question, the rest = options).
"""
import io
import re
from typing import Any, Optional

from docx import Document
from fastapi import FastAPI, File, HTTPException, Query, UploadFile

app = FastAPI(title="DOCX tests -> JSON")

L = r"A-Za-zА-Яа-яЁёІіЇїЄєҐґ"  # en + ru + uk
MARKS = "*+✓✔"
PREFIX_MARKS = MARKS + "="  # "=" is accepted only in front of an option label, e.g. "=c) text"

RE_Q_WORD = re.compile(
    r"^\s*(?:вопрос|питання|запитання|задание|завдання|question|task|q)\s*№?\s*\d+\s*[.):]?\s*",
    re.I,
)
RE_Q_NUM = re.compile(r"^\s*\d+\s*[.)]\s+")
RE_OPT = re.compile(
    rf"^\s*(?:[{PREFIX_MARKS}]\s*)?(?:\(?([{L}])\)\s*|([{L}])\.\s+|[-–—•●▪○]\s+)"
)
RE_ANS = re.compile(
    r"^\s*(?:(?:(?:правильн\w*|вірн\w*|верн\w*|correct)\s+)?"
    r"(?:ответ\w*|відповід\w*|answers?|key)|correct)\s*[:：=-]\s*(.+)$",
    re.I,
)
RE_PTS = re.compile(
    r"[(\[]\s*(\d+(?:[.,]\d+)?)\s*(?:бал\w*|б\.?|points?|pts?|marks?)\s*[)\]]", re.I
)
RE_PTS_LINE = re.compile(
    r"^\s*(?:бал\w*|points?|pts?|marks?)\s*[:=]\s*(\d+(?:[.,]\d+)?)\s*$", re.I
)
RE_TRAIL = re.compile(
    rf"(?:\s*[{MARKS}]\s*|\s*\((?:верн\w*|вірн\w*|правильн\w*|correct|right)\)\s*)$",
    re.I,
)

# Conjunctions (en/ru/uk) in an answer line, e.g. "a and c"; kept if the token is an option label
CONJ = {"и", "і", "й", "та", "and", "or", "або", "или"}
# Cyrillic letters that look like Latin ones: a Cyrillic label and its Latin look-alike are treated as equal
_LOOK = str.maketrans("асеорхкмтнві", "aceopxkmthbi")


def _norm(s: str) -> str:
    return s.lower().translate(_LOOK)


# Service lines (en/ru/uk): "question text" header, "options" header, "comment" block header.
# Separators/decoration around them (colons, dashes, asterisks, underscores, "#") are ignored.
_SEP = r"[\s:：.\-–—_*#]*"
RE_H_QUESTION = re.compile(
    rf"^{_SEP}(?:текст\s+(?:вопроса|вопросу|питання|запитання)"
    r"|(?:question|task)\s+text|text\s+of\s+(?:the\s+)?(?:question|task)"
    r"|вопрос|питання|запитання|question)"
    rf"(?:{_SEP}$|\s*[:：]\s*(?P<rest>\S.*)$)",
    re.I,
)
RE_H_OPTIONS = re.compile(
    r"^" + _SEP + r"(?:варіанти\s+відповід\w*|варианты\s+ответ\w*|варіанти|варианты"
    r"|відповіді|ответы|answer\s+options|answer\s+choices|possible\s+answers"
    r"|options|choices|answers)" + _SEP + r"$",
    re.I,
)
RE_H_COMMENT = re.compile(
    r"^" + _SEP + r"(?:коментар\w*|комментари\w*|comments?|explanations?|пояснення|пояснение"
    r"|notes?|feedback|rationale)(?:" + _SEP + r"$|\s*[:：].*$)",
    re.I,
)

ZERO_WIDTH = re.compile("[​‌‍⁠﻿]")


def _tidy(s: str) -> str:
    """Remove zero-width chars, collapse spaces, drop empty lines and leading decoration."""
    s = ZERO_WIDTH.sub("", s)
    lines = [re.sub(r"[ \t ]+", " ", l).strip() for l in s.split("\n")]
    s = "\n".join(l for l in lines if l)
    return s.lstrip(":：*#_ \n").strip()



# ---------- helpers ----------

def _num(s: str) -> float:
    v = float(s.replace(",", "."))
    return int(v) if v.is_integer() else v


def _list_level(p) -> Optional[int]:
    pPr = p._p.pPr
    if pPr is not None and pPr.numPr is not None:
        il = pPr.numPr.ilvl
        return il.val if il is not None else 0
    if p.style is not None and (p.style.name or "").lower().startswith("list"):
        return 0
    return None


def _format_ratios(p) -> dict:
    """Share of the paragraph's characters that are bold/underlined/highlighted/colored."""
    total = 0
    cnt = {"bold": 0, "underline": 0, "highlight": 0, "color": 0}
    style_bold = bool(p.style is not None and p.style.font.bold)
    for r in p.runs:
        t = r.text.strip()
        if not t:
            continue
        n = len(t)
        total += n
        f = r.font
        bold = r.bold if r.bold is not None else style_bold
        if bold:
            cnt["bold"] += n
        if f.underline:
            cnt["underline"] += n
        if f.highlight_color is not None:
            cnt["highlight"] += n
        try:
            rgb = f.color.rgb if f.color is not None else None
        except Exception:
            rgb = None
        if rgb is not None and str(rgb).upper() not in ("000000", "AUTO"):
            cnt["color"] += n
    return {k: (v / total if total else 0.0) for k, v in cnt.items()}


def _make_option(p, text: str) -> dict:
    flag = bool(re.match(rf"^\s*[{PREFIX_MARKS}]", text))
    label = None
    m = RE_OPT.match(text)
    if m:
        label = m.group(1) or m.group(2)
        text = text[m.end():]
    m2 = re.match(rf"^\s*[{MARKS}]\s*", text)
    if m2:
        flag = True
        text = text[m2.end():]
    m3 = RE_TRAIL.search(text)
    if m3:
        flag = True
        text = text[: m3.start()]
    return {
        "label": _norm(label) if label else None,
        "text": _tidy(text),
        "flag": flag,
        "fmt": _format_ratios(p),
    }


def _from_answer(raw: str, opts: list) -> Optional[list]:
    raw = raw.strip()
    # the answer is the full text of an option
    for i, o in enumerate(opts):
        if o["text"] and o["text"].lower() == raw.lower().rstrip("."):
            res = [False] * len(opts)
            res[i] = True
            return res
    labels = [o["label"] for o in opts]
    tokens = [t for t in re.split(r"[,;/\s]+", raw) if t]
    keys = []
    for t in tokens:
        t = _norm(t.strip(".)("))
        if not t:
            continue
        if t in CONJ and t not in labels:
            continue
        if t.isalpha() and len(t) > 1 and not any(t == (l or "") for l in labels):
            keys.extend(list(t))  # "ac" -> a, c
        else:
            keys.append(t)
    res = [False] * len(opts)
    for k in keys:
        if k in labels:
            res[labels.index(k)] = True
        elif k.isdigit() and 1 <= int(k) <= len(opts):
            res[int(k) - 1] = True
    return res if any(res) else None


def _from_format(opts: list) -> list:
    for kind in ("bold", "underline", "highlight", "color"):
        flags = [o["fmt"][kind] >= 0.5 for o in opts]
        if 0 < sum(flags) < len(flags):
            return flags
    return [False] * len(opts)


def _finalize(q: dict, default_points: float) -> dict:
    text = q["text"]
    points = q["points"]
    m = RE_PTS.search(text)
    if m:
        if points is None:
            points = _num(m.group(1))
        text = text[: m.start()] + text[m.end():]
    text = _tidy(text)
    opts = q["options"]

    correct = _from_answer(q["answer"], opts) if q["answer"] else None
    if not correct:
        if any(o["flag"] for o in opts):
            correct = [o["flag"] for o in opts]
        else:
            correct = _from_format(opts)

    return {
        "text": text,
        "type": "multiple" if sum(correct) > 1 else "single",
        "points": points if points is not None else default_points,
        "options": [{"text": o["text"], "correct": c} for o, c in zip(opts, correct)],
    }


def _new_q(text: str) -> dict:
    return {"text": text, "options": [], "answer": None, "points": None}


# ---------- parsing ----------

def _is_question_start(text: str, count: int) -> bool:
    """Inside a comment block: does this line start the next question?"""
    if RE_Q_WORD.match(text):
        return True
    if RE_Q_NUM.match(text):
        # a bare "N." line counts only if N is the expected next question number,
        # so numbered lists inside a comment are not mistaken for questions
        return int(re.match(r"\s*(\d+)", text).group(1)) == count + 1
    return False


def _service_line(text: str):
    """Classify a service line: returns (kind, rest) or None.

    kind is "question" (rest = inline question text, may be empty), "options" or "comment".
    """
    m = RE_H_QUESTION.match(text)
    if m:
        return "question", m.group("rest") or ""
    if RE_H_OPTIONS.match(text):
        return "options", ""
    if RE_H_COMMENT.match(text):
        return "comment", ""
    return None


def _parse_marked(items: list) -> list:
    levels = [_list_level(p) for p, _ in items]
    has_sub = any(l is not None and l > 0 for l in levels)

    questions: list = []
    cur: Optional[dict] = None
    pending_q = False   # "question text" header seen, the question itself comes next
    opts_mode = False   # "options" header seen: unmarked lines are options
    in_comment = False  # inside a comment block: lines are skipped

    for (p, text), lvl in zip(items, levels):
        svc = _service_line(text)

        if svc and svc[0] == "comment":
            in_comment = True
            continue
        if in_comment:
            if (svc and svc[0] == "question") or _is_question_start(text, len(questions)):
                in_comment = False
            else:
                continue

        if svc:
            kind, rest = svc
            if kind == "question":
                opts_mode = False
                if _tidy(rest):
                    cur = _new_q(rest)
                    questions.append(cur)
                    pending_q = False
                else:
                    pending_q = True
            elif cur is not None and not pending_q:
                opts_mode = True
            continue

        q_m = RE_Q_WORD.match(text) or RE_Q_NUM.match(text)

        if pending_q:
            cur = _new_q(text[q_m.end():] if q_m else text)
            questions.append(cur)
            pending_q = False
            opts_mode = False
            continue

        m = RE_ANS.match(text)
        if m and cur and cur["options"]:
            cur["answer"] = m.group(1)
            continue
        m = RE_PTS_LINE.match(text)
        if m and cur:
            cur["points"] = _num(m.group(1))
            continue

        opt_m = RE_OPT.match(text)

        if q_m:
            kind = "q"
        elif opt_m and cur:
            kind = "o"
        elif lvl is not None:
            if has_sub:
                kind = "o" if lvl > 0 and cur else "q"
            else:
                kind = "o" if cur else "q"
        elif opts_mode and cur:
            kind = "o"
        else:
            kind = "c" if cur and not cur["options"] else "q"

        if kind == "q":
            cur = _new_q(text[q_m.end():] if q_m else text)
            questions.append(cur)
            opts_mode = False
        elif kind == "o":
            cur["options"].append(_make_option(p, text))
        else:  # continuation of the question text
            cur["text"] += "\n" + text

    return [q for q in questions if q["options"]]


def _parse_blocks(items_with_empty: list) -> list:
    blocks, block = [], []
    for p, text in items_with_empty:
        if text:
            block.append((p, text))
        elif block:
            blocks.append(block)
            block = []
    if block:
        blocks.append(block)

    questions = []
    for raw_block in blocks:
        # drop service lines; everything after a comment header is skipped
        b = []
        for p, text in raw_block:
            svc = _service_line(text)
            if svc is None:
                b.append((p, text))
            elif svc[0] == "comment":
                break
            elif svc[0] == "question" and _tidy(svc[1]):
                b.append((p, svc[1]))
        if len(b) < 3:
            continue
        first = b[0][1]
        q_m = RE_Q_WORD.match(first) or RE_Q_NUM.match(first)
        q = _new_q(first[q_m.end():] if q_m else first)
        for p, text in b[1:]:
            m = RE_ANS.match(text)
            if m:
                q["answer"] = m.group(1)
                continue
            m = RE_PTS_LINE.match(text)
            if m:
                q["points"] = _num(m.group(1))
                continue
            q["options"].append(_make_option(p, text))
        if q["options"]:
            questions.append(q)
    return questions


def parse_docx(doc, default_points: float) -> list[dict[str, Any]]:
    # one item per visible line; soft line breaks inside a paragraph become separate lines
    all_items = []
    for p in doc.paragraphs:
        lines = [ZERO_WIDTH.sub("", l).strip() for l in p.text.split("\n")]
        lines = [l for l in lines if l]
        if lines:
            all_items.extend((p, l) for l in lines)
        else:
            all_items.append((p, ""))
    items = [(p, t) for p, t in all_items if t]

    raw = _parse_marked(items)
    if not any(len(q["options"]) >= 2 for q in raw):
        raw = _parse_blocks(all_items)

    result = []
    for i, q in enumerate(raw, 1):
        fq = _finalize(q, default_points)
        fq["number"] = i
        result.append(fq)
    return [
        {k: q[k] for k in ("number", "text", "type", "points", "options")}
        for q in result
    ]


# ---------- API ----------

@app.get("/health")
async def health():
    return {"status": "ok"}


@app.post("/parse")
async def convert(
    file: UploadFile = File(...),
    default_points: float = Query(1, ge=0, description="Default points per question"),
):
    if not (file.filename or "").lower().endswith(".docx"):
        raise HTTPException(status_code=400, detail="A .docx file is required")
    data = await file.read()
    try:
        doc = Document(io.BytesIO(data))
    except Exception:
        raise HTTPException(status_code=400, detail="Failed to read the docx file")
    return {"questions": parse_docx(doc, default_points)}