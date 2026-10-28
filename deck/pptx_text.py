"""Round-trip the text of a .pptx through an editable Markdown file.

  export:  python pptx_text.py export deck.pptx deck.txt.md
  apply:   python pptx_text.py apply  deck.pptx deck.txt.md out.pptx

Every text frame, table cell and speaker-notes block gets a tag ([#shape_id], [#shape_id r1c2],
[notes]); one line = one paragraph; **bold** marks bold runs. Apply only rewrites blocks whose text
changed, reusing the original run formatting (bold runs copy the first bold run, others the first
regular run), so fonts, colours and positions stay as they are.
"""
import copy
import re
import sys

from pptx import Presentation

TAG = re.compile(r"^\[(#\d+(?: r\d+c\d+)?|notes)\]\s?(.*)$")
SLIDE = re.compile(r"^=== SLIDE (\d+)")


def para_md(p):
    """Paragraph as Markdown; adjacent runs with the same weight are merged."""
    groups = []
    for r in p.runs:
        t = r.text.replace("\n", " ").replace("\x0b", " ")
        # runs without their own weight inherit it; follow the neighbouring run
        w = r.font.bold if r.font.bold is not None else (groups[-1][0] if groups else False)
        b = bool(w) and bool(t.strip())
        if groups and (groups[-1][0] == b or not t.strip()):
            groups[-1][1] += t
        else:
            groups.append([b, t])
    return "".join(f"**{t}**" if b else t for b, t in groups)


def norm(paras):
    paras = [x.rstrip() for x in paras]
    while paras and not paras[-1].strip():
        paras.pop()
    return paras


def frames(slide):
    """Yields (tag, text_frame) for every text-bearing frame on the slide, in reading order."""
    for sh in slide.shapes:
        if sh.has_text_frame and "".join(para_md(p) for p in sh.text_frame.paragraphs).strip():
            yield f"#{sh.shape_id}", sh.text_frame
        if sh.has_table:
            for ri, row in enumerate(sh.table.rows, 1):
                for ci, cell in enumerate(row.cells, 1):
                    if cell.text_frame.text.strip():
                        yield f"#{sh.shape_id} r{ri}c{ci}", cell.text_frame
    if slide.has_notes_slide and slide.notes_slide.notes_text_frame.text.strip():
        yield "notes", slide.notes_slide.notes_text_frame


def export(src, out):
    prs = Presentation(src)
    lines = ["# Slide text — edit freely after the [tags]; keep the tags and the === SLIDE lines.",
             "# One line = one paragraph/bullet. **bold** = bold text. Empty a block to clear that box.",
             "# Requests that text cannot express: add a line starting with  >> NOTE:", ""]
    for i, s in enumerate(prs.slides, 1):
        title = s.shapes.title.text.replace("\n", " ") if s.shapes.title is not None and s.shapes.title.has_text_frame else ""
        lines.append(f"=== SLIDE {i} === {title}")
        for tag, tf in frames(s):
            paras = [para_md(p) for p in tf.paragraphs]
            lines.append(f"[{tag}] {paras[0] if paras else ''}")
            lines += [f"      {p}" for p in paras[1:]]
        lines.append("")
    open(out, "w", encoding="utf-8").write("\n".join(lines))
    print("exported", len(prs.slides), "slides ->", out)


def parse(md):
    blocks, notes, slide, tag = {}, [], None, None
    for raw in open(md, encoding="utf-8").read().splitlines():
        if raw.startswith("#") and slide is None:
            continue
        m = SLIDE.match(raw)
        if m:
            slide, tag = int(m.group(1)), None
            continue
        if raw.startswith(">> NOTE"):
            notes.append((slide, raw))
            continue
        m = TAG.match(raw)
        if m and slide is not None:
            tag = m.group(1)
            blocks[(slide, tag)] = [m.group(2)]
        elif tag is not None and raw.startswith("      "):
            blocks[(slide, tag)].append(raw[6:])
        elif tag is not None and raw.strip():
            blocks[(slide, tag)].append(raw.strip())
    for k, v in blocks.items():
        while v and not v[-1].strip():
            v.pop()
    return blocks, notes


def set_para(p, md, proto_bold, proto_plain):
    for r in list(p.runs):
        r._r.getparent().remove(r._r)
    for i, part in enumerate(re.split(r"\*\*", md)):
        if not part:
            continue
        bold = i % 2 == 1
        # fall back to the other weight's formatting (bold is then set explicitly below)
        proto = next((x for x in ((proto_bold if bold else proto_plain), proto_plain, proto_bold) if x is not None), None)
        r_el = copy.deepcopy(proto) if proto is not None else None
        if r_el is not None:
            # runs must precede <a:endParaRPr>; PowerPoint ignores runs placed after it
            end = p._p.find("{http://schemas.openxmlformats.org/drawingml/2006/main}endParaRPr")
            if end is not None:
                end.addprevious(r_el)
            else:
                p._p.append(r_el)
            run = p.runs[-1]
            run.text = part
        else:
            run = p.add_run()
            run.text = part
        if run.font.bold != bold and (bold or run.font.bold):
            run.font.bold = bold


def apply(src, md, out):
    prs = Presentation(src)
    blocks, notes = parse(md)
    changed = []
    for i, s in enumerate(prs.slides, 1):
        for tag, tf in list(frames(s)):
            new = blocks.get((i, tag))
            if new is None:
                continue
            old = [para_md(p) for p in tf.paragraphs]
            if norm(new) == norm(old):
                continue
            runs = [r for p in tf.paragraphs for r in p.runs]
            pb = next((copy.deepcopy(r._r) for r in runs if r.font.bold), None)
            pp = next((copy.deepcopy(r._r) for r in runs if not r.font.bold), None)
            paras = list(tf.paragraphs)
            proto_p = copy.deepcopy(paras[-1]._p)
            for k, line in enumerate(new):
                if k < len(paras):
                    set_para(paras[k], line, pb, pp)
                else:
                    el = copy.deepcopy(proto_p)
                    tf._txBody.append(el)
                    set_para(tf.paragraphs[-1], line, pb, pp)
            for p in paras[len(new):]:
                p._p.getparent().remove(p._p)
            if not new:
                set_para(tf.paragraphs[0] if tf.paragraphs else tf.add_paragraph(), "", pb, pp)
            changed.append((i, tag))
    prs.save(out)
    print("changed blocks:", changed)
    print("notes for manual changes:", notes)


if __name__ == "__main__":
    cmd = sys.argv[1]
    export(*sys.argv[2:4]) if cmd == "export" else apply(*sys.argv[2:5])
