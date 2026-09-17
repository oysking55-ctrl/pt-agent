#!/usr/bin/env python3
"""Build a PPTX deck and a matching Reveal.js HTML deck from a content JSON file.

Usage:
    python build_deck.py content/01-visual-studio-agent.json

Outputs (relative to repo root):
    docs/<slug>/slides.pptx
    docs/<slug>/index.html
Also updates docs/index.html with a card linking to the new deck.
"""
import json
import re
import sys
import html
from pathlib import Path

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

REPO_ROOT = Path(__file__).resolve().parent.parent
DOCS_DIR = REPO_ROOT / "docs"

NAVY = RGBColor(0x11, 0x1C, 0x3B)
ACCENT = RGBColor(0x4F, 0x8C, 0xFF)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
SLATE = RGBColor(0x33, 0x3A, 0x4D)


def load_content(path: Path) -> dict:
    with open(path, encoding="utf-8") as f:
        return json.load(f)


# ---------------------------------------------------------------------------
# PPTX build
# ---------------------------------------------------------------------------

def set_background(slide, color: RGBColor):
    bg = slide.background
    bg.fill.solid()
    bg.fill.fore_color.rgb = color


def add_title_slide(prs: Presentation, topic: str, subtitle: str):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_background(slide, NAVY)

    box = slide.shapes.add_textbox(Inches(0.8), Inches(2.3), Inches(11.7), Inches(2.2))
    tf = box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = topic
    p.font.size = Pt(40)
    p.font.bold = True
    p.font.color.rgb = WHITE

    box2 = slide.shapes.add_textbox(Inches(0.8), Inches(4.3), Inches(11.7), Inches(1))
    tf2 = box2.text_frame
    p2 = tf2.paragraphs[0]
    p2.text = subtitle
    p2.font.size = Pt(22)
    p2.font.color.rgb = ACCENT

    rule = slide.shapes.add_shape(1, Inches(0.85), Inches(2.15), Inches(1.4), Pt(4))
    rule.fill.solid()
    rule.fill.fore_color.rgb = ACCENT
    rule.line.fill.background()
    return slide


def add_content_slide(prs: Presentation, index: int, total: int, title: str, bullets, code: str = None):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_background(slide, WHITE)

    # top accent bar
    bar = slide.shapes.add_shape(1, Inches(0), Inches(0), Inches(13.333), Inches(0.12))
    bar.fill.solid()
    bar.fill.fore_color.rgb = ACCENT
    bar.line.fill.background()

    title_box = slide.shapes.add_textbox(Inches(0.7), Inches(0.45), Inches(11.9), Inches(1.0))
    tf = title_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = title
    p.font.size = Pt(30)
    p.font.bold = True
    p.font.color.rgb = NAVY

    body_top = Inches(1.55)
    body_height = Inches(5.2)
    if code:
        body_width = Inches(7.0)
    else:
        body_width = Inches(11.9)

    body_box = slide.shapes.add_textbox(Inches(0.7), body_top, body_width, body_height)
    tf = body_box.text_frame
    tf.word_wrap = True
    for i, bullet in enumerate(bullets):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = f"•  {bullet}"
        p.font.size = Pt(17)
        p.font.color.rgb = SLATE
        p.space_after = Pt(12)

    if code:
        code_box = slide.shapes.add_textbox(Inches(7.9), body_top, Inches(4.7), Inches(3.2))
        code_box.fill.solid()
        code_box.fill.fore_color.rgb = NAVY
        tf = code_box.text_frame
        tf.word_wrap = True
        tf.margin_left = Pt(12)
        tf.margin_top = Pt(12)
        for i, line in enumerate(code.split("\n")):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.text = line
            p.font.size = Pt(13)
            p.font.name = "Consolas"
            p.font.color.rgb = RGBColor(0x9C, 0xD8, 0xFF)

    # page number
    pn = slide.shapes.add_textbox(Inches(12.5), Inches(7.05), Inches(0.6), Inches(0.35))
    p = pn.text_frame.paragraphs[0]
    p.text = f"{index}/{total}"
    p.font.size = Pt(11)
    p.font.color.rgb = SLATE
    return slide


def build_pptx(data: dict, out_path: Path):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    add_title_slide(prs, data["topic"], data.get("subtitle", ""))

    sections = data["sections"]
    total = len(sections) + 2  # title + closing
    for i, sec in enumerate(sections, start=2):
        add_content_slide(prs, i, total, sec["title"], sec["bullets"], sec.get("code"))

    closing = data.get("closing")
    if closing:
        add_content_slide(prs, total, total, closing["title"], closing["bullets"])

    out_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(out_path)


# ---------------------------------------------------------------------------
# Reveal.js HTML build
# ---------------------------------------------------------------------------

REVEAL_CDN = "https://cdnjs.cloudflare.com/ajax/libs/reveal.js/5.1.0"

HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>{topic}</title>
<link rel="stylesheet" href="{cdn}/reveal.css">
<link rel="stylesheet" href="{cdn}/theme/night.css" id="theme">
<link rel="stylesheet" href="{cdn}/plugin/highlight/monokai.css">
<style>
  :root {{ --accent: #4f8cff; }}
  .reveal h1, .reveal h2 {{ color: var(--accent); }}
  .reveal .slides section {{ text-align: left; }}
  .reveal ul {{ display: block; }}
  .reveal .title-slide {{ text-align: center; }}
  .reveal .title-slide h1 {{ font-size: 2.2em; }}
  .reveal .title-slide h3 {{ color: #ccd6f6; font-weight: 400; }}
  .topbar {{
    position: fixed; top: 0; left: 0; right: 0; padding: 10px 20px;
    display: flex; justify-content: flex-end; gap: 10px; z-index: 100;
  }}
  .topbar a {{
    background: var(--accent); color: #fff; text-decoration: none;
    padding: 8px 16px; border-radius: 6px; font-size: 14px; font-family: sans-serif;
  }}
  .topbar a:hover {{ opacity: 0.85; }}
</style>
</head>
<body>
<div class="topbar">
  <a href="slides.pptx" download>PPT 다운로드</a>
  <a href="../">← 목록으로</a>
</div>
<div class="reveal">
  <div class="slides">
{slides}
  </div>
</div>
<script src="{cdn}/reveal.js"></script>
<script src="{cdn}/plugin/notes/notes.js"></script>
<script src="{cdn}/plugin/highlight/highlight.js"></script>
<script>
  Reveal.initialize({{
    hash: true,
    slideNumber: true,
    transition: 'slide',
    plugins: [ RevealNotes, RevealHighlight ]
  }});
</script>
</body>
</html>
"""


def esc(s: str) -> str:
    return html.escape(s, quote=False)


def build_html(data: dict, out_path: Path):
    slides_html = []

    slides_html.append(f"""    <section class="title-slide">
      <h1>{esc(data['topic'])}</h1>
      <h3>{esc(data.get('subtitle', ''))}</h3>
    </section>""")

    for sec in data["sections"]:
        notes = sec.get("notes", "")
        notes_html = f"\n      <aside class=\"notes\">{esc(notes)}</aside>" if notes else ""
        bullets_html = "\n".join(f"        <li>{esc(b)}</li>" for b in sec["bullets"])
        code = sec.get("code")
        if code:
            body = f"""      <div style="display:flex; gap:2em; align-items:flex-start;">
        <ul style="flex:1.3">
{bullets_html}
        </ul>
        <pre style="flex:1"><code class="language-json">{esc(code)}</code></pre>
      </div>"""
        else:
            body = f"""      <ul>
{bullets_html}
      </ul>"""
        slides_html.append(f"""    <section>
      <h2>{esc(sec['title'])}</h2>
{body}{notes_html}
    </section>""")

    closing = data.get("closing")
    if closing:
        bullets_html = "\n".join(f"        <li>{esc(b)}</li>" for b in closing["bullets"])
        slides_html.append(f"""    <section>
      <h2>{esc(closing['title'])}</h2>
      <ul>
{bullets_html}
      </ul>
    </section>""")

    html_out = HTML_TEMPLATE.format(
        topic=esc(data["topic"]),
        cdn=REVEAL_CDN,
        slides="\n".join(slides_html),
    )
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text(html_out, encoding="utf-8")


# ---------------------------------------------------------------------------
# Gallery index update
# ---------------------------------------------------------------------------

GALLERY_TEMPLATE = """<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>PT Agent 발표자료 모음</title>
<style>
  :root {{ --accent:#4f8cff; --bg:#0b1020; --card:#141b31; --text:#e8ecf7; --muted:#9aa4c2; }}
  * {{ box-sizing: border-box; }}
  body {{ margin:0; background:var(--bg); color:var(--text); font-family:'Segoe UI',sans-serif; padding:40px 20px; }}
  h1 {{ text-align:center; margin-bottom:8px; }}
  p.sub {{ text-align:center; color:var(--muted); margin-top:0; }}
  .grid {{ display:grid; grid-template-columns:repeat(auto-fill,minmax(280px,1fr)); gap:20px; max-width:1000px; margin:40px auto 0; }}
  .card {{ background:var(--card); border-radius:12px; padding:24px; border:1px solid #262f4d; }}
  .card h2 {{ margin:0 0 8px; font-size:1.15em; color:var(--accent); }}
  .card p {{ color:var(--muted); font-size:0.9em; margin:0 0 16px; }}
  .card a {{ display:inline-block; margin-right:10px; padding:8px 14px; border-radius:6px; text-decoration:none; font-size:0.85em; }}
  .card a.view {{ background:var(--accent); color:#fff; }}
  .card a.dl {{ border:1px solid var(--accent); color:var(--accent); }}
</style>
</head>
<body>
<h1>PT Agent 발표자료 모음</h1>
<p class="sub">주제와 목차만 입력하면 리서치 → PPT → 웹 슬라이드까지 자동 생성됩니다</p>
<div class="grid">
{cards}
</div>
</body>
</html>
"""

CARD_TEMPLATE = """  <div class="card">
    <h2>{topic}</h2>
    <p>{subtitle}</p>
    <a class="view" href="{slug}/">웹으로 보기</a>
    <a class="dl" href="{slug}/slides.pptx" download>PPT 다운로드</a>
  </div>"""


def update_gallery(entries):
    cards = "\n".join(
        CARD_TEMPLATE.format(topic=esc(e["topic"]), subtitle=esc(e.get("subtitle", "")), slug=e["slug"])
        for e in entries
    )
    out = GALLERY_TEMPLATE.format(cards=cards)
    (DOCS_DIR / "index.html").write_text(out, encoding="utf-8")


def discover_existing_entries():
    entries = []
    content_dir = Path(__file__).resolve().parent / "content"
    for p in sorted(content_dir.glob("*.json")):
        d = load_content(p)
        entries.append({"slug": d["slug"], "topic": d["topic"], "subtitle": d.get("subtitle", "")})
    return entries


def main():
    if len(sys.argv) != 2:
        print("Usage: python build_deck.py <content.json>")
        sys.exit(1)

    content_path = Path(sys.argv[1])
    data = load_content(content_path)
    slug = data["slug"]

    out_dir = DOCS_DIR / slug
    build_pptx(data, out_dir / "slides.pptx")
    build_html(data, out_dir / "index.html")

    update_gallery(discover_existing_entries())

    print(f"Built: {out_dir / 'slides.pptx'}")
    print(f"Built: {out_dir / 'index.html'}")
    print(f"Updated: {DOCS_DIR / 'index.html'}")


if __name__ == "__main__":
    main()
