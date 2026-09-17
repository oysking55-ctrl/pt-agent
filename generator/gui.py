"""PT Agent desktop launcher.

A small Tkinter app with two tabs:

1. "새 발표자료 만들기" — collects a topic, audience, an optional
   scenario/concept, and an optional outline (one item per line). If
   no outline is given, Claude Code is asked to propose one from the
   concept. Everything is handed off to Claude Code: it writes a
   content skeleton JSON plus a request note, and opens a new terminal
   running `claude` in the project folder so Claude Code can research,
   suggest/generate the outline, fill in the content, and build the
   PPTX + web deck.

2. "기존 발표자료 수정/보완" — lists decks already built under
   generator/content/, lets you open the web preview for one, and
   send a follow-up revision request to Claude Code.

This file is frozen into a standalone .exe with PyInstaller.
"""
import json
import re
import subprocess
import sys
import tkinter as tk
import webbrowser
from pathlib import Path
from tkinter import messagebox, scrolledtext, ttk

CONFIG_NAME = "pt_agent_config.json"
DEFAULT_REPO_ROOT = r"C:\Users\SKTelecom\4. PT 자료 만들기"


def app_dir() -> Path:
    """Directory the exe (or this script) lives in — used to find/save config."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def load_repo_root() -> Path:
    cfg_path = app_dir() / CONFIG_NAME
    if cfg_path.exists():
        try:
            cfg = json.loads(cfg_path.read_text(encoding="utf-8"))
            root = Path(cfg.get("repo_root", DEFAULT_REPO_ROOT))
            if root.exists():
                return root
        except Exception:
            pass
    return Path(DEFAULT_REPO_ROOT)


def save_repo_root(root: Path):
    cfg_path = app_dir() / CONFIG_NAME
    cfg_path.write_text(json.dumps({"repo_root": str(root)}, ensure_ascii=False, indent=2), encoding="utf-8")


def slugify(topic: str, existing: set) -> str:
    ascii_part = re.sub(r"[^a-zA-Z0-9]+", "-", topic).strip("-").lower()
    n = 1
    while f"{n:02d}" in {e[:2] for e in existing}:
        n += 1
    base = ascii_part if ascii_part else "deck"
    slug = f"{n:02d}-{base}"[:50].rstrip("-")
    while slug in existing:
        n += 1
        slug = f"{n:02d}-{base}"[:50].rstrip("-")
    return slug


class PTAgentApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.repo_root = load_repo_root()
        root.title("PT Agent — 발표자료 생성")
        root.geometry("680x720")
        root.minsize(600, 600)

        notebook = ttk.Notebook(root)
        notebook.pack(fill="both", expand=True)

        self.new_tab = ttk.Frame(notebook)
        self.edit_tab = ttk.Frame(notebook)
        notebook.add(self.new_tab, text="새 발표자료 만들기")
        notebook.add(self.edit_tab, text="기존 발표자료 수정/보완")
        notebook.bind("<<NotebookTabChanged>>", lambda e: self.refresh_deck_list())

        self.build_new_tab(self.new_tab)
        self.build_edit_tab(self.edit_tab)

    # ------------------------------------------------------------------
    # Tab 1 — new deck
    # ------------------------------------------------------------------
    def build_new_tab(self, parent):
        pad = {"padx": 12, "pady": 6}

        ttk.Label(parent, text="발표 주제", font=("Segoe UI", 11, "bold")).pack(anchor="w", **pad)
        self.topic_entry = ttk.Entry(parent, font=("Segoe UI", 11))
        self.topic_entry.pack(fill="x", **pad)
        self.topic_entry.insert(0, "Visual Studio를 이용한 나만의 Agent 만들기")

        ttk.Label(parent, text="부제 (선택)", font=("Segoe UI", 10)).pack(anchor="w", **pad)
        self.subtitle_entry = ttk.Entry(parent, font=("Segoe UI", 10))
        self.subtitle_entry.pack(fill="x", **pad)
        self.subtitle_entry.insert(0, "w/ Claude Code")

        ttk.Label(
            parent,
            text="발표 대상 (누구에게 발표하나요? 수준에 맞춰 내용을 조절합니다)",
            font=("Segoe UI", 11, "bold"),
        ).pack(anchor="w", **pad)
        self.audience_combo = ttk.Combobox(
            parent,
            font=("Segoe UI", 10),
            values=["초등학생", "중학생", "고등학생", "대학생/일반인", "실무자/전문가"],
        )
        self.audience_combo.pack(fill="x", **pad)
        self.audience_combo.set("중학생")

        ttk.Label(
            parent,
            text="발표 시나리오/컨셉 (선택 — 이 내용을 보고 Claude Code가 발표 항목을 추천합니다)",
            font=("Segoe UI", 11, "bold"),
        ).pack(anchor="w", **pad)
        self.concept_box = scrolledtext.ScrolledText(parent, font=("Segoe UI", 10), height=4, wrap="word")
        self.concept_box.pack(fill="x", **pad)

        ttk.Label(
            parent,
            text="발표 항목 (한 줄에 하나씩, 선택 — 비워두면 위 컨셉을 바탕으로 추천받습니다)",
            font=("Segoe UI", 11, "bold"),
        ).pack(anchor="w", **pad)
        self.items_box = scrolledtext.ScrolledText(parent, font=("Segoe UI", 10), height=10, wrap="word")
        self.items_box.pack(fill="both", expand=True, **pad)

        opts = ttk.Frame(parent)
        opts.pack(fill="x", **pad)
        self.suggest_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            opts,
            text="항목을 입력했어도 Claude Code에게 추가/변형을 추천받기",
            variable=self.suggest_var,
        ).pack(anchor="w")

        root_row = ttk.Frame(parent)
        root_row.pack(fill="x", **pad)
        ttk.Label(root_row, text="프로젝트 폴더:").pack(side="left")
        self.repo_label = ttk.Label(root_row, text=str(self.repo_root), foreground="#555")
        self.repo_label.pack(side="left", padx=(6, 0))
        ttk.Button(root_row, text="변경", command=self.change_repo_root).pack(side="right")

        btn_row = ttk.Frame(parent)
        btn_row.pack(fill="x", **pad)
        ttk.Button(btn_row, text="Claude Code로 만들기 시작", command=self.submit).pack(
            side="right", ipadx=10, ipady=6
        )

    def change_repo_root(self):
        from tkinter import filedialog

        chosen = filedialog.askdirectory(title="PT Agent 프로젝트 폴더 선택", initialdir=str(self.repo_root))
        if chosen:
            self.repo_root = Path(chosen)
            self.repo_label.config(text=str(self.repo_root))
            save_repo_root(self.repo_root)
            self.refresh_deck_list()

    def submit(self):
        topic = self.topic_entry.get().strip()
        subtitle = self.subtitle_entry.get().strip()
        audience = self.audience_combo.get().strip()
        concept = self.concept_box.get("1.0", "end").strip()
        items = [line.strip() for line in self.items_box.get("1.0", "end").splitlines() if line.strip()]

        if not topic:
            messagebox.showwarning("입력 필요", "발표 주제를 입력해주세요.")
            return
        if not audience:
            messagebox.showwarning("입력 필요", "발표 대상을 입력해주세요. (예: 중학생, 실무자 등)")
            return
        if not items and not concept:
            messagebox.showwarning(
                "입력 필요",
                "발표 항목을 입력하거나, 발표 시나리오/컨셉을 입력해주세요.\n"
                "컨셉만 입력하면 Claude Code가 항목을 추천해줍니다.",
            )
            return

        content_dir = self.repo_root / "generator" / "content"
        if not content_dir.exists():
            messagebox.showerror(
                "폴더를 찾을 수 없음",
                f"{content_dir} 를 찾을 수 없습니다. '변경' 버튼으로 PT Agent 프로젝트 폴더를 다시 지정해주세요.",
            )
            return

        existing = {p.stem for p in content_dir.glob("*.json")}
        slug = slugify(topic, existing)

        skeleton = {
            "slug": slug,
            "topic": topic,
            "subtitle": subtitle,
            "audience": audience,
            "sections": [{"title": item, "bullets": [], "terms": [], "notes": ""} for item in items],
            "closing": {"title": "정리", "bullets": [], "terms": []},
        }
        content_path = content_dir / f"{slug}.json"
        content_path.write_text(json.dumps(skeleton, ensure_ascii=False, indent=2), encoding="utf-8")

        request_path = content_dir / f"{slug}.request.txt"
        request_path.write_text(
            self.build_request_text(slug, topic, subtitle, audience, concept, items), encoding="utf-8"
        )

        self.launch_claude(request_path)

        messagebox.showinfo(
            "요청 전달 완료",
            f"'{content_path.name}' 초안을 저장하고 Claude Code를 새 터미널에서 열었습니다.\n"
            "터미널 창에서 대화를 이어가면 항목 추천, 리서치, PPT/웹 슬라이드 생성까지 진행됩니다.",
        )
        self.refresh_deck_list()

    def build_request_text(self, slug, topic, subtitle, audience, concept, items):
        concept_block = f"\n발표 시나리오/컨셉:\n{concept}\n" if concept else ""

        if items:
            item_lines = "\n".join(f"{i+1}. {it}" for i, it in enumerate(items))
            if self.suggest_var.get():
                outline_instruction = (
                    f"입력된 발표 항목:\n{item_lines}\n\n"
                    "먼저 이 항목 구성에 빠진 게 없는지, 순서나 표현을 다듬을 부분이 없는지"
                    + (" 위 컨셉을 참고해서" if concept else "")
                    + " 검토해서 추가하거나 바꾸면 좋을 항목을 제안해줘. "
                    "내가 확인/수정한 뒤 다음 단계로 진행해줘.\n"
                )
            else:
                outline_instruction = f"입력된 발표 항목(그대로 사용):\n{item_lines}\n"
        else:
            outline_instruction = (
                "발표 항목이 아직 없어. 위 시나리오/컨셉을 바탕으로 이 발표에 어울리는 "
                "발표 항목(목차)을 새로 제안해줘. 제안한 목차를 먼저 나에게 보여주고 "
                "확인/수정을 받은 다음 다음 단계로 진행해줘.\n"
            )

        return f"""PT Agent 요청 — generator/content/{slug}.json

주제: {topic}
부제: {subtitle}
발표 대상: {audience}
{concept_block}
{outline_instruction}
그다음 (항목이 확정되면) 각 항목을 웹 검색으로 리서치해서 generator/content/{slug}.json 의
sections[].bullets (항목별 핵심 내용 4~7개)와 closing.bullets 를 채워줘.
항목이 새로 추가/변경됐다면 sections 배열 자체를 새 목차에 맞게 다시 써줘.
필요하면 sections[].code 필드에 짧은 예시 코드/설정을 추가해도 좋아.

**도식화 (중요):**
순서/단계, 흐름, 비교처럼 글보다 그림으로 보여주면 더 이해하기 쉬운 내용은
sections[].diagram 필드에 ["1단계", "2단계", "3단계", ...] 형태로 넣어줘.
화면에는 화살표로 이어지는 박스 흐름도로 자동으로 그려져. 각 항목은 3~6단어로 짧게.
모든 슬라이드에 넣을 필요는 없고, 프로세스/절차/비교를 설명하는 슬라이드에만 자연스럽게 사용해줘
(하나의 발표에 1~3개 정도가 적당해). 도식화가 어울리지 않는 슬라이드는 diagram 없이 bullets만 써도 돼.

**아주 중요 — 대상 맞춤:**
모든 문장의 단어 선택과 설명 수준을 "{audience}"에 맞춰줘.
예를 들어 대상이 초/중/고 학생이면 전문 용어를 최소화하고 쉬운 말과 비유로 풀어서 설명하고,
실무자/전문가면 전문 용어를 그대로 쓰고 더 깊이 있는 내용을 다뤄줘.

**용어 설명 (필수):**
전문 용어나 어려운 단어를 한 개라도 쓴 슬라이드에는 sections[].terms 배열에
{{"term": "용어", "def": "쉬운 설명"}} 형태로 추가해줘. terms 안의 설명도 "{audience}" 눈높이로 써줘.
용어가 없는 슬라이드는 terms를 빈 배열로 둬도 돼.

내용을 다 채운 뒤에는 아래 명령으로 PPT와 웹 슬라이드를 빌드해줘:
    python generator/build_deck.py generator/content/{slug}.json

빌드 후 docs/{slug}/index.html 을 열어 슬라이드가 한 화면에 잘 들어가는지, 용어 설명이 각 슬라이드 아래쪽에 잘 보이는지 확인해줘.
"""

    # ------------------------------------------------------------------
    # Tab 2 — review / revise an existing deck
    # ------------------------------------------------------------------
    def build_edit_tab(self, parent):
        pad = {"padx": 12, "pady": 6}

        ttk.Label(parent, text="발표자료 선택", font=("Segoe UI", 11, "bold")).pack(anchor="w", **pad)

        pick_row = ttk.Frame(parent)
        pick_row.pack(fill="x", **pad)
        self.deck_combo = ttk.Combobox(pick_row, font=("Segoe UI", 10), state="readonly")
        self.deck_combo.pack(side="left", fill="x", expand=True)
        ttk.Button(pick_row, text="새로고침", command=self.refresh_deck_list).pack(side="left", padx=(8, 0))

        preview_row = ttk.Frame(parent)
        preview_row.pack(fill="x", **pad)
        ttk.Button(preview_row, text="웹으로 미리보기", command=self.preview_deck).pack(side="left")
        ttk.Button(preview_row, text="PPT 파일 열기", command=self.open_pptx).pack(side="left", padx=(8, 0))

        ttk.Label(
            parent,
            text="수정/보완 요청 내용 (예: '3번째 슬라이드에 예시를 하나 더 추가해줘', '전체적으로 더 재미있게 해줘')",
            font=("Segoe UI", 11, "bold"),
        ).pack(anchor="w", **pad)
        self.revise_box = scrolledtext.ScrolledText(parent, font=("Segoe UI", 10), height=12, wrap="word")
        self.revise_box.pack(fill="both", expand=True, **pad)

        btn_row = ttk.Frame(parent)
        btn_row.pack(fill="x", **pad)
        ttk.Button(btn_row, text="Claude Code에게 수정 요청", command=self.submit_revision).pack(
            side="right", ipadx=10, ipady=6
        )

        self.refresh_deck_list()

    def content_dir(self) -> Path:
        return self.repo_root / "generator" / "content"

    def list_decks(self):
        decks = []
        cdir = self.content_dir()
        if not cdir.exists():
            return decks
        for p in sorted(cdir.glob("*.json")):
            try:
                data = json.loads(p.read_text(encoding="utf-8"))
                decks.append((data.get("slug", p.stem), data.get("topic", p.stem)))
            except Exception:
                continue
        return decks

    def refresh_deck_list(self):
        if not hasattr(self, "deck_combo"):
            return
        self._decks = self.list_decks()
        display = [f"{topic}  ({slug})" for slug, topic in self._decks]
        current = self.deck_combo.get()
        self.deck_combo["values"] = display
        if display and current not in display:
            self.deck_combo.current(len(display) - 1)

    def selected_slug(self):
        idx = self.deck_combo.current()
        if idx < 0 or idx >= len(getattr(self, "_decks", [])):
            messagebox.showwarning("선택 필요", "수정할 발표자료를 먼저 선택해주세요.")
            return None
        return self._decks[idx][0]

    def preview_deck(self):
        slug = self.selected_slug()
        if not slug:
            return
        html_path = self.repo_root / "docs" / slug / "index.html"
        if not html_path.exists():
            messagebox.showwarning(
                "아직 만들어지지 않음", f"{html_path} 를 찾을 수 없습니다. 먼저 발표자료를 빌드해주세요."
            )
            return
        webbrowser.open(html_path.resolve().as_uri())

    def open_pptx(self):
        slug = self.selected_slug()
        if not slug:
            return
        pptx_path = self.repo_root / "docs" / slug / "slides.pptx"
        if not pptx_path.exists():
            messagebox.showwarning(
                "아직 만들어지지 않음", f"{pptx_path} 를 찾을 수 없습니다. 먼저 발표자료를 빌드해주세요."
            )
            return
        try:
            import os

            os.startfile(str(pptx_path))
        except Exception as e:
            messagebox.showerror("열기 실패", str(e))

    def submit_revision(self):
        slug = self.selected_slug()
        if not slug:
            return
        changes = self.revise_box.get("1.0", "end").strip()
        if not changes:
            messagebox.showwarning("입력 필요", "수정/보완 요청 내용을 입력해주세요.")
            return

        topic = next((t for s, t in self._decks if s == slug), slug)
        request_path = self.content_dir() / f"{slug}.revise.request.txt"
        request_path.write_text(
            f"""PT Agent 수정 요청 — generator/content/{slug}.json

이미 만들어진 발표자료 '{topic}' (generator/content/{slug}.json, docs/{slug}/)를 아래 요청대로 수정해줘.
필요하면 generator/content/{slug}.json 의 sections/terms/closing 을 직접 고치고,
슬라이드 수나 순서를 바꿔야 하면 그렇게 해도 좋아.
순서/절차/비교처럼 글보다 그림이 이해하기 쉬운 내용은 sections[].diagram 필드에
["1단계", "2단계", ...] 형태로 넣으면 화살표로 이어지는 흐름도가 자동으로 그려져.

요청 내용:
{changes}

수정 후에는 아래 명령으로 다시 빌드해줘:
    python generator/build_deck.py generator/content/{slug}.json

빌드 후 docs/{slug}/index.html 을 열어 요청한 대로 잘 반영됐는지, 슬라이드가 한 화면에 잘 들어가는지 확인해줘.
""",
            encoding="utf-8",
        )

        self.launch_claude(request_path)

        messagebox.showinfo(
            "수정 요청 전달 완료",
            "수정 요청을 저장하고 Claude Code를 새 터미널에서 열었습니다.\n"
            "터미널 창에서 대화를 이어가면 수정 사항이 반영됩니다.",
        )

    # ------------------------------------------------------------------
    # Shared: launch Claude Code with a request file
    # ------------------------------------------------------------------
    def launch_claude(self, request_path: Path):
        rel = request_path.relative_to(self.repo_root)
        prompt = f"{rel.as_posix()} 파일을 읽고, 그 안의 요청대로 진행해줘."
        # Pass the prompt as a real argv entry (not through a batch file) so
        # Windows delivers the Korean text via CreateProcessW untouched —
        # writing it into a .bat and letting cmd.exe re-parse the file from
        # disk corrupts multi-byte characters before `chcp 65001` can help.
        subprocess.Popen(
            ["cmd", "/k", "claude", prompt],
            cwd=str(self.repo_root),
            creationflags=subprocess.CREATE_NEW_CONSOLE,
        )


def main():
    root = tk.Tk()
    try:
        style = ttk.Style()
        if "vista" in style.theme_names():
            style.theme_use("vista")
    except Exception:
        pass
    PTAgentApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
