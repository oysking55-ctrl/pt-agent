"""PT Agent desktop launcher.

A small Tkinter form that collects a presentation topic and an outline
(one item per line), lets the user review/edit that outline, then hands
everything off to Claude Code: it writes a content skeleton JSON plus a
request note, and opens a new terminal running `claude` in the project
folder so Claude Code can research, suggest additions/tweaks to the
outline, fill in the content, and build the PPTX + web deck.

This file is frozen into a standalone .exe with PyInstaller; see
generator/build_exe.py.
"""
import json
import re
import subprocess
import sys
import tempfile
import tkinter as tk
from datetime import datetime
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
        root.geometry("640x620")
        root.minsize(560, 520)

        pad = {"padx": 12, "pady": 6}

        ttk.Label(root, text="발표 주제", font=("Segoe UI", 11, "bold")).pack(anchor="w", **pad)
        self.topic_entry = ttk.Entry(root, font=("Segoe UI", 11))
        self.topic_entry.pack(fill="x", **pad)
        self.topic_entry.insert(0, "Visual Studio를 이용한 나만의 Agent 만들기")

        ttk.Label(root, text="부제 (선택)", font=("Segoe UI", 10)).pack(anchor="w", **pad)
        self.subtitle_entry = ttk.Entry(root, font=("Segoe UI", 10))
        self.subtitle_entry.pack(fill="x", **pad)
        self.subtitle_entry.insert(0, "w/ Claude Code")

        ttk.Label(
            root,
            text="발표 항목 (한 줄에 하나씩, 순서대로 입력)",
            font=("Segoe UI", 11, "bold"),
        ).pack(anchor="w", **pad)
        self.items_box = scrolledtext.ScrolledText(root, font=("Segoe UI", 10), height=14, wrap="word")
        self.items_box.pack(fill="both", expand=True, **pad)

        opts = ttk.Frame(root)
        opts.pack(fill="x", **pad)
        self.suggest_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            opts,
            text="Claude Code에게 항목 추가/변형 추천을 먼저 요청하기",
            variable=self.suggest_var,
        ).pack(anchor="w")

        root_row = ttk.Frame(root)
        root_row.pack(fill="x", **pad)
        ttk.Label(root_row, text="프로젝트 폴더:").pack(side="left")
        self.repo_label = ttk.Label(root_row, text=str(self.repo_root), foreground="#555")
        self.repo_label.pack(side="left", padx=(6, 0))
        ttk.Button(root_row, text="변경", command=self.change_repo_root).pack(side="right")

        btn_row = ttk.Frame(root)
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

    def submit(self):
        topic = self.topic_entry.get().strip()
        subtitle = self.subtitle_entry.get().strip()
        items = [line.strip() for line in self.items_box.get("1.0", "end").splitlines() if line.strip()]

        if not topic:
            messagebox.showwarning("입력 필요", "발표 주제를 입력해주세요.")
            return
        if not items:
            messagebox.showwarning("입력 필요", "발표 항목을 한 줄에 하나씩 입력해주세요.")
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
            "sections": [{"title": item, "bullets": [], "notes": ""} for item in items],
            "closing": {"title": "정리", "bullets": []},
        }
        content_path = content_dir / f"{slug}.json"
        content_path.write_text(json.dumps(skeleton, ensure_ascii=False, indent=2), encoding="utf-8")

        request_path = content_dir / f"{slug}.request.txt"
        request_path.write_text(self.build_request_text(slug, topic, subtitle, items), encoding="utf-8")

        self.launch_claude(slug)

        messagebox.showinfo(
            "요청 전달 완료",
            f"'{content_path.name}' 초안을 저장하고 Claude Code를 새 터미널에서 열었습니다.\n"
            "터미널 창에서 대화를 이어가면 항목 추천, 리서치, PPT/웹 슬라이드 생성까지 진행됩니다.",
        )

    def build_request_text(self, slug, topic, subtitle, items):
        item_lines = "\n".join(f"{i+1}. {it}" for i, it in enumerate(items))
        suggest_block = (
            "\n먼저 이 항목 구성에 빠진 게 없는지, 순서나 표현을 다듬을 부분이 없는지 검토해서 "
            "추가하거나 바꾸면 좋을 항목을 제안해줘. 내가 확인/수정한 뒤 다음 단계로 진행해줘.\n"
            if self.suggest_var.get()
            else "\n"
        )
        return f"""PT Agent 요청 — generator/content/{slug}.json

주제: {topic}
부제: {subtitle}

입력된 발표 항목:
{item_lines}
{suggest_block}
그다음 각 항목을 웹 검색으로 리서치해서 generator/content/{slug}.json 의
sections[].bullets (항목별 핵심 내용 4~7개)와 closing.bullets 를 채워줘.
필요하면 sections[].code 필드에 짧은 예시 코드/설정을 추가해도 좋아.

내용을 다 채운 뒤에는 아래 명령으로 PPT와 웹 슬라이드를 빌드해줘:
    python generator/build_deck.py generator/content/{slug}.json

빌드 후 docs/{slug}/index.html 을 열어 슬라이드가 한 화면에 잘 들어가는지 확인해줘.
"""

    def launch_claude(self, slug: str):
        prompt = f"generator/content/{slug}.request.txt 파일을 읽고, 그 안의 요청대로 진행해줘."
        bat_path = Path(tempfile.gettempdir()) / f"pt_agent_launch_{slug}_{datetime.now():%H%M%S}.bat"
        bat_path.write_text(
            "@echo off\r\n"
            "chcp 65001 > nul\r\n"
            f'cd /d "{self.repo_root}"\r\n'
            f'claude "{prompt}"\r\n',
            encoding="utf-8",
        )
        subprocess.Popen(
            ["cmd", "/c", "start", "PT Agent - Claude Code", str(bat_path)],
            cwd=str(self.repo_root),
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
