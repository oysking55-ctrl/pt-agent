# PT Agent

주제와 발표 항목(목차)만 입력하면 리서치 → PPT(.pptx) → 웹 슬라이드(HTML)까지 자동으로 만들어주는 발표자료 생성기입니다.

- 웹으로 보기: GitHub Pages (Reveal.js 슬라이드)
- 다운로드: 각 슬라이드 페이지 상단의 "PPT 다운로드" 버튼 (.pptx)

## 새 발표자료 만드는 방법

1. `generator/content/` 폴더에 새 JSON 파일을 만듭니다. (`01-visual-studio-agent.json` 참고)
   - `slug`: 폴더/URL에 쓰일 영문 식별자 (예: `02-my-topic`)
   - `topic`, `subtitle`: 표지 제목/부제
   - `sections[]`: 각 항목의 `title`과 `bullets`(발표 내용). 필요하면 `code`(예시 코드/설정)나 `notes`(발표자 노트, 출처)도 추가
   - `closing`: 마무리 슬라이드
2. 아래 명령으로 빌드합니다.

   ```bash
   python generator/build_deck.py generator/content/02-my-topic.json
   ```

3. `docs/<slug>/slides.pptx`, `docs/<slug>/index.html`이 생성되고, `docs/index.html`(전체 목록)도 자동 갱신됩니다.

내용 리서치는 Claude Code(웹 검색)로 각 항목의 최신 정보를 찾아 JSON에 채워 넣는 방식으로 진행합니다. 즉 "주제 + 항목 목록"을 Claude Code에게 주면, Claude Code가 리서치 후 JSON을 작성하고 위 스크립트로 빌드까지 한 번에 해줍니다.

## GitHub Pages로 배포하기

이 저장소는 `docs/` 폴더를 GitHub Pages 소스로 사용합니다.

1. GitHub에서 새 저장소 생성: `pt-agent` (Public)
2. 아래 명령으로 원격 저장소 연결 및 푸시

   ```bash
   git remote add origin https://github.com/<GitHub 계정명>/pt-agent.git
   git branch -M main
   git push -u origin main
   ```

3. GitHub 저장소 → Settings → Pages → Build and deployment
   - Source: `Deploy from a branch`
   - Branch: `main` / `docs`
   - Save

4. 몇 분 후 `https://<GitHub 계정명>.github.io/pt-agent/` 에서 발표자료 목록을 볼 수 있습니다.

## 로컬에서 미리보기

```bash
cd docs
python -m http.server 8000
```

브라우저에서 `http://localhost:8000` 접속.
