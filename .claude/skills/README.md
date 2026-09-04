# Installed Agent Skills

60 skills installed under `.claude/skills/`. They are discovered automatically by
Claude Code when a session starts in this repository.

## Contents

| Group | Skills |
| --- | --- |
| Documents & media | `docx`, `pptx`, `xlsx`, `pdf`, `canvas-design`, `algorithmic-art`, `slack-gif-creator`, `theme-factory`, `brand-guidelines` |
| Web & frontend | `web-artifacts-builder`, `web-design-guidelines`, `frontend-design`, `webapp-testing` |
| Firecrawl (web data) | `firecrawl`, `firecrawl-agent`, `firecrawl-crawl`, `firecrawl-download`, `firecrawl-interact`, `firecrawl-map`, `firecrawl-monitor`, `firecrawl-parse`, `firecrawl-scrape`, `firecrawl-search` |
| Engineering practice | `test-driven-development-tdd`, `systematic-debugging`, `root-cause-tracing`, `testing-anti-patterns`, `condition-based-waiting`, `defense-in-depth-validation`, `verification-before-completion`, `using-git-worktrees`, `finishing-a-development-branch` |
| Planning & review | `writing-plans`, `executing-plans`, `brainstorming-ideas-into-designs`, `requesting-code-review`, `code-review-reception`, `subagent-driven-development`, `dispatching-parallel-agents` |
| Thinking tools | `collision-zone-thinking`, `inversion-exercise`, `meta-pattern-recognition`, `preserving-productive-tensions`, `scale-game`, `simplification-cascades`, `tracing-knowledge-lineages`, `when-stuck-problem-solving-dispatch` |
| Skill authoring | `skill-creator`, `writing-skills`, `testing-skills-with-subagents`, `getting-started-with-skills`, `find-skills`, `sharing-skills`, `gardening-skills-wiki`, `pulling-updates-from-skills-repository` |
| Other | `claude-api`, `mcp-builder`, `doc-coauthoring`, `internal-comms`, `remembering-conversations` |

## Normalization applied on install

The bundle mixed two frontmatter conventions. Claude Code only accepts
`name`, `description`, `license`, `allowed-tools`, `compatibility` and `metadata`,
requires `name` to be a lowercase-hyphen slug matching the directory, and caps
`description` at 1024 characters. 32 skills needed fixing to load:

- **Names slugified** (31 skills) — e.g. `Test-Driven Development (TDD)` → `test-driven-development-tdd`.
- **`when_to_use` folded into `description`** (31 skills) — Claude Code only reads
  `description` when deciding whether to trigger a skill, so the trigger text would
  otherwise have been ignored.
- **Non-standard keys moved under `metadata`** (31 skills) — `version`, `languages`, `context`.
- **`claude-api` description trimmed** from 1071 to 747 characters.
- **Broken `firecrawl-cli` cross-links** in 6 firecrawl skills repointed at `firecrawl/`.
- **`pdf/scripts/convert_pdf_to_images.py`** now creates its output directory.

All 60 pass `python3 skill-creator/scripts/quick_validate.py <skill>`.

## Runtime prerequisites

Skills that shell out need these on PATH. They are installed in the current
container but the container is ephemeral — reinstall in a fresh environment.

```bash
# Document, PDF and media skills
pip install python-docx python-pptx openpyxl pypdf pdfplumber PyMuPDF \
            pillow imageio imageio-ffmpeg numpy lxml markitdown \
            pytesseract pdf2image
apt-get install -y libreoffice-writer libreoffice-calc libreoffice-impress \
                   poppler-utils tesseract-ocr

# webapp-testing
pip install playwright        # browser already at /opt/pw-browsers/chromium

# mcp-builder
pip install anthropic mcp

# Firecrawl skills
npm install -g firecrawl-cli@1.19.6
export FIRECRAWL_API_KEY=fc-...
```

### Firecrawl behind the sandbox HTTPS proxy

`firecrawl-cli@1.19.6` bundles axios 1.15.2, which sends plain-HTTP
(non-CONNECT) requests to an HTTPS proxy and gets `405 Method Not Allowed` on
every POST — `--status` works, `scrape`/`search` do not. Upgrading the bundled
axios to >= 1.16.1 fixes it:

```bash
FC=$(npm root -g)/firecrawl-cli/node_modules
cd /tmp && npm init -y && npm install axios@latest https-proxy-agent
for d in /tmp/node_modules/*; do rm -rf "$FC/$(basename $d)"; cp -a "$d" "$FC/"; done
```

The API key belongs in `.claude/settings.local.json`, which is gitignored:

```json
{ "env": { "FIRECRAWL_API_KEY": "fc-..." } }
```
