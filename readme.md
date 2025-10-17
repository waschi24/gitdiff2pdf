# GitDiff2PDF

Generate a clean, PR‑style **PDF** from normal **unified `git diff`** output.

- **Unified view (default)** like Bitbucket: `-` lines in **red**, `+` lines in **green**, context **gray**  
- Optional **side‑by‑side** view  
- Crisp typography with **system fonts** (Consolas / Segoe UI on Windows; DejaVu on Linux) and safe fallbacks  
- **Header/Footer**, **page numbers**, **line numbers**, **file badges**, and **hunk headers**  
- Thoughtful pagination: **keep-together** (avoid splitting a small file across pages), **widow/orphan protection**  
- Robust parsing and **encoding handling** (UTF‑8 & UTF‑16 PowerShell redirects)

---

## Why?

For exams (e.g., **IPA EFZ**) or audits, you often must submit the **changed code**. PDFs are easy to archive, share, review, and annotate. This tool turns your `git diff` into a **readable, paginated PDF** that feels like a code review tool.

---

## Features

- **Unified view** (default): One column; deletions in red, additions in green, context lines gray  
- **Side-by-side view** (optional): Old (left) | New (right)  
- **Grouping**: by **file → hunk**  
- **Line numbers** (gutter, subtle)  
- **Subtle backgrounds** and **color bars** like modern PR UIs  
- **Header & Footer**: title + timestamp; `Page X / N`  
- **Pagination**:
  - **Keep-together** (unified): a file that fits on next page won’t be split at the end of the previous page  
  - **Widow/Orphan protection** so hunk headers don’t appear alone at page bottoms  
- **Fonts**: uses system TTFs if available; otherwise falls back to Courier safely (no font exceptions)  
- **Encoding**: handles UTF‑8, UTF‑8‑BOM, and PowerShell’s UTF‑16 LE automatically  
- **Sanitization**: removes BOMs, zero‑widths, NBSP/rare spaces, and common “dot/ellipsis” artifacts at the start

---

## Requirements

- **Python 3.9+**
- **PyMuPDF** (aka `pymupdf`)

Install:

```bash
pip install pymupdf
```

> **Windows (PowerShell):**
>
> ```powershell
> py -m pip install pymupdf
> ```

---

## Quick Start

### 1 Create a diff

**PowerShell 5.1** (ensure UTF‑8!):

```powershell
git diff <from-commit> <to-commit> -- . ':(exclude)path/to/big/folder' |
  Out-File -Encoding utf8 compare.diff
```

**PowerShell 7+ / Git Bash / WSL / Linux / macOS**:

```bash
git diff <from-commit> <to-commit> -- . ':(exclude)path/to/big/folder' > compare.diff
```

Or pipe directly to the script (no temp file):

```bash
git diff <from-commit> <to-commit> | python gitdiff2pdf.py - -o diff.pdf --title "Changed Code"
```

---

### 2 Generate the PDF

**Unified (Bitbucket‑style):**

```bash
python gitdiff2pdf.py compare.diff -o ipa-unified.pdf --title "Changed Code (Unified)" --landscape
```

**Only changes (hide context):**

```bash
python gitdiff2pdf.py compare.diff -o ipa-unified-changes.pdf --view unified --hide-context
```

**Side-by-side:**

```bash
python gitdiff2pdf.py compare.diff -o ipa-sbs.pdf --view side-by-side
```

---

## CLI Options

- `inputs`  
  One or more diff files, or `-` for STDIN.

- `-o, --output` **(required)**  
  Output PDF path.

- `--title`  
  Document title (header).

- `--view unified|side-by-side`  
  Default: `unified`.

- `--hide-context`  
  Show only changed lines (`+`/`-`) in unified view.

- `--landscape`  
  A4 landscape (useful for long lines).

- `--font-size` *(float, default 9.5)*  
  Monospace font size for code & numbers.

- `--tabsize` *(int, default 4)*  
  Tab expansion (tabs → spaces for width measurement).

- `--theme light|dark`  
  Light is default; dark is optimized for on-screen viewing.

- `--debug`  
  Print parser debug info (stderr).

- Font overrides (optional; Windows examples):
  - `--mono-font-file "C:\Windows\Fonts\consola.ttf"`
  - `--mono-bold-font-file "C:\Windows\Fonts\consolab.ttf"`
  - `--ui-font-file "C:\Windows\Fonts\segoeui.ttf"`
  - `--ui-bold-font-file "C:\Windows\Fonts\segoeuib.ttf"`

---

## Output Layout & Pagination

- **Unified view**  
  - File “badge” (blue band) **close** to the first hunk header  
  - Small gap between **blue hunk header** and **code**  
  - Subtle **block gaps** between hunks / files (to avoid cramping)
  - **Keep-together (unified)**: before rendering a file, the script **measures its height**.  
    - If it **fits on the current page** → render here  
    - If it **doesn’t fit here but would fit on a *new* page** → **page break before the file**  
    - If it **doesn’t fit on a single page** → normal flow (split across pages)
  - **Widow/Orphan protection** ensures the hunk header is not stranded at a page bottom without at least a few code lines following.

- **Side-by-side view**  
  - Old (left) | New (right) with line numbers per side  
  - (Keep-together heuristic is implemented for **unified**; side-by-side uses continuous flow)

---

## Fonts & Typography

- Automatically uses system fonts if available:
  - **Windows**: Consolas (mono), Segoe UI (UI)
  - **Linux**: DejaVu Sans Mono (mono), DejaVu Sans (UI)
- If fonts can’t be resolved or names are unsafe (contain spaces), the script falls back to **Courier** / **Courier‑Bold** and **avoids font errors**.
- All width measurements use the actual font to keep wrapping and alignment precise.

---

## Encoding & Diff Tips (Windows PowerShell)

- **PowerShell 5.1** redirects (`>`) in **UTF‑16 LE** by default → **use `Out-File -Encoding utf8`** or pipe into the script.
- The script **auto-detects** UTF‑8/UTF‑8‑BOM/UTF‑16LE/BE, but making the diff **UTF‑8** is cleaner.
- **Not supported**: `--word-diff`, `--name-only`, `--name-status`, or pure binary-only changes (no hunks).

Recommended:

```powershell
# UTF‑8 redirection (PS 5.1)
git diff <from> <to> | Out-File -Encoding utf8 compare.diff
```

---

## Troubleshooting

### “No parsable diffs found.”

- Ensure it’s a **unified diff** containing `@@` hunk headers (e.g., `git diff -U3` or `git show <commit>`).
- Not supported: `--word-diff`, `--name-only`, `--name-status`.
- Check encoding (especially PS 5.1). Try `--debug`.

### “bad fontname chars {' '}" or “need font file or buffer”

- You passed a font with spaces or a font that can’t be loaded.  
  The script already **falls back** to Courier.  
  To force specific fonts, use `--mono-font-file` / `--ui-font-file` with full TTF paths.

### Weird dots / ellipsis at the top (`···`, `…`, `•`)

- The script removes BOMs, zero‑widths, NBSP/rare spaces, and common dot/ellipsis bullets.  
  If you still see such characters inside actual code content, share a snippet (they might be **part of the code**).

### Broken alignment or wrapping**

- Use `--landscape` for longer lines.  
- Adjust `--font-size` (e.g., 9.0–10.5).  
- Increase `--tabsize` for better visual alignment if your diffs use tabs.

---

## Known Limitations

- Requires **unified diffs** with `@@` hunks.  
- **Keep-together** is implemented for **unified** mode; **side-by-side** uses flowing layout (can be added if needed).
- Inline **word-level** highlights (within a single line) are not included (can be added as a future enhancement).

---

## Example Workflows

**Compare a commit range, exclude large folders, and export PDF (Unified):**

```powershell
git diff <a-commit> <b-commit> -- . ':(exclude)src/BigAssets' |
  Out-File -Encoding utf8 compare.diff

python gitdiff2pdf.py compare.diff -o changed-code.pdf --title "Changed Code (Unified)" --view unified --landscape
```

**Direct pipe to STDIN:**

```bash
git diff <a-commit> <b-commit> | python gitdiff2pdf.py - -o diff.pdf --title "Changed Code"
```
