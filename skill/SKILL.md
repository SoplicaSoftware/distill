---
name: distill
description: Token-efficient analysis of documents dropped into chat (pptx, pdf; more to come). Modes - map (quick look), work (actionable work plan), edu (JSON concept map with figures for study notes). Use for /distill FILE [mode].
---

# /distill — turn heavy documents into cheap, structured output

Usage: `/distill [FILE] [map|work|edu] [extra: slides 3-5, focus, language]`
- FILE: path, usually `/mnt/user-data/uploads/...`. Missing → newest file in `/mnt/user-data/uploads/`.
- Mode missing → `map`. Answer in the user's language.

## Rules (all modes)
1. Never read binaries or raw XML/HTML. Run a script that outputs **compact structure as text** first.
2. Look at images only where text is not enough. An image costs ≈ w×h/750 tokens: renders ≤1400 px.
3. Small text in an image → crop the region (`crop.py`), never re-render everything larger.
4. Read selectively: map first, then only the renders/crops the task needs. Contact sheets let you see many figures in one read.
5. Cache in `/tmp/distill/<name>/`. If the output already exists, do not convert again.

## Tools (once per session)
If `/tmp/distill/repo/scripts` is missing:
```bash
git clone -q --depth 1 https://github.com/SoplicaSoftware/distill /tmp/distill/repo
```
If cloning fails and `mcp__remote-devices__*` tools exist (chat linked to the user's computer), stage `~/Documents/GitHub/distill/scripts/*.py` with `device_stage_files` and copy them from `stagedPath` to `/tmp/distill/repo/scripts/`. If both fail, tell the user.
Needs: `soffice`, `pdftoppm`/`pdfinfo`, `python-pptx`, `pdfplumber`, `Pillow`. Without `soffice` use `--render none`. Below `S=/tmp/distill/repo/scripts`.

## Mode `map` — quick structured look (PPTX)
1. `python3 $S/pptx_map.py FILE.pptx -o /tmp/distill/<name> [--slides 2-5] [--render auto|all|none]` (for a PDF use `edu_extract.py` and read `slides.md`)
2. Read `map.md`: coords = % of slide (x→right, y→down); `TEXT [...] on IMG1 @(x%,y% of image)` = comment on a screenshot; `picture` = screenshot, `snippet` = pasted UI fragment, `on IMGk` = lies on it; `ANNOTATIONS` = arrows `FROM … TO …` with image position, frames/ellipses with centre, lines (often strike-throughs); `RENDER: slide-NN.png` = visual slide, rendered.
3. Open renders only for slides the map does not explain. Unreadable text: `python3 $S/crop.py /tmp/distill/<name>/deck.pdf N x0 y0 x1 y1` (coords from map + ~2 % margin).
4. Reading annotations: arrow from comment to element = the change concerns that element; strike-through / red line = usually "remove".

## Mode `work` — complete work plan
For requirement decks, client feedback with screenshots, specs, audit reports.
1. Do `map` for the whole file. View every slide whose meaning depends on arrows or screenshots (renders + crops). Coverage over savings: every comment must land in a task.
2. Build the plan from `/tmp/distill/repo/templates/work-plan.md`: one task per atomic change (split "move AND rename"), with where (screen → element), current → target state, acceptance criteria, dependencies, priority, estimate in hours (ranges when unsure).
3. Collect every ambiguity or contradiction between slides as an **open question** naming the tasks it blocks; never guess silently.
4. Check: each slide/comment maps to ≥1 task or is listed under "out of scope". Report the counts.
5. Deliver as a document (Docs artifact if available, else `plan-prac.md` / `work-plan.md` file). Also save `tasks.json` (`[{id,title,where,change,acceptance[],depends[],priority,estimate_h,source_slides[]}]`) for import into trackers. In chat: short summary + biggest risks.

## Mode `edu` — concept map for study notes
For lectures, textbooks, course slides (PDF handouts with several slides per page are split automatically).
1. `python3 $S/edu_extract.py FILE -o /tmp/distill/<name>` → `slides.md` (text per slide, figures with `«nearby text»`, `!! TEXT GARBLED` flags), `extract.json`, `images/`, `slides/`, `contact/sheet-N.png`.
2. Read `slides.md`. Read all contact sheets (cheap) to caption figures. Open `slides/sNN.png` only for slides flagged GARBLED (formulas) or where diagrams have no figure.
3. Write `concepts.json` per `/tmp/distill/repo/schemas/edu-v1.schema.json`:
   - one concept per term worth a flashcard; `id` kebab-case; `category` = lecture section.
   - `status`: `defined` (definition in source) / `partial` (only example, figure, list or half-sentence) / `mentioned` (only named). `partial` and `mentioned` **must** have `to_find` saying what to look up. Unfinished slide sentences ("…", "gdy ...") → `to_find`.
   - `definition`: from the source only. Your own knowledge never goes into `definition`; doubts or likely slide errors go to `model_note`.
   - `figures`: only meaningful ones. Drop title banners, section dividers, logos, text rendered as an image, reference lists. Caption every kept figure, link it to concepts.
   - `related` with short relation labels (makes the mind map); every id must exist.
   - `flashcards` for `defined` concepts (1–3 short Q/A). `exercises` for tasks on slides.
4. Build the package `/mnt/user-data/outputs/<name>-distill/`: `concepts.json`, `images/` (kept figures only), `slides/`. Validate JSON (ids, figure files exist). Zip it and send. In chat: counts (defined/partial/mentioned) and the "to find" list.

## Adding formats (contributors)
New `scripts/<format>_*.py` must print compact text structure and write renders only on request. Document it here as a mode step.
