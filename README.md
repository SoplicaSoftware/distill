# distill

**Turn heavy documents into cheap, structured input for Claude** — and get back a work plan or a study-ready concept map.

`distill` is a [Claude skill](https://docs.claude.com/en/docs/agents-and-tools/agent-skills/overview) plus a few small Python scripts. Instead of feeding Claude whole slide decks as images or raw XML, the scripts extract **compact text structure** (text, layout, arrows, which comment points at which part of a screenshot) and render **only the visuals that matter**. Claude reads less, sees more precisely, and you burn fewer tokens.

Built for people who work or study with lots of documents: engineers reviewing client feedback decks, developers turning annotated screenshots into tickets, students turning lecture slides into notes and flashcards.

## Modes

| Mode | Command | Input | Output |
|---|---|---|---|
| `map` (default) | `/distill deck.pptx` | any deck | answer in chat, based on `map.md` + selected renders |
| `work` | `/distill feedback.pptx work` | requirement decks, annotated screenshots, specs | **work plan** (tasks, acceptance criteria, dependencies, estimates, open questions) + `tasks.json` |
| `edu` | `/distill lecture.pdf edu` | lecture slides / handouts (PDF or PPTX) | **`concepts.json`** + `images/` + `slides/` — concepts with definitions, linked figures, relations, flashcards, and **what still needs to be looked up** |

### `edu` output
```
lecture-04-distill/
├── concepts.json      # schema: schemas/edu-v1.schema.json
├── images/s04_f2.png  # figures cropped from slides, linked to concepts
└── slides/s04.png     # full slide renders
```
Every concept has a `status`:
- `defined` — the definition is in the material,
- `partial` — only an example, a figure or half a sentence,
- `mentioned` — only named; `to_find` says what to look up.

Definitions come **only from the source**. Claude's own remarks (e.g. a likely error on a slide) go to a separate `model_note`. See [`examples/edu-excerpt.json`](examples/edu-excerpt.json). The format is meant to feed your own reader, mind-map tool or flashcard app (Anki export is an easy next step).

### What the PPTX map looks like
```
## Slide 3
TEXT [x85.9–100% y60.3–71.1%]: Please remove this and move the last-update record
IMG1: picture x6.8–85.6% y19.2–76.6%
IMG3: snippet x66.1–76% y21.8–23.9% on IMG1 @(82%,6% of image)
ANNOTATIONS:
- arrow FROM (93.9,78.5) on IMG4 TO (76,22.9) on IMG3 @(100%,52% of image)
- rect red [x6.5–10.8% y20.2–26.6%] centre on IMG1 @(2%,7% of image)
RENDER: slide-03.png
```

## Why: token cost

Rough numbers measured on real files (image ≈ width×height/750 tokens):

| Scenario | Without distill | With distill | Saving |
|---|---|---|---|
| 7-slide feedback deck, 3 slides need a look | ~22k | ~9k | ~60 % |
| 30-slide deck, 5 slides relevant | ~68k | ~20k | ~70 % |
| 30-slide deck, text-only question | ~68k | ~12k | ~80 % |

Plus quality: exact text instead of text read from pixels, and unambiguous "this arrow points at that button".

## Install

**Claude (app / claude.ai):** add [`skill/SKILL.md`](skill/SKILL.md) as a skill, e.g. paste it into a chat and ask Claude to save it as a skill. The skill clones this repo into its sandbox on first use, so scripts are always up to date.

**Claude Code:** copy `skill/SKILL.md` to `~/.claude/skills/distill/SKILL.md`.

Requirements in the sandbox / locally: Python 3, `python-pptx`, `pdfplumber`, `Pillow`, LibreOffice (`soffice`), poppler (`pdftoppm`, `pdfinfo`). Claude's sandbox normally has all of them.

## Use the scripts directly
```bash
python3 scripts/pptx_map.py deck.pptx -o out/deck --slides 2-5        # map.md + slide-NN.png
python3 scripts/crop.py out/deck/deck.pdf 3 66 18 86 26                # zoom: page 3, x0 y0 x1 y1 in %
python3 scripts/edu_extract.py lecture.pdf -o out/lecture             # slides.md, extract.json, images/, slides/, contact/
```

`edu_extract.py` handles handouts (2–6 slides per page) by detecting slide frames, crops figures from the page render (so formulas with transparency masks come out right), detects vector diagrams, removes logos repeated across slides, and flags slides whose text layer is garbled (chemical / math formulas).

## Repository layout
```
scripts/     pptx_map.py, crop.py, edu_extract.py
skill/       SKILL.md — the skill itself
schemas/     edu-v1.schema.json
templates/   work-plan.md
examples/    edu-excerpt.json
```

## Roadmap
- [ ] `edu`: Anki / CSV flashcard export, Markmap / Mermaid mind-map export
- [ ] DOCX and XLSX maps
- [ ] PDF `work` mode for scanned specs
- [ ] Merging concept maps from several lectures into one course map

## Contributing
Issues and PRs welcome. A new format = a script in `scripts/` that prints compact text structure + a mode step in `skill/SKILL.md`. Please don't commit real lecture or client material, as it is usually someone else's copyright.

## License
MIT © [Soplica Software](https://github.com/SoplicaSoftware)
