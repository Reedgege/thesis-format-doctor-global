# PaperFormat Pro · Formatting Scope & Boundaries (APA 7th + MLA 9th, English)

> Audience: users of **PaperFormat Pro (overseas edition of "Mr. Document Formatter")** — the APA 7th *Student* and *Professional* editions, the MLA 9th edition (student course papers / theses), and our own product, operations, and support teams.
> Purpose: to state clearly **what we do, what we do not do, what prompts you must read, and what you must do yourself** — so that no overseas customer mistakes this for a "one-click submit" or "auto-grade" tool.

---

## 0. One-line positioning

PaperFormat Pro is an **offline, zero-network, zero-LLM** **formatting assistant**. It checks and auto-fixes your document's **layout/formatting** (margins, line spacing, alignment, indentation, fonts, heading levels, reference list layout, etc.) against the chosen citation style (e.g., APA 7th). It **does not touch or rewrite your body content**, and it **is not responsible for the correctness of your citation content**.

It is not a plagiarism checker, not a grammar checker, not an auto-grading system, and not a "one-click submit" button.

---

## 1. What we DO (layout/formatting layer — commercially graded and verified)

These capabilities are built into the rule engine, backed by regression tests, and **verified with zero false-positives** on the official APA annotated samples (`student-paper-annotated.docx` / `professional-paper-annotated.docx`).

### 1.1 Page & body layout (APA 7th)
| Dimension | APA 7th rule | Our handling |
|---|---|---|
| Margins | 1″ on all sides (§2.14) | Checked (as a typical value hint, not a hard fail) |
| Line spacing | Double-spaced (§2.21) | Checked + auto-fixable |
| Body alignment | **Left-aligned, NOT justified** (§2.22) | Checked (under APA this is a **warn**) + auto-fixable |
| First-line indent | 0.5″ indent on the first line of each paragraph (§2.24) | Checked (under APA this is a **warn**) + auto-fixable |
| Body font | 6 accessible fonts: Times New Roman / Calibri / Arial / Georgia / Lucida Sans Unicode / Computer Modern (§2.19) | Checked (warns if the font is outside the list) |
| Font size | Typically 12pt (11pt for Calibri/Arial/Georgia, 10pt for Lucida Sans Unicode/Computer Modern) | Checked (as a typical value hint) |

> Note: page-level values (margins / size / spacing) follow our **template-driven principle** — they are only hard-enforced when your school template or the format questionnaire explicitly states them; by default they are treated as typical-value hints and will not arbitrarily fail you.

### 1.2 Heading levels (5 levels, APA 7th §2.26–2.27)
| Level | Format |
|---|---|
| L1 | Centered, bold, title case |
| L2 | Flush-left, bold, title case |
| L3 | Flush-left, bold italic, title case |
| L4 | Indented, bold, title case, **ending with a period (run-in heading)** |
| L5 | Indented, bold italic, title case, **ending with a period (run-in heading)** |

- Checked (under APA this is a **warn**: incorrect heading format is clearly flagged) + auto-applied per the 5-level rules.
- After fixing, regression tests **read back the direct formatting level by level** (L1 centered+bold, L2 left+bold, L3 left+bold+italic, L4 indent+bold+period, L5 indent+bold+italic+period — all confirmed).

### 1.3 Reference list (References)
| Dimension | Rule | Our handling |
|---|---|---|
| Heading format | "References" **bold + centered + on a new page** (§2.9 / §9.2 / §2.17) | Checked (if a heading *style* is used → reported as pass and the style is trusted; if typed manually without bold/centering → flagged) |
| Hanging indent | 0.5″ hanging indent per entry (§9.43–9.44) | Checked + auto-fixable |
| Order | Alphabetical by first author's surname (§9.44) | Checked (**info** hint, not enforced) |
| Journal / book titles | Should be italicized (§6.22 / §10.1) | **info** robust hint (flagged only when *no* italics appear anywhere — will not false-positive on official samples) |
| Author connector | Use "&" before the last author (§9.7) | **info** robust hint (flagged only when ", and " appears in the *author* segment — will not false-positive on "and" inside a book title) |

> Why are journal-italics / "&" connector only **info** and only macroscopic hints? Because they are *content-semantics* judgments (which words are a journal name, which segment is the author). The official annotated samples themselves contain legitimate entries with no DOI and no issue number. We deliberately issue an info hint only on large-scale suspicious signals — we **never rewrite your reference content** and never false-positive on gold-standard papers.

### 1.4 Front matter (see §4 for Student vs Professional differences)
- Title-page structure checks (title / author / affiliation / course / instructor / date, etc.).
- Professional edition: running head (short title + page number), Abstract + Keywords, author note, running head ≤ 50 characters (§2.8).
- Block quotation 0.5″ indent (§8.25–8.27).

---

## 2. What we do NOT do (citation content / syntax — explicitly out of scope)

The following belong to **body content**. By product policy we **never auto-rewrite** them; the tool takes no responsibility for their correctness and does not claim to check them. **You must verify these yourself:**

1. **Author–date in-text citation syntax**: whether (Author, Year) is formed correctly.
2. **et al. rules**: how many authors warrant et al., how many names to spell out.
3. **"&" vs "and" inside parentheses**: in-text parentheses should use "&" not "and" (§8.11).
4. **a/b/c suffixes for same-author-same-year**: disambiguation markers.
5. **Group / anonymous author** handling (§9.11–9.12).
6. **Author name capitalization & inversion**: surname-first vs initials ordering.
7. **20 / 21+ author** omission rules (§9.8).
8. **DOI / URL presence & format** (§9.34–9.35): official samples legitimately allow entries without DOIs, so we deliberately do not check this to avoid false positives.
9. **Article title sentence case** correctness (§6.14 / §10.2).
10. **Issue number** inclusion/exclusion (§9.31).
11. **Reference-to-text correspondence**: missing, extra, or mismatched citations.
12. **Spelling, grammar, academic phrasing, logical coherence** — language quality in general.

> In short: anything about "is the content right / is the syntax compliant" is **your responsibility**. We only work on "is the layout clean / are the levels correct."

---

## 3. Prompts you MUST read (otherwise you waste the tool)

After a run you get a **findings report** graded by severity. **At minimum, read these:**

- **warn (warning)**: under APA these concentrate in three dimensions — `alignment` (must be left, not justified), `first_line_indent` (must be 0.5″), `heading` (5-level format). These are the layout issues you most need to fix.
- **info (hint)**: especially `Reference italics` (journal/book italicized?), `Reference author connector` (should it be &?), `reference_sort` (alphabetical?), `reference_heading` (References bold + centered?). They are not enforced, but reading them spares you criticism from advisors/reviewers.
- **pass (passed)**: these items are already correct; no action needed.

> Technical note: Word "Heading styles" hide bold/centering *inside the style*, which the tool cannot read directly — so if your References or headings use a style (Heading 1–5), the tool **trusts the style** and reports pass; if you typed direct formatting, it checks the direct values. In neither case do we alter your content.

---

## 4. Differences between the two APA editions (pick wrong → check is meaningless)

| Item | Student (course paper) | Professional (submission / publication) |
|---|---|---|
| Running head (short title in header) | Not required | **Required** (≤50 characters) |
| Abstract + Keywords | Not required (unless instructor asks) | **Required** |
| Author note | Not required | **Required** |
| Title page | Title (bold, centered) + Name + Department + Course No.:Course Name + Instructor + Date | Author + Affiliation (superscript) + Author note |
| Body / reference rules | **Identical** | **Identical** |

> If you pick the wrong edition, the Professional check will flag missing running head / abstract / author note; the Student check will not. Choose per your school's / journal's actual requirement.

---

## 5. What you MUST do yourself (pre-submission checklist)

1. **Verify citation content**: for every reference, confirm author, year, title, source, and DOI/URL are real and accurate.
2. **Verify one-to-one correspondence**: every (Author, Year) in the text appears in References, and every References entry is cited in the text.
3. **Read all warn / info prompts from §3**; fix the layout via "auto-fix" or manually.
4. **Read the whole paper through**: spelling, grammar, phrasing, logic — the tool does none of this.
5. **Do a final human review before submitting**: this tool is an assistant only; it is not responsible for your final grade / acceptance.

---

## 6. What it is NOT (please pass this to every customer)

- ❌ Not a "one-click submit" button: after a run you still must review and correct manually.
- ❌ Not an auto-grading / plagiarism system: it does not judge how good your writing is or whether you copied.
- ❌ Not a spelling / grammar checker.
- ❌ Does **not** guarantee your paper is "100% compliant with your school": if your school template differs from typical APA values, follow your school template and state it explicitly in the questionnaire / template.
- ✅ It **is** an **offline, private, repeatable formatting check + auto-fix assistant**.

---

## 7. Verification status (as of this document)

- Rule-engine unit tests (headless): **35 passed** (including APA 5-level heading landing verification and References bold+centered regression tests).
- Official gold-standard samples (APA Student / Professional annotated docs): **17 / 19 findings respectively, with 0 warn / 0 fail** — i.e., **zero false-positives**.
- Font list corrected to the official APA 7th §2.19 six fonts; body alignment / first-line indent hard rules added.
- **MLA 9th rule engine**: **15 MLA-specific unit tests all pass** (including the header-distance ½″ pass / info pair); non-UI regression suite **194 passed / 30 skipped** (skipped = GUI tests that cannot launch a Tk window in this environment, unrelated to the rules). The MLA official samples were cross-checked word-by-word against all 11 Purdue OWL index links, covering every machine-checkable layout rule in G1–G5.

> This document evolves with the software; for rule changes, the in-app findings report and the latest versioned docs prevail.

---

## 8. MLA 9th Scope & Boundaries (parallel to APA)

> MLA 9th is covered by the same layout assistant as APA 7th, but with a **different strictness**: MLA's page / heading dimensions are **always info / pass only — never warn / fail, and never auto-rewrite body content**. This is a deliberate design choice because MLA official samples allow high freedom in school templates — we only give reference hints, never hard judgments.

### 8.1 What MLA checks (all info / pass — never fail)

| Dimension | MLA 9th rule | Our handling |
|---|---|---|
| Margins | 1″ on all sides | Checked (info hint, not enforced) |
| Body font | Times New Roman (common in samples) | Checked (info hint if outside the list) |
| Font size | 12pt | Checked (info hint) |
| Line spacing | Double-spaced | Checked + auto-fixable (layout) |
| Body alignment | Left-aligned | Checked (info) + auto-fixable |
| First-line indent | 0.5″ on first line of each paragraph | Checked (info) + auto-fixable |
| Space after paragraph | 0 (no blank between paragraphs) | Checked (info) + auto-fixable |
| Header distance from top | ½″ (MLA 9th requires the header 0.5″ from the top) | Checked (info / pass, **added 2026-10-18**) |
| Page number + running head | Top-right every page: your last name + space + page number | Checked (info) |
| First-page heading block (G2) | Upper-left before title: name / instructor / course / date, double-spaced, ~4 lines | Checked (info) |
| Title format (G3) | Centered, not bold / italic / underlined | Checked (info) |
| Works Cited on new page (G4a) | Reference list starts on a new page | Checked (info) |
| Works Cited heading (G4b) | "Works Cited" centered, **not bold** (MLA explicitly forbids bold) | Checked (info) |
| Reference order | Alphabetical by author's last name | Checked (info) |
| Reference hanging indent | 0.5″ hanging indent per entry | Checked + auto-fixable |
| Block quotation (G5) | Quotations longer than 4 lines: 0.5″ left indent, no quotation marks | Checked (info) |

> Note: unlike APA, MLA's `alignment` / `first_line_indent` / `heading` are **info only** (not upgraded to warn), because MLA grants more freedom in body layout and school templates vary widely — we do not hard-judge for you.

### 8.2 What MLA fixes (layout only — never body content)

The fixer writes **layout attributes only**; it never rewrites your body text, in-text citations, or reference entry text:

- Margins, font, font size, line spacing, alignment, first-line indent, space after paragraph
- Block-quotation indent (0.5″ left indent, strip quotation marks)
- Reference hanging indent (0.5″)
- Title-page layout (MLA title / author / affiliation direct formatting)
- **Does not** apply numbered heading formats: MLA `heading_levels=0` — no I./1. style numbered headings, so your section heading levels are left untouched.

### 8.3 What MLA does NOT check (citation content — explicitly out of scope)

Same as APA §2 — these belong to **body content**; by policy we **never auto-rewrite** them; the tool takes no responsibility for their correctness:

1. In-text parenthetical citation syntax: whether (Author page) is formed correctly.
2. et al. rules.
3. Italics vs quotation marks for book / article titles.
4. Publication-info order and formatting (Author. Title. Publisher, Year.).
5. DOI / URL presence & format.
6. Title case correctness.
7. Reference-to-text correspondence.
8. Spelling, grammar, academic phrasing, logical coherence — language quality in general.

### 8.4 MLA prompts you MUST read

An MLA run produces a findings report with **no warn / fail** — only info / pass. At minimum read these info hints:

- `running_head` (last name + page number top-right every page)
- `first_page_heading` (G2 first-page block complete?)
- `title_format` (G3 title centered, not bold?)
- `reference_heading` (Works Cited centered, not bold — MLA forbids bold)
- `reference_sort` (alphabetical?)
- `block_quote` (G5 long quotations indented and unquoted?)

They are not enforced, but reading them spares you criticism from advisors / reviewers.
