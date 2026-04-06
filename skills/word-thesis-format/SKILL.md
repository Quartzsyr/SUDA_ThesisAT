---
name: word-thesis-format
description: Use when formatting a graduation thesis in Word, including cover pages, heading levels, page layout, fonts, spacing, captions, tables, equations, cross-references, table of contents, section breaks, page numbers, and thesis compliance cleanup. For this workspace, default to the Soochow University undergraduate thesis print and binding format (2020 edition) unless the user provides a newer school rule.
---

# Word Thesis Format

Use this skill for Word formatting and thesis layout cleanup.

For this workspace, default to the Soochow University undergraduate thesis print and binding format (2020 edition). When generic advice conflicts with school-specific rules, prefer the Soochow University rules.

If the user provides an actual school template file such as a cover `.doc` or `.docx`, treat that file as the highest-priority source for the covered section. Do not "beautify" away template-specific spacing, alignment, field order, or table geometry.

## Scope

- Cover page completion and cleanup
- Heading hierarchy and styles
- Chinese and English fonts and sizes
- Paragraph spacing and indentation
- Page margins, headers, footers, and page numbers
- Figure captions, table captions, and cross-references
- Equation numbering and alignment
- Table of contents generation and refresh
- Final compliance check before submission

## Workflow

1. Confirm the formatting source.
   Prefer the university template or explicit school requirements. In this workspace, use the Soochow University 2020 format as the default baseline unless the user explicitly switches to another standard.

2. Decide the authority by section.
   - Cover, originality statement, authorization statement: school template first.
   - Body pages and appendices: school rules first, then the working document styles.
   - If a supplied template conflicts with the generic baseline, keep the template for that section and note the conflict.

3. Normalize styles before micro-fixes.
   Fix built-in styles or a consistent custom style system first. Do not patch every paragraph manually unless the document is tiny.

4. Enforce hierarchy.
   Typical pattern:
   - thesis title
   - chapter titles
   - section titles
   - subsection titles
   - body text
   - captions
   - references

5. Standardize objects.
   - Figures: unified caption position, numbering, and spacing.
   - Tables: unified title position, borders, alignment, and notes.
   - Equations: centered formula with right-aligned number if required.
   - References: one format only, no mixed punctuation styles.

6. Final cleanup.
   Check page break logic, blank lines, orphan headings, inconsistent numbering, punctuation width, and full-width/half-width misuse.

## Cover Page Workflow

When the user provides a cover file, use this workflow instead of giving generic cover advice:

1. Preserve the original template structure.
   Do not rebuild the cover from scratch if the provided `.doc` or `.docx` already contains positioned text, lines, tables, or WordArt. Edit inside the template unless it is clearly broken.

2. Identify what is fixed and what is editable.
   Typical fixed elements:
   - school name
   - document type such as `本科毕业设计（论文）`
   - decorative rules, borders, logo, or fixed labels
   Typical editable elements:
   - thesis title
   - student name
   - student number
   - college or department
   - major
   - class or grade
   - supervisor
   - completion date
   Always verify the actual field names from the template before renaming anything.

3. Match by visual baseline, not only by font size.
   For cover pages, preserve:
   - line positions
   - center alignment versus distributed alignment
   - underline length
   - textbox width and internal margins
   - character spacing and line spacing
   - table row height, cell width, and vertical alignment
   A cover can look wrong even when the nominal font size is correct.

4. Prefer template-native mechanisms.
   If the cover uses tables, tab stops, text boxes, shapes, or underlined placeholder runs, keep using that same mechanism. Do not replace a table layout with spaces or repeated underscores.

5. Keep title fitting controlled.
   If the thesis title is long, prefer this order:
   - adjust line breaks at semantic boundaries
   - slightly tune paragraph spacing or textbox width if the template allows it
   - only then consider a small font reduction
   Avoid shrinking the title aggressively just to force one line.

6. Protect compatibility for old `.doc` files.
   Many school cover templates are Word 97-2003 `.doc` files. Avoid operations that commonly break them:
   - wholesale copy-paste from rich web content
   - replacing layout objects with newer SmartArt or content controls
   - converting aligned fields into manually spaced text

## Concrete Cover Template Notes

For the supplied cover template `D:\syr\毕设\相关文档\本科毕业设计（论文）封面-word版本.docx`, the stable layout cues are:

- The page title is `本 科 毕 业 设 计（论 文）`.
- The title paragraph is centered.
- The title run uses `方正小标宋_GBK` and size `64` half-points.
- The editable information area is a single table, not free text.
- Extracted field labels include:
  - `学院(部)`
  - `题   目`
  - `年  级`
  - `专业`
  - `班  级`
  - `学号`
  - `姓  名`
  - `指导老师`
  - `职称`
  - `论文提交日期`
- In this template, labels are mainly Song at size `30` half-points.
- Several values, including the thesis title and student data, use HeiTi around size `28` to `30` half-points.

When helping with this exact cover style, prefer editing the existing table cell content rather than rebuilding the form.

## Cover-Specific Output Pattern

When the user asks for cover help, provide:
- which elements should remain exactly as in the template
- which fields can be edited and how to enter them cleanly
- which visual details must be rechecked manually in Word before submission

If the user supplies a concrete cover file path, mention that file as the working authority for the cover section.

## Soochow University 2020 Baseline

Apply the following rules by default for this thesis project:

- Printing: A4, single-sided.
- Cover and declarations: use the school-provided electronic templates.
- Margins: top 3.3 cm, bottom 2.7 cm, left 2.5 cm, right 2.5 cm.
- Binding line: left side, 0.5 cm.
- Header distance: 2.6 cm.
- Footer distance: 2.0 cm.
- Line spacing: 1.5 lines.
- Paragraph spacing: before 0 pt, after 0 pt.
- First-line indent: 2 characters.
- Header: centered, small-five Song, text should be `苏州大学本科生毕业设计（论文）`, with a horizontal line.
- Body font: Chinese in Song, English in Times New Roman.
- Body size: small-four.
- Heading font: bold.
- Chapter title size: small-three.
- Leave one standard blank line between headings and the following body text.

## Soochow University Front Matter and Order

For undergraduate thesis binding order, use:

1. Cover
2. Originality statement and authorization statement
3. Thesis
4. Task assignment
5. Foreign-language literature and Chinese translation
6. Literature review or reading report
7. Midterm inspection form
8. Defense record form
9. Grade evaluation form
10. Detection report (concise version)

Inside the thesis package, use:

1. Chinese title, abstract, and keywords
2. English title, abstract, and keywords
3. Table of contents
4. Main text
5. Conclusion
6. References
7. Acknowledgements
8. Appendices

## Soochow University Font Rules

- Chinese title: small-two, HeiTi.
- Chinese abstract and keywords: small-four, Song.
- English title: small-two, Times New Roman.
- English abstract and keywords: small-four, Times New Roman.
- TOC main entries: four-point Song by default.
- In a chaptered TOC, chapter titles use four-point HeiTi and section titles use four-point Song.

## Soochow University Page Number Rules

- No page numbers on the cover, originality statement, or authorization statement.
- Abstract and table-of-contents pages use lowercase Roman numerals in continuous order: `i`, `ii`, `iii`.
- Start Arabic numbering from the main text.
- Main-text page numbers should be centered and written with hyphens on both sides, such as `-1-`.
- References, acknowledgements, and appendices continue the Arabic numbering from the main text.
- Foreign-language translation and literature review follow the same body formatting rules but do not need page numbers.

## Soochow University Citation Rules

- In-text citations use upper-right bracketed numbers, such as `[1]`.
- Multiple citations in one place use one pair of brackets with English commas between numbers.
- Consecutive citation ranges use a short hyphen.
- Repeated citation of the same source should appear only once in the final references list.
- Reference formatting should follow GB/T 7714-2015.

## TOC and Heading Guidance

- If the thesis is chaptered in science and engineering style, prefer forms such as `第1章` and `1.1`.
- The table of contents should display starting page numbers and connect title text to the page number with dot leaders.
- Keep heading numbering stable across the document before refreshing the TOC.

## Preferred Principles

- Use styles, not manual visual patching, whenever possible.
- Keep one numbering system for chapters and captions.
- Keep Chinese typography consistent and readable.
- Avoid mixed fonts, mixed indent logic, and mixed caption patterns.

## Common Checks

- Does the cover page still match the school template after text replacement?
- Are editable cover fields aligned to the original placeholder geometry?
- Are heading levels mapped consistently?
- Do TOC entries match actual headings?
- Are captions sequential and cross-references correct?
- Are page numbers suppressed where the template requires?
- Are abstract, keywords, references, acknowledgements, and appendices placed correctly?
- Do margin, header, footer, and page-number settings match the Soochow University 2020 rules?

## Output Patterns

When helping with Word formatting, provide:
- the formatting rule to apply
- the exact objects or sections to normalize
- a short list of compliance risks still needing manual confirmation against the school template

If exact school wording or a compact checklist is needed, read `references/suda-2020-format.md`.
