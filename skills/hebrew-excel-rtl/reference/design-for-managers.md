# Designing a workbook a manager can read

The audience is someone senior, in a hurry, who did not build the file and does not think
in spreadsheets. They will look at it for ninety seconds and form an opinion.

---

## The colour rule

**One accent colour. Colour on a number means "this needs attention" and nothing else.**

Most bad workbooks fail here. Each concept gets its own colour, so after five concepts the
grid is a rainbow and nothing stands out — which means nothing is emphasised at all.

| Use | Colour |
|---|---|
| Header band | the accent, dark, white text — the only dark fill on the sheet |
| Section band | the accent at ~10% — `E8EDF4` |
| Data rows | white, or a near-invisible zebra tint `F7F9FC` |
| All numbers | one ink colour `1A1A1A` |
| Captions, units | grey `6B7280` |
| Problem | red `DC2626` |
| Warning | amber `D97706` |
| Good | green `16A34A` |
| Worth knowing, not urgent | blue `2563EB` |

Amber and blue must not be interchangeable. "About to run out" and "far too much stock" are
opposite problems; giving them the same colour makes the column unreadable.

Palette adapted from the `israeli-ui-design-system` tokens, so a workbook and a web
dashboard for the same company look related.

## Do not colour-code inputs vs formulas

The investment-banking convention — blue for hardcoded inputs, black for formulas, green for
cross-sheet links — is a **modeller's debugging aid**. To a manager it looks like the file
has two different kinds of money in it, and it is the first thing they ask about.

Put every input on one clearly-named assumptions sheet instead. Then "what can I change?"
has a location as its answer, not a colour. If you also want the convention, confine it to
that sheet and label it.

## Borders and gridlines

- Turn Excel's gridlines **off** on any sheet you have formatted
  (`ws.sheet_view.showGridLines = False`). They fight with your own rules.
- Horizontal rules only, thin, light grey. Vertical lines inside data add nothing and make
  the sheet look like a form.
- One rule under each row, or none at all with a zebra tint. Not both.

## Never merge inside data

Merged cells break sorting, filtering and pivot tables — the three things a BI workbook
exists to support. Use `Alignment(horizontal="centerContinuous")` across the range instead:
identical appearance, and the range stays clean.

This is also how you stop a large KPI number rendering as `######` — a 16pt number will not
fit a 15-wide column, and centerContinuous lets it spill across the tile without a merge.

## Layout

An RTL sheet puts `A1` at the top **right**. So:

- The label column is `A` — the rightmost, where the eye starts.
- Totals go at the far left (highest column).
- Charts anchored past your data, e.g. `K4`, sit to the **left** of it.

Vertical order on a dashboard, top to bottom:

1. Title, and one line saying what this is and how fresh it is
2. **The numbers that matter** — 4–6 KPI tiles, no more
3. **What you should know** — plain-Hebrew sentences, computed live
4. **What needs action** — a short table, sorted by urgency
5. Charts

A manager who reads only section 2 and 3 should be able to run the meeting.

## KPI tiles

Three rows each: small grey label, large accent number, tiny grey caption saying what it
means. The caption is the part people skip and the part that makes the tile self-explanatory
— `לפי מחיר עלות` under an inventory value answers the question before it is asked.

Four to six tiles. Beyond that they stop being highlights.

## The insight block

This is what separates a report from a data dump. Four to six sentences, each computed by
formula so they cannot drift:

> 1. החנות המובילה: עזריאלי תל אביב — 170,721 ש״ח בחודש.
> 3. הפריט שנגמר הכי מהר: מכנסי ג׳ינס סלים מידה 32 בקניון הקריון — נשארו 0.5 שבועות מלאי.
> 6. הרווח התפעולי הוא 8.6% מההכנסות. זה שולי — כל ירידה במכירות תמחק אותו.

Sentence 6 is the model: state the number, then say whether it is good. An `IF()` that
picks between two verdicts is a legitimate and very high-value use of a formula.

## Status columns

Write the **word**, then colour it. `אזל` / `קריטי` / `נמוך` / `תקין` / `עודף` are readable
with the colour stripped out — printed in mono, or by someone colour-blind. A bare coloured
cell forces the reader to hunt for a legend.

Derive the status from a formula against a threshold on the assumptions sheet, never
hardcode the boundary in the `IF()`.

## Give every number its unit

`שבועות כיסוי` not `כיסוי`. `מלאי (יח׳)` not `מלאי`. `מכירות בחודש (₪)` not `מכירות`.
Ambiguity about units is the single most common cause of a manager misreading a report by
an order of magnitude.

## Rows and spacing

- Header rows 30–32pt tall, data rows 20–22pt. Excel's default 15pt is cramped for Hebrew.
- A blank row between sections. Whitespace is the cheapest readability you can buy.
- Hebrew column widths: roughly `len × 1.3 + 4`, floor 12, cap 50.

## The test

Show it to someone who has never seen it and say nothing. If they ask "what am I looking
at?" the title is wrong. If they ask "is that good?" you are missing the insight block. If
they ask "what do I do?" you are missing the action table.
