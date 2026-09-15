# BI HTML Dashboard Template

Shared design system and component catalog for the Risk BI / Pulse family of
self-contained HTML dashboards.

**These files aren't a codebase you open and edit. They're context you hand to
Claude.** You attach them, describe the dashboard you want, and Claude does the
copying, wiring and adapting — producing a single self-contained `.html` file
with its dataset embedded as a CSV string and parsed client-side. No server, no
build step, no external dependencies at runtime.

📖 **[Read PROMPT_GUIDE.md](PROMPT_GUIDE.md)** — the actual how-to: what to
attach, what to say, what's fair game to ask for and what to push back on.
Start there.

---

## Quick start

1. Create a **Claude Project** for this dashboard system.
2. Add all four working files below to its project knowledge — once, not per
   chat.
3. In a new chat inside that project, describe what you're building.

> "Build me a [X] dashboard using this exact structure — keep the topbar,
> sidenav and filters as-is, add sections for [Y, Z], and use the existing
> button/badge/pill classes already defined rather than inventing new styles."

Preloading all four matters: it means Claude can flag *"there's already a driver
table widget for this"* instead of you having to remember the catalog exists.

---

## What's in here

| File | What it is | When it's used |
|---|---|---|
| `boilerplate.html` | The page shell — topbar, sidenav, filters, page titles, section switching. Already wired and working. Design tokens are inlined, so this one file covers the visual language on its own. | **Always.** Every new dashboard starts here. |
| `design-tokens.css` | The shared palette, fonts, spacing, shadows and radii (Flutter brand colours). | Reference; already embedded in the boilerplate. |
| `widgets.css` | Opt-in catalog of fancier per-report CSS components — chart modals, driver tables & waterfalls, pyramid/metric grids, definition tables, threshold chips, brand filtering, dropzones, and more. Not linked by the boilerplate. | When you want a component beyond the base shell. |
| `widget-scripts.js` | The same idea, for widgets needing real JS: PPTX export engine, Excel/CSV ingestion helpers, in-browser AI data assistant, MoM/Mo3M/YoY lag math, brand-split overlays, data dictionary rendering. | When you want behaviour, not just styling. |
| `PROMPT_GUIDE.md` | The team guide — how to prompt, set up, retrofit, and what not to touch. | Read it. |
| `PROMPT_GUIDE.pdf` | Same guide, plus rendered visual samples of the main widgets with placeholder data — handy for pointing and saying "like this one." | When you'd rather see a widget than describe it. |

---

## The one rule

**Attach `boilerplate.html` and build from it. Don't start from a blank chat,
and don't ask Claude to "match the style of Product Pulse" from memory.**

A description gets you something *similar*. The file gets you something
*identical*. Everything in the guide assumes you're already doing this.

---

## Where the components came from

`widgets.css` and `widget-scripts.js` are extracted from three shipped
dashboards — Product Pulse, Commercial Pulse, and UK Brand Pulse. Each block is
labelled with whether it's **verbatim-identical across sources** or
**single-source** (built for one report only).

That distinction matters when you're picking a widget: single-source components
are still safe to use, but they've only been proven against one report's data
shape, so they're worth a second look. Ask Claude which category a widget falls
into — it's noted inline in the files.

One exception worth knowing: the *"export chart as raw HTML"* widget was never
actually built in any source dashboard. It's written fresh rather than
extracted, so treat it as a starting point to test, not as proven.

---

## Adding back to the catalog

If you build something genuinely reusable that isn't in the catalog, ask Claude
to add it back into `widgets.css` / `widget-scripts.js` in the same labelled
format — so the next dashboard doesn't rebuild it from scratch. That's how this
library grew in the first place.

---

Maintained by Risk BI · questions in
[#uki-risk-bi-helpdesk](https://flutter.enterprise.slack.com/archives/C08G45XQK8C)
