# UI/UX Pro Max – Complete Skill Definition

## Overview

**UI/UX Pro Max** is a design intelligence skill providing searchable guidance across UI structure, visual design, interaction patterns, and user experience quality. It contains **79 styles, 192 product palettes, 74 font pairings, 119 UX guidelines, 105 icons, 17 GSAP presets, 25 chart types, and 22 technology stacks**.

## When to Use

Apply this skill for tasks involving **UI structure, visual design decisions, interaction patterns, or user experience quality control**—designing pages, creating components, choosing color/typography/spacing, reviewing UI for accessibility/consistency, implementing navigation/animation/responsive behavior, or improving usability.

Skip it for pure backend logic, API/database design, infrastructure, or non-visual scripts—unless the task changes how something "looks, feels, moves, or is interacted with."

## Priority Rule Categories (1–10)

| Priority | Category | Key Focus |
|----------|----------|-----------|
| 1 | **Accessibility** | Contrast 4.5:1, alt text, keyboard navigation, ARIA labels |
| 2 | **Touch & Interaction** | 44×44px minimum, 8px+ spacing, loading feedback |
| 3 | **Performance** | WebP/AVIF, lazy loading, CLS < 0.1 |
| 4 | **Style Selection** | Product-type match, SVG icons, consistency |
| 5 | **Layout & Responsive** | Mobile-first, viewport meta, no horizontal scroll |
| 6 | **Typography & Color** | 16px base, 1.5 line-height, semantic tokens |
| 7 | **Animation** | Context-aware timing, meaningful motion, reduced-motion support |
| 8 | **Forms & Feedback** | Visible labels, error proximity, helper text |
| 9 | **Navigation Patterns** | Predictable back, ≤5 bottom nav items, deep linking |
| 10 | **Charts & Data** | Legends, tooltips, accessible color encoding |

## Core Workflow

### Step 1: Analyze Requirements
Extract product type, audience, style keywords, and detected technology stack from `package.json`, `pubspec.yaml`, `*.xcodeproj`, or framework markers. **Never assume a stack**.

### Step 2: Generate Design System (for new projects)
```bash
python "${CLAUDE_PLUGIN_ROOT}/.claude/skills/ui-ux-pro-max/scripts/search.py" \
  "<product_type> <industry> <keywords>" --design-system [-p "Project Name"]
```
Returns pattern, style, colors, typography, effects, and anti-patterns.

**Persist across sessions:**
```bash
python "${CLAUDE_PLUGIN_ROOT}/.claude/skills/ui-ux-pro-max/scripts/search.py" \
  "<query>" --design-system --persist -p "Project" --output-dir "<root>"
```
Creates `design-system/<project-slug>/MASTER.md` and optional page overrides.

**Design Dials** (optional tuning sliders):
- `--variance 1-10`: Centered (1-3) to Bold (8-10)
- `--motion 1-10`: Subtle to Complex choreography
- `--density 1-10`: Spacious to Dense/dashboard layouts

### Step 3: Targeted Searches (as needed)
```bash
python "${CLAUDE_PLUGIN_ROOT}/.claude/skills/ui-ux-pro-max/scripts/search.py" \
  "<keyword>" --domain <domain> [-n <max_results>]
```

**Common domains:** `product`, `style`, `color`, `typography`, `ux`, `landing`, `icons`, `gsap`, `chart`, `react`, `nextjs`, `vue`, `svelte`, `nuxtjs`, `html-tailwind`, `swiftui`, `flutter`.

### Step 4: Stack-Specific Implementation
```bash
python "${CLAUDE_PLUGIN_ROOT}/.claude/skills/ui-ux-pro-max/scripts/search.py" \
  "<keyword>" --stack <stack>
```

## Search Best Practices

- **One dominant intent** per query (2–5 meaningful terms)
- **Add constraints** (product type, platform, interaction)
- **Retry once** with narrower terms or explicit domain if results are empty or off-topic
- **Do not persist unverified output**
- For accessibility: search the **semantic outcome first** (`"error summary validation" --domain ux`), then stack details
- For text/compact components: search **UX outcome**, then stack implementation

## Output Formats

- `-f ascii` (default, terminal display)
- `-f markdown` (documentation)
- `--json` (machine-readable)

## Pre-Delivery Checklist

Read `references/pro-rules.md` before shipping app UI. Covers icon discipline, interaction feedback, contrast, safe-area layout, and accessibility for native/mobile apps.

## Handling Empty Results

1. Retry with narrower phrasing or explicit domain/stack
2. If still empty, fall back to built-in defaults (§1–10 priority table) and explicitly label it as a general recommendation, not a database match
3. Never fabricate output from a 0-result search
