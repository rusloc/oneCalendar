<!--
ORIGINAL NOTES (kept for reference):

Create README.MD file, structure it,markdown nicely (check for spelling mistakes before creating the file).

General logic and functionality:
* calendar has two views: Standard & Last
* To toggle between views: click button on top (attach the button SVG)
* when switching between views - filters are reset (dropped)
* each view has the same set of date span blocks: Year, Quarter, Month, Week, Date selection (disabled on Last view)
* buttons on top control: Minimize / Reveal the whole calendar; Collapse / Expand all blocks; Switch views; Reset filters
* in the developer side pane (formatting options) you can:
	> select Accent color (controls the color of the pills with date items & filters info block on top)
	> position controlling buttons on the left side or right side (order changes accordingly so to keep the Minimize button in the corner always)
	> main container border weight & color
	> dates block (the bottom one) accent color / background color



Standard view logic:
* general logic runs from top (Year) to bottom (Date selection)
* if you start selecting your date range with year all below blocks will allign within selected Years span
* the same logic applies if you start selecting from Quarter or Month
* if you select Month first and then decide to select year not applicable to the selected month - the Month filter will be removed and the selection will start from the very beginning (since the Year selection is dominant)
* to select one single item you need to click on it (eg.: "Jan 2026")
* to select several items you can keep on clicking without holding CTRL or SHIFT buttons (you can select 'ragged intervals' eg.: "Sep 2025", "Jan 2026", "Mar 2026")
* you can select a range by clicking the first element and the next one with SHIFT pressed
* when date span is selected you can adjust the span using precise date filters
* precise date filter selection is automatically grades out period not in the selection (eg.: if you selected "Jan 2026" the precise date selection will allow to select dates from January 2026 only)




Last view logic:
* the overall visual structure is similar to Standard view
* main idea is to give a quick tool of selecting previous periods in blocks (excluding current period)
* abbreviation meaning:
	> LY - Last Year
	> LQ - Last Quarter
	> LM - Last Month
	> LW - Last Week
* selection of "last" item does NOT include current period (ie.: "1 LY" does NOT include current year; "3 LW" do NOT include current week and so on)
* date selection is disabled (for precise selection use Standard view)
* only single selection is possible vertically and horizontally (ie.: only "2 LQ" or "1 LY" or "3 LW" can be selected but not any combination of 2 and more)
-->

<p align="center">
  <img src="oneCalendar/assets/icon.png" alt="oneCalendar" width="120" />
</p>

<h1 align="center">oneCalendar</h1>

<p align="center">
  <strong>A hierarchical date slicer for Power BI</strong><br/>
  Filter by Year, Quarter, Month, Week, or exact date range — all in one compact visual.
</p>

<p align="center">
  <a href="https://github.com/rusloc/oneCalendar">GitHub</a> &middot;
  <a href="https://www.linkedin.com/in/kirill-bezzubkin-b6860b235">Author</a>
</p>

---

## Two Views

oneCalendar has two modes, toggled via the **Switch view** button in the toolbar. Switching views **resets all active filters**.

| Standard | Last Periods |
|----------|-------------|
| Select specific years, quarters, months, weeks | Select "last N" periods relative to your data |
| Multi-select with click or Shift+click | Single selection only (one item across all blocks) |
| Date range picker enabled | Date range picker disabled |

---

## Toolbar

<p align="center">
  <img src="oneCalendar/assets/Buttons info.png" alt="Toolbar buttons" width="480" />
</p>

| Button | Action |
|--------|--------|
| **Minimize all** | Collapse or reveal the entire calendar panel |
| **Collapse blocks** | Toggle all blocks (Year, Quarter, Month, Week) open or closed at once |
| **Switch view** | Toggle between **Standard** and **Last periods** mode |
| **Reset** | Clear all selections and remove all filters |
| **Filter alert display** | Shows which filter levels are active: `YR`, `QT`, `MT`, `WK`, `DT` |

> The button position (left or right) is configurable in the formatting pane. The **Minimize** button always stays in the corner regardless of position.

---

## Standard View

<p align="center">
  <img src="oneCalendar/assets/standard view expanded.png" alt="Standard view — expanded" width="480" />
</p>

### How Cascading Filters Work

Filters flow **top to bottom**: Year > Quarter > Month > Week > Date range. Selecting at any level automatically scopes everything below it.

| Action | Result |
|--------|--------|
| Select **2025** in Year | Quarter, Month, Week blocks show only 2025 periods |
| Then select **Q1 2025** | Month and Week narrow to Jan–Mar 2025 |
| Then select **Feb 2025** | Week narrows to weeks within February 2025 |

**Upward override rule:** If you select a Month first (e.g., `Mar 2026`) and then select a Year that doesn't contain that month (e.g., `2025`), the Month filter is dropped and filtering restarts from the Year level. The higher level always takes priority.

### Selection Methods

- **Single click** — toggle one item on/off (e.g., click `Jan 2026` to select it, click again to deselect)
- **Multiple clicks** — keep clicking different items to build a non-contiguous selection (e.g., `Sep 2025`, `Jan 2026`, `Mar 2026` — no Ctrl needed)
- **Shift+click** — select a continuous range between the last clicked item and the current one

### Date Range Picker

The bottom section provides precise from/to date inputs. These **constrain automatically** to the active selection above:

> *Example:* If `Jan 2026` is selected in the Month block, the date picker only allows dates within January 2026.

<p align="center">
  <img src="oneCalendar/assets/standard view collapsed.png" alt="Standard view — collapsed" width="480" />
</p>

<p align="center"><em>Collapsed view — all blocks minimized for a compact footprint. Date range still visible.</em></p>

---

## Last Periods View

A quick-select tool for filtering by **previous periods**, excluding the current one. All ranges are relative to your **dataset's most recent date** (not today's calendar date).

### Abbreviations

| Label | Meaning | Example |
|-------|---------|---------|
| `1 LY` | Last 1 Year | If data ends in 2026 → full year 2025 |
| `2 LQ` | Last 2 Quarters | If data ends in Q1 2026 → Q3 2025 + Q4 2025 |
| `3 LM` | Last 3 Months | If data ends in Mar 2026 → Dec 2025, Jan 2026, Feb 2026 |
| `1 LW` | Last 1 Week | The last complete Mon–Sun week before the dataset's max date |

### Key Rules

- **Excludes the current period.** `1 LY` does **not** include the year of the max date. `3 LW` does **not** include the week containing the max date.
- **Single selection only.** You can pick `2 LQ` **or** `1 LY` **or** `3 LW`, but never a combination. Selecting a new item clears the previous one across all blocks.
- **Date range picker is disabled.** For precise date filtering, switch back to Standard view.

---

## Formatting Options

Accessible from the Power BI formatting pane (paint roller icon):

| Setting | What it controls | Default |
|---------|-----------------|---------|
| **Accent Color** | Selected pill color + filter status badge | `#3b82f6` (blue) |
| **Button Position** | Toolbar alignment — Left or Right | Left |
| **Container Border Weight** | Border thickness around the main panel | `1px` |
| **Container Border Color** | Border color of the main panel | `#e5e7eb` |
| **Dates Block Background** | Background of the date range section at the bottom | `#fafafa` |

---

## Data Setup

Drag a single **Date** column into the visual's data well:

| Field | Role | Required |
|-------|------|----------|
| Date Value | Grouping | Yes |

The visual accepts Date, DateTime, or parseable date strings. No measures required.

---

## Installation

```bash
git clone https://github.com/rusloc/oneCalendar.git
cd oneCalendar
npm install

# Development
pbiviz start

# Package
pbiviz package
# Output: dist/oneCalendar.pbiviz
```

Import the `.pbiviz` file via **Visualizations pane > ... > Import a visual from a file**.

---

## Additional Features

- **Sync Slicers** — filter state propagates across report pages
- **Bookmark support** — selections are saved and restored with Power BI bookmarks
- **Landing page** — branded welcome screen when no data is bound
- **Hybrid filter engine** — uses lightweight min-max range filters for contiguous selections; falls back to enumerated date lists only for non-contiguous gaps
- **Responsive scaling** — font size and layout adapt to the visual's viewport

---

## Author

**Kirill Bezzubkin**
[LinkedIn](https://www.linkedin.com/in/kirill-bezzubkin-b6860b235) &middot; [GitHub](https://github.com/rusloc/oneCalendar)
