/*
*  Power BI Visual CLI
*
*  Copyright (c) Microsoft Corporation
*  All rights reserved.
*/
"use strict";

import powerbi from "powerbi-visuals-api";
import { FormattingSettingsService } from "powerbi-visuals-utils-formattingmodel";
import "./../style/visual.less";

import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisual = powerbi.extensibility.visual.IVisual;
import IVisualHost = powerbi.extensibility.visual.IVisualHost;
import ISelectionId = powerbi.visuals.ISelectionId;

import { VisualFormattingSettingsModel } from "./settings";

interface DateDataPoint {
    date: Date;
    selectionId: ISelectionId;
    // Pre-computed cached properties for fast filtering
    yearStr: string;
    quarterLabel: string;
    monthLabel: string;
    weekNum: number;
    weekLabel: string;
    isoDateStr: string;
}

export class Visual implements IVisual {
    private static readonly BLOCK_IDS = ['year', 'quarter', 'month', 'week'] as const;
    private static readonly MONTH_NAMES = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

    // SVG icon constants — stroke-width="3" for bold look
    // Main panel: "+" to expand, "−" to collapse
    private static readonly SVG_EXPAND = '<svg viewBox="0 0 24 24" width="16" height="16" stroke="currentColor" stroke-width="3" fill="none" stroke-linecap="round" stroke-linejoin="round"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>';
    private static readonly SVG_COLLAPSE = '<svg viewBox="0 0 24 24" width="16" height="16" stroke="currentColor" stroke-width="3" fill="none" stroke-linecap="round" stroke-linejoin="round"><line x1="5" y1="12" x2="19" y2="12"/></svg>';
    // Block toggle: chevron-down to expand all, chevron-up to collapse all
    private static readonly SVG_TOGGLE_EXPAND = '<svg viewBox="0 0 24 24" width="16" height="16" stroke="currentColor" stroke-width="3" fill="none" stroke-linecap="round" stroke-linejoin="round"><polyline points="6 9 12 15 18 9"/></svg>';
    private static readonly SVG_TOGGLE_COLLAPSE = '<svg viewBox="0 0 24 24" width="16" height="16" stroke="currentColor" stroke-width="3" fill="none" stroke-linecap="round" stroke-linejoin="round"><polyline points="18 15 12 9 6 15"/></svg>';
    // Reset: stays as-is, just bolder
    private static readonly SVG_RESET = '<svg viewBox="0 0 24 24" width="16" height="16" stroke="currentColor" stroke-width="3" fill="none" stroke-linecap="round" stroke-linejoin="round"><path d="M3 12a9 9 0 1 0 9-9 9.75 9.75 0 0 0-6.74 2.74L3 8"/><path d="M3 3v5h5"/></svg>';
    // Block header arrows
    private static readonly SVG_ARROW_DOWN = '<svg viewBox="0 0 24 24" width="14" height="14" stroke="currentColor" stroke-width="3" fill="none" stroke-linecap="round" stroke-linejoin="round"><polyline points="6 9 12 15 18 9"/></svg>';
    private static readonly SVG_ARROW_UP = '<svg viewBox="0 0 24 24" width="14" height="14" stroke="currentColor" stroke-width="3" fill="none" stroke-linecap="round" stroke-linejoin="round"><polyline points="18 15 12 9 6 15"/></svg>';
    // View switch: two overlapping rectangles
    private static readonly SVG_VIEW_SWITCH = '<svg viewBox="0 0 24 24" width="16" height="16" stroke="currentColor" stroke-width="3" fill="none" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="7" width="12" height="10" rx="2"/><rect x="9" y="3" width="12" height="10" rx="2"/></svg>';

    private target: HTMLElement;
    private host: IVisualHost;
    private formattingSettings: VisualFormattingSettingsModel;
    private formattingSettingsService: FormattingSettingsService;

    private dataViewParsed: boolean = false;
    private mainPanelExpanded: boolean = true;
    private blocksExpanded: { [key: string]: boolean } = { year: false, quarter: false, month: false, week: false };
    private viewMode: 'normal' | 'last' = 'normal';

    private dataPoints: DateDataPoint[] = [];
    private masterDataPoints: DateDataPoint[] = [];  // Full unfiltered dataset cache
    private currentCategory: powerbi.DataViewCategoryColumn | null = null;

    // Cached DOM nodes (per PBI perf best practice)
    private mainPanelEl: HTMLElement | null = null;
    private btnCollapseEl: HTMLButtonElement | null = null;
    private btnToggleBlocksEl: HTMLButtonElement | null = null;
    private btnResetEl: HTMLButtonElement | null = null;
    private btnViewSwitchEl: HTMLButtonElement | null = null;
    private filterStatusEl: HTMLElement | null = null;
    private containerEl: HTMLElement | null = null;

    private selections: { [key: string]: Set<string> } = {
        year: new Set<string>(),
        quarter: new Set<string>(),
        month: new Set<string>(),
        week: new Set<string>()
    };
    private lastSelections: { [key: string]: Set<string> } = {
        year: new Set<string>(),
        quarter: new Set<string>(),
        month: new Set<string>(),
        week: new Set<string>()
    };
    private lastSelectedIndex: { [key: string]: number | null } = {
        year: null, quarter: null, month: null, week: null
    };

    // Explicit manual date ranges
    private dateFrom: Date | null = null;
    private dateTo: Date | null = null;

    constructor(options: VisualConstructorOptions) {
        this.host = options.host;
        this.formattingSettingsService = new FormattingSettingsService();
        this.target = options.element;

        this.target.innerHTML = `
            <div class="calendar-container">
                <div class="calendar-header">
                    <button id="btn-collapse" class="btn-icon" title="Toggle Main Panel" disabled>${Visual.SVG_EXPAND}</button>
                    <button id="btn-toggle-blocks" class="btn-icon" title="Toggle All Blocks" disabled>${Visual.SVG_TOGGLE_EXPAND}</button>
                    <button id="btn-view-switch" class="btn-icon" title="Switch View" disabled>${Visual.SVG_VIEW_SWITCH}</button>
                    <button id="btn-reset" class="btn-icon" title="Reset Selections" disabled>${Visual.SVG_RESET}</button>
                    <div id="filter-status" class="filter-status">No filters</div>
                </div>
                <div id="main-panel" class="main-panel collapsed">
                    <div class="empty-state">Please add a Date field.</div>
                </div>
            </div>
        `;

        // Cache DOM nodes once (PBI perf tip: avoid querySelector in update/render)
        this.mainPanelEl = this.target.querySelector('#main-panel') as HTMLElement;
        this.btnCollapseEl = this.target.querySelector('#btn-collapse') as HTMLButtonElement;
        this.btnToggleBlocksEl = this.target.querySelector('#btn-toggle-blocks') as HTMLButtonElement;
        this.btnResetEl = this.target.querySelector('#btn-reset') as HTMLButtonElement;
        this.btnViewSwitchEl = this.target.querySelector('#btn-view-switch') as HTMLButtonElement;
        this.filterStatusEl = this.target.querySelector('#filter-status') as HTMLElement;
        this.containerEl = this.target.querySelector('.calendar-container') as HTMLElement;

        this.bindStaticEvents();
    }

    private bindStaticEvents() {
        this.btnCollapseEl?.addEventListener('click', () => {
            if (!this.dataViewParsed) return;
            this.mainPanelExpanded = !this.mainPanelExpanded;
            if (this.mainPanelExpanded) {
                this.mainPanelEl?.classList.remove('collapsed');
                if (this.btnCollapseEl) this.btnCollapseEl.innerHTML = Visual.SVG_COLLAPSE;
            } else {
                this.mainPanelEl?.classList.add('collapsed');
                if (this.btnCollapseEl) this.btnCollapseEl.innerHTML = Visual.SVG_EXPAND;
            }
        });

        this.btnResetEl?.addEventListener('click', () => {
            if (!this.dataViewParsed) return;
            this.clearAllSelections();
            this.host.applyJsonFilter(null, "general", "filter", powerbi.FilterAction.remove);
            this.refreshUI(false);
        });

        this.btnViewSwitchEl?.addEventListener('click', () => {
            if (!this.dataViewParsed) return;
            this.viewMode = this.viewMode === 'normal' ? 'last' : 'normal';
            this.clearAllSelections();
            this.host.applyJsonFilter(null, "general", "filter", powerbi.FilterAction.remove);
            this.refreshUI(false);
        });

        this.btnToggleBlocksEl?.addEventListener('click', () => {
            if (!this.dataViewParsed) return;
            const anyCollapsed = Object.values(this.blocksExpanded).some(state => state === false);
            this.blocksExpanded = { year: anyCollapsed, quarter: anyCollapsed, month: anyCollapsed, week: anyCollapsed };
            this.refreshUI(false);
        });
    }

    private clearAllSelections() {
        this.selections = { year: new Set<string>(), quarter: new Set<string>(), month: new Set<string>(), week: new Set<string>() };
        this.lastSelections = { year: new Set<string>(), quarter: new Set<string>(), month: new Set<string>(), week: new Set<string>() };
        this.lastSelectedIndex = { year: null, quarter: null, month: null, week: null };
        this.dateFrom = null;
        this.dateTo = null;
    }

    private refreshUI(applyFilter: boolean = false) {
        if (!this.dataPoints || this.dataPoints.length === 0) {
            this.clearPanel("No dates available.");
            return;
        }

        if (this.viewMode === 'last') {
            this.refreshUILast(applyFilter);
            return;
        }

        // Use cached properties for fast filtering — no extractDateInformation calls
        const dp = this.dataPoints;

        // Root years
        const yearSet = new Set<string>();
        for (let i = 0; i < dp.length; i++) yearSet.add(dp[i].yearStr);
        const availableYears = Array.from(yearSet).sort();

        // Clean invalid year selections
        for (const y of this.selections.year) {
            if (!yearSet.has(y)) this.selections.year.delete(y);
        }

        // Filter by Year
        const selYear = this.selections.year;
        const dpAfterYear = selYear.size === 0 ? dp : dp.filter(p => selYear.has(p.yearStr));

        // Extract quarters from year-filtered set
        const qMap = new Map<string, number>();
        for (let i = 0; i < dpAfterYear.length; i++) {
            const p = dpAfterYear[i];
            if (!qMap.has(p.quarterLabel)) qMap.set(p.quarterLabel, p.date.getTime());
        }
        const availableQuarters = Array.from(qMap.entries()).sort((a, b) => a[1] - b[1]).map(e => e[0]);
        const qSet = new Set(availableQuarters);

        for (const q of this.selections.quarter) {
            if (!qSet.has(q)) this.selections.quarter.delete(q);
        }

        // Filter by Quarter
        const selQ = this.selections.quarter;
        const dpAfterQuarter = selQ.size === 0 ? dpAfterYear : dpAfterYear.filter(p => selQ.has(p.quarterLabel));

        // Extract months
        const mMap = new Map<string, number>();
        for (let i = 0; i < dpAfterQuarter.length; i++) {
            const p = dpAfterQuarter[i];
            if (!mMap.has(p.monthLabel)) mMap.set(p.monthLabel, p.date.getTime());
        }
        const availableMonths = Array.from(mMap.entries()).sort((a, b) => a[1] - b[1]).map(e => e[0]);
        const mSet = new Set(availableMonths);

        for (const m of this.selections.month) {
            if (!mSet.has(m)) this.selections.month.delete(m);
        }

        // Filter by Month
        const selM = this.selections.month;
        const dpAfterMonth = selM.size === 0 ? dpAfterQuarter : dpAfterQuarter.filter(p => selM.has(p.monthLabel));

        // Extract weeks
        const wSet = new Set<number>();
        for (let i = 0; i < dpAfterMonth.length; i++) wSet.add(dpAfterMonth[i].weekNum);
        const availableWeeks = Array.from(wSet).sort((a, b) => a - b).map(w => `W${w}`);
        const wLabelSet = new Set(availableWeeks);

        for (const w of this.selections.week) {
            if (!wLabelSet.has(w)) this.selections.week.delete(w);
        }

        // Filter by Week
        const selW = this.selections.week;
        const dpAfterWeek = selW.size === 0 ? dpAfterMonth : dpAfterMonth.filter(p => selW.has(p.weekLabel));

        // Filter by date range
        const dFrom = this.dateFrom;
        const dTo = this.dateTo;
        const finalDataPoints = (!dFrom && !dTo) ? dpAfterWeek : dpAfterWeek.filter(p => {
            if (dFrom && p.date < dFrom) return false;
            if (dTo && p.date > dTo) return false;
            return true;
        });

        // Find min/max without spread (stack-safe for any size)
        let minTime = Infinity, maxTime = -Infinity;
        for (let i = 0; i < dpAfterWeek.length; i++) {
            const t = dpAfterWeek[i].date.getTime();
            if (t < minTime) minTime = t;
            if (t > maxTime) maxTime = t;
        }
        const unconstrainedMinDate = dpAfterWeek.length ? new Date(minTime) : null;
        const unconstrainedMaxDate = dpAfterWeek.length ? new Date(maxTime) : null;

        const displayInfo = {
            years: availableYears,
            quarters: availableQuarters,
            months: availableMonths,
            weeks: availableWeeks,
            minDate: unconstrainedMinDate,
            maxDate: unconstrainedMaxDate
        };

        this.renderDynamicContent(displayInfo);

        // Apply JSON filter
        if (applyFilter && this.currentCategory) {
            const hasSelections = this.selections.year.size > 0 ||
                this.selections.quarter.size > 0 ||
                this.selections.month.size > 0 ||
                this.selections.week.size > 0 ||
                this.dateFrom !== null ||
                this.dateTo !== null;
            this.applyJsonFilterFromDataPoints(hasSelections ? finalDataPoints : null);
        }
    }

    private refreshUILast(applyFilter: boolean) {
        const now = new Date();
        const currentYear = now.getFullYear();
        const currentMonth = now.getMonth();
        const currentQuarter = Math.floor(currentMonth / 3) + 1;

        // Build available items for each block
        const yearItems = ['1 LY', '2 LY', '3 LY', '4 LY', '5 LY'];
        const quarterItems = ['1 LQ', '2 LQ', '3 LQ', '4 LQ', '5 LQ', '6 LQ'];
        const monthItems = ['1 LM', '2 LM', '3 LM', '4 LM', '5 LM', '6 LM'];
        const weekItems: string[] = [];
        for (let i = 1; i <= 12; i++) weekItems.push(`${i} LW`);

        // Compute date range based on "last" selections
        let rangeStart: Date | null = null;
        let rangeEnd: Date | null = null;

        // Year: N LY = last N years, excluding current year
        if (this.lastSelections.year.size > 0) {
            let maxN = 0;
            this.lastSelections.year.forEach(v => {
                const n = parseInt(v);
                if (n > maxN) maxN = n;
            });
            rangeStart = new Date(currentYear - maxN, 0, 1);
            rangeEnd = new Date(currentYear - 1, 11, 31, 23, 59, 59, 999);
        }

        // Quarter: N LQ = last N quarters, excluding current quarter
        if (this.lastSelections.quarter.size > 0) {
            let maxN = 0;
            this.lastSelections.quarter.forEach(v => {
                const n = parseInt(v);
                if (n > maxN) maxN = n;
            });
            // End of last quarter (quarter before current)
            const endQ = currentQuarter - 1;
            const endYear = endQ <= 0 ? currentYear - 1 : currentYear;
            const endQAdj = endQ <= 0 ? endQ + 4 : endQ;
            const qEnd = new Date(endYear, endQAdj * 3, 0, 23, 59, 59, 999); // last day of end quarter

            // Start of (maxN quarters before current)
            let startQTotal = (currentYear * 4 + currentQuarter) - maxN;
            const startYear = Math.floor((startQTotal - 1) / 4);
            const startQ = ((startQTotal - 1) % 4) + 1;
            const qStart = new Date(startYear, (startQ - 1) * 3, 1);

            rangeStart = rangeStart ? (qStart > rangeStart ? qStart : rangeStart) : qStart;
            rangeEnd = rangeEnd ? (qEnd < rangeEnd ? qEnd : rangeEnd) : qEnd;
        }

        // Month: N LM = last N months, excluding current month
        if (this.lastSelections.month.size > 0) {
            let maxN = 0;
            this.lastSelections.month.forEach(v => {
                const n = parseInt(v);
                if (n > maxN) maxN = n;
            });
            // End of last month
            const mEnd = new Date(currentYear, currentMonth, 0, 23, 59, 59, 999);
            // Start: N months before current month
            const startDate = new Date(currentYear, currentMonth - maxN, 1);

            rangeStart = rangeStart ? (startDate > rangeStart ? startDate : rangeStart) : startDate;
            rangeEnd = rangeEnd ? (mEnd < rangeEnd ? mEnd : rangeEnd) : mEnd;
        }

        // Week: N LW = last N weeks, excluding current week
        if (this.lastSelections.week.size > 0) {
            let maxN = 0;
            this.lastSelections.week.forEach(v => {
                const n = parseInt(v);
                if (n > maxN) maxN = n;
            });
            // Find Monday of current ISO week
            const dayOfWeek = now.getDay() || 7; // 1=Mon..7=Sun
            const mondayThisWeek = new Date(now.getFullYear(), now.getMonth(), now.getDate() - dayOfWeek + 1);
            // End: Sunday before current week
            const wEnd = new Date(mondayThisWeek.getTime() - 1); // last ms of previous week
            wEnd.setHours(23, 59, 59, 999);
            // Start: Monday of (maxN weeks before current)
            const wStart = new Date(mondayThisWeek.getTime() - maxN * 7 * 86400000);

            rangeStart = rangeStart ? (wStart > rangeStart ? wStart : rangeStart) : wStart;
            rangeEnd = rangeEnd ? (wEnd < rangeEnd ? wEnd : rangeEnd) : wEnd;
        }

        // Filter data points
        let finalDataPoints: DateDataPoint[];
        if (rangeStart && rangeEnd) {
            finalDataPoints = this.dataPoints.filter(dp =>
                dp.date >= rangeStart! && dp.date <= rangeEnd!
            );
        } else {
            finalDataPoints = [];
        }

        // Compute display info for Last mode (just pass static items)
        const displayInfo = {
            years: yearItems,
            quarters: quarterItems,
            months: monthItems,
            weeks: weekItems,
            minDate: rangeStart,
            maxDate: rangeEnd
        };

        this.renderDynamicContent(displayInfo);

        // Apply JSON filter
        if (applyFilter && this.currentCategory) {
            const hasSelections = this.lastSelections.year.size > 0 ||
                this.lastSelections.quarter.size > 0 ||
                this.lastSelections.month.size > 0 ||
                this.lastSelections.week.size > 0;

            this.applyJsonFilterFromDataPoints(hasSelections ? finalDataPoints : null);
        }
    }

    /** Shared filter application — deduplicates logic between Normal & Last modes */
    private applyJsonFilterFromDataPoints(dataPoints: DateDataPoint[] | null) {
        if (!this.currentCategory) return;
        if (dataPoints && dataPoints.length > 0) {
            const queryName = this.currentCategory.source.queryName || "Calendar.Date";
            let tableName = queryName;
            if (queryName.includes('.')) tableName = queryName.substring(0, queryName.indexOf('.'));
            tableName = tableName.replace(/"/g, "");

            // Use pre-computed isoDateStr — no Date construction needed
            const seen = new Set<string>();
            const uniqueDates: string[] = [];
            for (let i = 0; i < dataPoints.length; i++) {
                const iso = dataPoints[i].isoDateStr;
                if (!seen.has(iso)) { seen.add(iso); uniqueDates.push(iso); }
            }

            this.host.applyJsonFilter({
                $schema: "http://powerbi.com/product/schema#basic",
                target: { table: tableName, column: this.currentCategory.source.displayName },
                operator: "In",
                values: uniqueDates,
                filterType: 1
            } as any, "general", "filter", powerbi.FilterAction.merge);
        } else {
            this.host.applyJsonFilter(null, "general", "filter", powerbi.FilterAction.remove);
        }
    }

    private getISOWeekYear(date: Date): number {
        const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
        const dayNum = d.getUTCDay() || 7;
        d.setUTCDate(d.getUTCDate() + 4 - dayNum);
        return d.getUTCFullYear();
    }

    private extractDateInformation(dates: Date[]) {
        const years = new Set<string>();
        const quarters = new Map<string, Date>();
        const months = new Map<string, Date>();
        const weeks = new Set<number>();

        dates.forEach(date => {
            if (!date || !(date instanceof Date)) return;
            const y = date.getFullYear();
            years.add(String(y));

            const m = date.getMonth();
            const q = Math.floor(m / 3) + 1;

            const qLabel = `Q${q} ${y}`;
            if (!quarters.has(qLabel)) quarters.set(qLabel, new Date(y, m, 1));

            const mLabel = `${date.toLocaleString('default', { month: 'short' })} ${y}`;
            if (!months.has(mLabel)) months.set(mLabel, new Date(y, m, 1));

            const weekNum = this.getISOWeekNumber(date);
            weeks.add(weekNum);
        });

        const sortedQuarters = Array.from(quarters.entries()).sort((a, b) => a[1].getTime() - b[1].getTime()).map(e => e[0]);
        const sortedMonths = Array.from(months.entries()).sort((a, b) => a[1].getTime() - b[1].getTime()).map(e => e[0]);

        return {
            years: Array.from(years).sort(),
            quarters: sortedQuarters,
            months: sortedMonths,
            weeks: Array.from(weeks).sort((a, b) => a - b).map(w => `W${w}`)
        };
    }

    private getISOWeekNumber(date: Date): number {
        const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
        const dayNum = d.getUTCDay() || 7;
        d.setUTCDate(d.getUTCDate() + 4 - dayNum);
        const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
        return Math.ceil((((d.getTime() - yearStart.getTime()) / 86400000) + 1) / 7);
    }

    private formatDateForInput(date: Date | null): string {
        if (!date) return "";
        const y = date.getFullYear();
        const m = String(date.getMonth() + 1).padStart(2, '0');
        const d = String(date.getDate()).padStart(2, '0');
        return `${y}-${m}-${d}`;
    }

    private renderDynamicContent(dateInfo: any) {
        const mainPanel = this.mainPanelEl;
        const btnCollapse = this.btnCollapseEl;
        const btnToggleBlocks = this.btnToggleBlocksEl;
        const btnReset = this.btnResetEl;
        const btnViewSwitch = this.btnViewSwitchEl;

        if (btnCollapse) btnCollapse.disabled = false;
        if (btnReset) btnReset.disabled = false;
        if (btnViewSwitch) btnViewSwitch.disabled = false;

        const allExpanded = Object.values(this.blocksExpanded).every(s => s);

        if (btnToggleBlocks) {
            btnToggleBlocks.disabled = false;
            const currentExpandedAttr = btnToggleBlocks.getAttribute('data-expanded');
            if (currentExpandedAttr !== String(allExpanded)) {
                btnToggleBlocks.innerHTML = allExpanded ? Visual.SVG_TOGGLE_COLLAPSE : Visual.SVG_TOGGLE_EXPAND;
                btnToggleBlocks.setAttribute('data-expanded', String(allExpanded));
            }
        }

        // Apply custom accent color from settings to root container via CSS variable
        const accentColor = this.formattingSettings?.visualSettingsCard?.accentColor?.value?.value || "#3b82f6";
        const btnPosition = this.formattingSettings?.visualSettingsCard?.buttonPosition?.value?.value || "left";
        const borderWeight = this.formattingSettings?.visualSettingsCard?.containerBorderWeight?.value ?? 1;
        const borderColor = this.formattingSettings?.visualSettingsCard?.containerBorderColor?.value?.value || "#e5e7eb";
        const datesBgColor = this.formattingSettings?.visualSettingsCard?.datesBgColor?.value?.value || "#fafafa";

        const container = this.containerEl;
        if (container) {
            container.style.setProperty('--accent-color', accentColor);
            container.style.setProperty('--container-border-weight', `${borderWeight}px`);
            container.style.setProperty('--container-border-color', borderColor);
            container.style.setProperty('--dates-bg-color', datesBgColor);

            if (btnPosition === "right") {
                container.classList.add('btns-right');
            } else {
                container.classList.remove('btns-right');
            }
        }

        // Capture current scroll positions to avoid resetting to left
        const scrollPositions: { [key: string]: number } = {};
        if (this.dataViewParsed && mainPanel) {
            Visual.BLOCK_IDS.forEach(blockId => {
                const contentNode = mainPanel.querySelector(`#content-${blockId}`);
                if (contentNode) scrollPositions[blockId] = contentNode.scrollLeft;
            });
        }

        if (!this.dataViewParsed) {
            this.mainPanelExpanded = true;
            if (mainPanel) mainPanel.classList.remove('collapsed');
            if (btnCollapse) btnCollapse.innerHTML = Visual.SVG_COLLAPSE;
            this.dataViewParsed = true;
        }

        const currentSelections = this.viewMode === 'last' ? this.lastSelections : this.selections;

        const buildBlock = (id: string, label: string, items: string[]) => {
            const isExpanded = this.blocksExpanded[id];

            let itemHtml = '';
            items.forEach((item: string, idx: number) => {
                const isSelected = currentSelections[id].has(item);
                const selClass = isSelected ? 'selected' : '';
                itemHtml += `<div class="block-item ${selClass}" data-id="${id}" data-index="${idx}" data-val="${item}">${item}</div>`;
            });

            return `
                <div class="calendar-block">
                    <div class="block-header" id="header-${id}">
                        <span class="block-title">${label}</span>
                        <span class="icon" id="icon-${id}">${isExpanded ? Visual.SVG_ARROW_UP : Visual.SVG_ARROW_DOWN}</span>
                    </div>
                    <div class="block-content ${isExpanded ? '' : 'collapsed'}" id="content-${id}">
                        ${itemHtml}
                    </div>
                </div>
            `;
        };

        const isLast = this.viewMode === 'last';
        let html = '';
        html += buildBlock('year', isLast ? 'Year (last)' : 'Year', dateInfo.years);
        html += buildBlock('quarter', isLast ? 'Quarter (last)' : 'Quarter', dateInfo.quarters);
        html += buildBlock('month', isLast ? 'Month (last)' : 'Month', dateInfo.months);
        html += buildBlock('week', isLast ? 'Week (last)' : 'Week', dateInfo.weeks);

        // Apply visual constraints but populate with active specific selections if made
        const inputMinBound = this.formatDateForInput(dateInfo.minDate);
        const inputMaxBound = this.formatDateForInput(dateInfo.maxDate);

        const currentFromValue = this.dateFrom ? this.formatDateForInput(this.dateFrom) : inputMinBound;
        const currentToValue = this.dateTo ? this.formatDateForInput(this.dateTo) : inputMaxBound;

        // Date inputs: in Last mode, disable and show today's date
        const todayStr = this.formatDateForInput(new Date());
        if (isLast) {
            html += `
                <div class="dates-block dates-disabled">
                    <div class="dates-inputs">
                        <input id="date-from" type="date" value="${todayStr}" disabled />
                        <span class="separator">–</span>
                        <input id="date-to" type="date" value="${todayStr}" disabled />
                    </div>
                </div>
            `;
        } else {
            html += `
                <div class="dates-block">
                    <div class="dates-inputs">
                        <input id="date-from" type="date" value="${currentFromValue}" min="${inputMinBound}" max="${inputMaxBound}" />
                        <span class="separator">–</span>
                        <input id="date-to" type="date" value="${currentToValue}" min="${inputMinBound}" max="${inputMaxBound}" />
                    </div>
                </div>
            `;
        }

        if (mainPanel) {
            mainPanel.innerHTML = html;

            // Update filter status indicator
            const filterStatus = this.filterStatusEl;
            if (filterStatus) {
                const cs = this.viewMode === 'last' ? this.lastSelections : this.selections;
                const activeFilters: string[] = [];
                if (cs.year.size > 0) activeFilters.push('YR');
                if (cs.quarter.size > 0) activeFilters.push('QT');
                if (cs.month.size > 0) activeFilters.push('MT');
                if (cs.week.size > 0) activeFilters.push('WK');
                if (this.dateFrom || this.dateTo) activeFilters.push('DT');
                const prefix = this.viewMode === 'last' ? 'Last: ' : 'Filters: ';
                filterStatus.textContent = activeFilters.length > 0 ? prefix + activeFilters.join(', ') : 'No date filters applied';
                filterStatus.classList.toggle('active', activeFilters.length > 0);
            }

            // Restore scroll positions
            Visual.BLOCK_IDS.forEach(blockId => {
                if (scrollPositions[blockId]) {
                    const contentNode = mainPanel.querySelector(`#content-${blockId}`);
                    if (contentNode) contentNode.scrollLeft = scrollPositions[blockId];
                }
            });
        }

        // Bind dynamic event listeners
        Visual.BLOCK_IDS.forEach(blockId => {
            const header = this.target.querySelector(`#header-${blockId}`);
            const content = this.target.querySelector(`#content-${blockId}`);
            const icon = this.target.querySelector(`#icon-${blockId}`);

            header?.addEventListener('click', () => {
                this.blocksExpanded[blockId] = !this.blocksExpanded[blockId];
                if (this.blocksExpanded[blockId]) {
                    content?.classList.remove('collapsed');
                    if (icon) icon.innerHTML = Visual.SVG_ARROW_UP;
                } else {
                    content?.classList.add('collapsed');
                    if (icon) icon.innerHTML = Visual.SVG_ARROW_DOWN;
                }
            });
        });

        // Date input listeners
        const dFrom = this.target.querySelector('#date-from') as HTMLInputElement;
        const dTo = this.target.querySelector('#date-to') as HTMLInputElement;

        dFrom?.addEventListener('change', () => {
            this.dateFrom = dFrom.value ? new Date(dFrom.value + 'T00:00:00') : null;
            this.refreshUI(true);
        });

        dTo?.addEventListener('change', () => {
            this.dateTo = dTo.value ? new Date(dTo.value + 'T00:00:00') : null;
            this.refreshUI(true);
        });

        const blockItems = this.target.querySelectorAll('.block-item');
        const selSet = this.viewMode === 'last' ? this.lastSelections : this.selections;
        blockItems.forEach(item => {
            item.addEventListener('click', (e: Event) => {
                const mouseEvent = e as MouseEvent;
                const target = mouseEvent.target as HTMLElement;
                const blockId = target.getAttribute('data-id');
                const indexStr = target.getAttribute('data-index');
                const valStr = target.getAttribute('data-val');

                if (!blockId || !indexStr || !valStr) return;
                const idx = parseInt(indexStr);

                if (this.viewMode === 'last') {
                    // Single-select in Last mode: toggle off or replace
                    if (selSet[blockId].has(valStr)) {
                        selSet[blockId].delete(valStr);
                    } else {
                        selSet[blockId].clear();
                        selSet[blockId].add(valStr);
                    }
                    this.lastSelectedIndex[blockId] = idx;
                } else if (mouseEvent.shiftKey && this.lastSelectedIndex[blockId] !== null) {
                    const start = Math.min(this.lastSelectedIndex[blockId]!, idx);
                    const end = Math.max(this.lastSelectedIndex[blockId]!, idx);

                    const currentGroupItems = Array.from(this.target.querySelectorAll(`.block-item[data-id="${blockId}"]`));
                    currentGroupItems.forEach((el) => {
                        const elIdx = parseInt(el.getAttribute('data-index')!);
                        const elVal = el.getAttribute('data-val')!;
                        if (elIdx >= start && elIdx <= end) {
                            selSet[blockId].add(elVal);
                        }
                    });
                } else {
                    if (selSet[blockId].has(valStr)) {
                        selSet[blockId].delete(valStr);
                        this.lastSelectedIndex[blockId] = null;
                    } else {
                        selSet[blockId].add(valStr);
                        this.lastSelectedIndex[blockId] = idx;
                    }
                }

                this.refreshUI(true);
            });
        });
    }

    public update(options: VisualUpdateOptions) {
        this.formattingSettings = this.formattingSettingsService.populateFormattingSettingsModel(VisualFormattingSettingsModel, options.dataViews[0]);

        // Dynamically scale UI base font-size to the viewport width
        // Assume 350px width is our "standard" 14px font size layout scaling
        if (options.viewport) {
            const scale = options.viewport.width / 350;
            const baseFontSize = Math.max(6, 14 * scale); // clamp to 6px min
            const container = this.target.querySelector('.calendar-container') as HTMLElement;
            if (container) {
                container.style.setProperty('--base-font-size', `${baseFontSize}px`);
            }
        }

        try {
            const dataViews = options.dataViews;
            if (dataViews && dataViews[0] && dataViews[0].categorical && dataViews[0].categorical.categories) {
                const category = dataViews[0].categorical.categories[0];
                this.currentCategory = category;
                const rawValues = category.values;

                if (rawValues && rawValues.length > 0) {
                    // Parse incoming data
                    const incomingDataPoints: DateDataPoint[] = [];
                    rawValues.forEach((v: any, index: number) => {
                        let d: Date | null = null;
                        if (v instanceof Date) {
                            d = v;
                        } else if (typeof v === 'string' || typeof v === 'number') {
                            const pd = new Date(v);
                            if (!isNaN(pd.getTime())) d = pd;
                        }

                        if (d) {
                            const selectionId = this.host.createSelectionIdBuilder()
                                .withCategory(category, index)
                                .createSelectionId();

                            // Pre-compute cached properties
                            const y = d.getFullYear();
                            const m = d.getMonth();
                            const q = Math.floor(m / 3) + 1;
                            const wn = this.getISOWeekNumber(d);
                            const normalizedDate = new Date(y, m, d.getDate());

                            incomingDataPoints.push({
                                date: d,
                                selectionId: selectionId,
                                yearStr: String(y),
                                quarterLabel: `Q${q} ${y}`,
                                monthLabel: `${Visual.MONTH_NAMES[m]} ${y}`,
                                weekNum: wn,
                                weekLabel: `W${wn}`,
                                isoDateStr: normalizedDate.toISOString()
                            });
                        }
                    });

                    // Only update master cache if incoming data is >= cached (full dataset)
                    // If incoming is smaller, another visual is cross-filtering us — ignore it
                    if (incomingDataPoints.length >= this.masterDataPoints.length || this.masterDataPoints.length === 0) {
                        this.masterDataPoints = incomingDataPoints;
                    }

                    // Always render from the master (full) cache
                    this.dataPoints = this.masterDataPoints;

                    if (this.dataPoints.length > 0) {
                        this.refreshUI(false); // Parse UI visually but don't re-apply filters back
                    } else {
                        this.clearPanel("Supplied data could not be parsed as dates.");
                    }
                } else {
                    this.clearPanel("No data supplied.");
                }
            } else {
                this.clearPanel("Please add a Date field.");
            }
        } catch (e) {
            console.error("Error updating visual", e);
            this.clearPanel("An error occurred extracting data.");
        }
    }

    private clearPanel(msg: string) {
        this.dataViewParsed = false;
        if (this.btnCollapseEl) {
            this.btnCollapseEl.disabled = true;
            this.btnCollapseEl.innerHTML = Visual.SVG_EXPAND;
        }
        if (this.btnToggleBlocksEl) {
            this.btnToggleBlocksEl.disabled = true;
            this.btnToggleBlocksEl.innerHTML = Visual.SVG_TOGGLE_EXPAND;
            this.btnToggleBlocksEl.removeAttribute('data-expanded');
        }
        if (this.btnResetEl) this.btnResetEl.disabled = true;
        if (this.btnViewSwitchEl) this.btnViewSwitchEl.disabled = true;
        this.mainPanelExpanded = true;
        if (this.mainPanelEl) {
            this.mainPanelEl.classList.remove('collapsed');
            this.mainPanelEl.innerHTML = `<div class="empty-state">${msg}</div>`;
        }
    }

    public getFormattingModel(): powerbi.visuals.FormattingModel {
        return this.formattingSettingsService.buildFormattingModel(this.formattingSettings);
    }
}