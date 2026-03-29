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
    dateInt: number; // YYYYMMDD for timezone-safe comparison
}

interface DisplayInfo {
    years: string[];
    quarters: string[];
    months: string[];
    weeks: string[];
    minDate: Date | null;
    maxDate: Date | null;
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

    // DOM-diffing state for fast-path rendering (P4)
    private renderedBlockItems: { [key: string]: string[] } = {};
    private renderedViewMode: string = '';
    private renderedDateBounds: string = '';
    private delegatedListenersBound: boolean = false;

    // Bookmark: track last filter applied by the visual itself, so we can distinguish
    // a PBI echo-back from a genuine bookmark restore in update()
    private lastAppliedFilterKey: string = '';

    // Landing page state
    private isLandingPageOn: boolean = false;
    private landingPageEl: HTMLElement | null = null;

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
            // P4: toggle CSS classes directly instead of full refreshUI
            if (this.mainPanelEl) {
                Visual.BLOCK_IDS.forEach(blockId => {
                    const content = this.mainPanelEl!.querySelector(`#content-${blockId}`);
                    const icon = this.mainPanelEl!.querySelector(`#icon-${blockId}`);
                    if (anyCollapsed) {
                        content?.classList.remove('collapsed');
                        if (icon) icon.innerHTML = Visual.SVG_ARROW_UP;
                    } else {
                        content?.classList.add('collapsed');
                        if (icon) icon.innerHTML = Visual.SVG_ARROW_DOWN;
                    }
                });
            }
            if (this.btnToggleBlocksEl) {
                this.btnToggleBlocksEl.innerHTML = anyCollapsed ? Visual.SVG_TOGGLE_COLLAPSE : Visual.SVG_TOGGLE_EXPAND;
                this.btnToggleBlocksEl.setAttribute('data-expanded', String(anyCollapsed));
            }
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
        // Anchor to dataset's most recent date, not wall-clock
        let maxDateInt = 0;
        for (let i = 0; i < this.dataPoints.length; i++) {
            if (this.dataPoints[i].dateInt > maxDateInt) maxDateInt = this.dataPoints[i].dateInt;
        }
        const anchorYear = Math.floor(maxDateInt / 10000);
        const anchorMonth = Math.floor((maxDateInt % 10000) / 100) - 1;
        const anchorDay = maxDateInt % 100;
        const now = new Date(anchorYear, anchorMonth, anchorDay);
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
            const startQTotal = (currentYear * 4 + currentQuarter) - maxN;
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

        // Week: N LW = last N complete weeks (Mon-Sun), excluding current week
        if (this.lastSelections.week.size > 0) {
            let maxN = 0;
            this.lastSelections.week.forEach(v => {
                const n = parseInt(v);
                if (n > maxN) maxN = n;
            });
            // Find Monday of current week (getDay: 0=Sun, 1=Mon..6=Sat)
            const dow = now.getDay() || 7; // convert Sun=0 to 7 for ISO
            const mondayThisWeek = new Date(now.getFullYear(), now.getMonth(), now.getDate() - dow + 1);
            // End: Sunday before this Monday = yesterday for Monday, last Sunday otherwise
            const wEnd = new Date(mondayThisWeek.getFullYear(), mondayThisWeek.getMonth(), mondayThisWeek.getDate() - 1, 23, 59, 59, 999);
            // Start: Monday of N weeks ago
            const wStart = new Date(mondayThisWeek.getFullYear(), mondayThisWeek.getMonth(), mondayThisWeek.getDate() - maxN * 7);

            rangeStart = rangeStart ? (wStart > rangeStart ? wStart : rangeStart) : wStart;
            rangeEnd = rangeEnd ? (wEnd < rangeEnd ? wEnd : rangeEnd) : wEnd;
        }

        // Filter data points using integer YYYYMMDD comparison (timezone-immune)
        const toDateInt = (d: Date) => d.getFullYear() * 10000 + (d.getMonth() + 1) * 100 + d.getDate();
        let finalDataPoints: DateDataPoint[];
        if (rangeStart && rangeEnd) {
            const startInt = toDateInt(rangeStart);
            const endInt = toDateInt(rangeEnd);
            finalDataPoints = this.dataPoints.filter(dp =>
                dp.dateInt >= startInt && dp.dateInt <= endInt
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

    /** Shared filter application — deduplicates logic between Normal & Last modes.
     *  Uses hybrid strategy: AdvancedFilter (min-max range, ~200 bytes) when selected dates
     *  are contiguous; falls back to BasicFilter "In" (full date list) when gaps exist. */
    private applyJsonFilterFromDataPoints(dataPoints: DateDataPoint[] | null) {
        if (!this.currentCategory) return;
        if (dataPoints && dataPoints.length > 0) {
            const queryName = this.currentCategory.source.queryName || "Calendar.Date";
            let tableName = queryName;
            if (queryName.includes('.')) tableName = queryName.substring(0, queryName.indexOf('.'));
            tableName = tableName.replace(/"/g, "");

            // Deduplicate and find min/max dateInt in a single pass
            const seen = new Set<string>();
            const uniqueDates: string[] = [];
            let minInt = Infinity, maxInt = -Infinity;
            for (let i = 0; i < dataPoints.length; i++) {
                const dp = dataPoints[i];
                if (!seen.has(dp.isoDateStr)) {
                    seen.add(dp.isoDateStr);
                    uniqueDates.push(dp.isoDateStr);
                }
                if (dp.dateInt < minInt) minInt = dp.dateInt;
                if (dp.dateInt > maxInt) maxInt = dp.dateInt;
            }

            // Contiguity check: count master data points within the same min-max range.
            // If masterCountInRange === uniqueDates.length, no gaps exist → use range filter.
            let masterCountInRange = 0;
            const masterSeen = new Set<string>();
            for (let i = 0; i < this.masterDataPoints.length; i++) {
                const mdp = this.masterDataPoints[i];
                if (mdp.dateInt >= minInt && mdp.dateInt <= maxInt && !masterSeen.has(mdp.isoDateStr)) {
                    masterSeen.add(mdp.isoDateStr);
                    masterCountInRange++;
                }
            }

            const isContiguous = masterCountInRange === uniqueDates.length;
            const target = { table: tableName, column: this.currentCategory.source.displayName };

            // Track filter fingerprint for bookmark echo-back detection
            this.lastAppliedFilterKey = isContiguous
                ? `range|${minInt}|${maxInt}`
                : `in|${uniqueDates.length}|${uniqueDates[0]}|${uniqueDates[uniqueDates.length - 1]}`;

            if (isContiguous) {
                // AdvancedFilter: 2 values (~200 bytes) instead of N dates (~600KB)
                const minIso = uniqueDates.reduce((a, b) => a < b ? a : b);
                const maxIso = uniqueDates.reduce((a, b) => a > b ? a : b);

                console.log(`[oneCalendar] Range filter (${this.viewMode}): ${minIso} → ${maxIso} (${uniqueDates.length} dates, contiguous)`);

                this.host.applyJsonFilter({
                    $schema: "http://powerbi.com/product/schema#advanced",
                    target: target,
                    logicalOperator: "And",
                    conditions: [
                        { operator: "GreaterThanOrEqual", value: minIso },
                        { operator: "LessThanOrEqual", value: maxIso }
                    ],
                    filterType: 0
                } as any, "general", "filter", powerbi.FilterAction.merge);
            } else {
                // BasicFilter "In": non-contiguous selection, must enumerate all dates
                console.log(`[oneCalendar] In filter (${this.viewMode}): ${uniqueDates.length} dates (non-contiguous)`);

                this.host.applyJsonFilter({
                    $schema: "http://powerbi.com/product/schema#basic",
                    target: target,
                    operator: "In",
                    values: uniqueDates,
                    filterType: 1
                } as any, "general", "filter", powerbi.FilterAction.merge);
            }
        } else {
            this.lastAppliedFilterKey = '';
            console.log(`[oneCalendar] Filter removed (${this.viewMode} mode)`);
            this.host.applyJsonFilter(null, "general", "filter", powerbi.FilterAction.remove);
        }
    }

    // getISOWeekYear and extractDateInformation removed — dead code after optimization

    /** Bookmark support: restore internal selection state from saved jsonFilters.
     *  Determines the lowest selection level that fully explains the bookmarked dates,
     *  respecting the cascading filter logic (year > quarter > month > week).
     *  Upper levels auto-select only to scope; lower levels stay unselected (filtered, not selected). */
    private restoreBookmarkFilter(options: VisualUpdateOptions): boolean {
        const jsonFilters = options.jsonFilters;
        if (!jsonFilters || jsonFilters.length === 0) {
            return false;
        }

        const filter = jsonFilters[0] as any;
        if (!filter || filter.filterType !== 1 || !filter.values || filter.values.length === 0) {
            return false;
        }

        // Build a key from incoming filter to compare against last self-applied filter
        const incomingKey = filter.values.slice().sort().join('|');
        if (incomingKey === this.lastAppliedFilterKey) {
            // This is our own filter echoed back by PBI — not a bookmark restore
            return false;
        }

        // Build set of bookmarked date integers (YYYYMMDD)
        const bookmarkedDates = new Set<number>();
        for (const val of filter.values) {
            const d = new Date(val);
            if (!isNaN(d.getTime())) {
                bookmarkedDates.add(d.getFullYear() * 10000 + (d.getMonth() + 1) * 100 + d.getDate());
            }
        }
        if (bookmarkedDates.size === 0) return false;

        // Collect which dataPoints match the bookmark
        const matchedDPs: DateDataPoint[] = [];
        for (const dp of this.dataPoints) {
            if (bookmarkedDates.has(dp.dateInt)) {
                matchedDPs.push(dp);
            }
        }
        if (matchedDPs.length === 0) return false;

        // Determine the lowest level that fully explains the bookmarked dates.
        // For each level, check if selecting the unique values at that level would
        // produce exactly the bookmarked date set (given the full dataset).
        this.clearAllSelections();

        const matchedYears = new Set(matchedDPs.map(dp => dp.yearStr));
        const matchedQuarters = new Set(matchedDPs.map(dp => dp.quarterLabel));
        const matchedMonths = new Set(matchedDPs.map(dp => dp.monthLabel));
        const matchedWeeks = new Set(matchedDPs.map(dp => dp.weekLabel));

        // Helper: count all dataPoints matching a set at a given level
        const countByLevel = (level: 'yearStr' | 'quarterLabel' | 'monthLabel' | 'weekLabel', vals: Set<string>) => {
            let count = 0;
            for (const dp of this.dataPoints) { if (vals.has(dp[level])) count++; }
            return count;
        };

        const targetCount = matchedDPs.length;

        // Try year-only: if selecting these years gives exactly the bookmarked dates
        if (countByLevel('yearStr', matchedYears) === targetCount) {
            this.selections.year = matchedYears;
        }
        // Try quarter (with parent years auto-selected for scoping)
        else if (countByLevel('quarterLabel', matchedQuarters) === targetCount) {
            this.selections.year = matchedYears;
            this.selections.quarter = matchedQuarters;
        }
        // Try month
        else if (countByLevel('monthLabel', matchedMonths) === targetCount) {
            this.selections.year = matchedYears;
            this.selections.quarter = matchedQuarters;
            this.selections.month = matchedMonths;
        }
        // Fall back to week (most granular)
        else {
            this.selections.year = matchedYears;
            this.selections.quarter = matchedQuarters;
            this.selections.month = matchedMonths;
            this.selections.week = matchedWeeks;
        }

        console.log(`[oneCalendar] Bookmark restored: ${bookmarkedDates.size} dates, selections:`,
            { years: [...this.selections.year], quarters: [...this.selections.quarter],
              months: [...this.selections.month], weeks: [...this.selections.week] });

        return true;
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

    /** Update the filter status badge in the header */
    private updateFilterStatus(cs: { [key: string]: Set<string> }) {
        const filterStatus = this.filterStatusEl;
        if (!filterStatus) return;
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

    /** Bind delegated event listeners on mainPanel once (P4) */
    private bindDelegatedListeners() {
        const panel = this.mainPanelEl;
        if (!panel) return;

        // Click delegation: block headers + block items
        panel.addEventListener('click', (e: MouseEvent) => {
            const target = e.target as HTMLElement;

            // --- Block header click (expand/collapse) ---
            const header = target.closest('.block-header') as HTMLElement | null;
            if (header) {
                const blockId = header.id.replace('header-', '');
                this.blocksExpanded[blockId] = !this.blocksExpanded[blockId];
                const content = panel.querySelector(`#content-${blockId}`);
                const icon = panel.querySelector(`#icon-${blockId}`);
                if (this.blocksExpanded[blockId]) {
                    content?.classList.remove('collapsed');
                    if (icon) icon.innerHTML = Visual.SVG_ARROW_UP;
                } else {
                    content?.classList.add('collapsed');
                    if (icon) icon.innerHTML = Visual.SVG_ARROW_DOWN;
                }
                const allExpanded = Object.values(this.blocksExpanded).every(s => s);
                if (this.btnToggleBlocksEl) {
                    this.btnToggleBlocksEl.innerHTML = allExpanded ? Visual.SVG_TOGGLE_COLLAPSE : Visual.SVG_TOGGLE_EXPAND;
                    this.btnToggleBlocksEl.setAttribute('data-expanded', String(allExpanded));
                }
                return;
            }

            // --- Block item click (selection) ---
            const item = target.closest('.block-item') as HTMLElement | null;
            if (!item) return;
            const blockId = item.getAttribute('data-id');
            const indexStr = item.getAttribute('data-index');
            const valStr = item.getAttribute('data-val');
            if (!blockId || !indexStr || !valStr) return;
            const idx = parseInt(indexStr);
            const selSet = this.viewMode === 'last' ? this.lastSelections : this.selections;

            if (this.viewMode === 'last') {
                const wasSelected = selSet[blockId].has(valStr);
                for (const bid of Visual.BLOCK_IDS) selSet[bid].clear();
                if (!wasSelected) selSet[blockId].add(valStr);
                this.lastSelectedIndex[blockId] = idx;
            } else if (e.shiftKey && this.lastSelectedIndex[blockId] !== null) {
                const start = Math.min(this.lastSelectedIndex[blockId]!, idx);
                const end = Math.max(this.lastSelectedIndex[blockId]!, idx);
                const items = this.renderedBlockItems[blockId] || [];
                for (let i = start; i <= end && i < items.length; i++) {
                    selSet[blockId].add(items[i]);
                }
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

        // Change delegation: date inputs
        panel.addEventListener('change', (e: Event) => {
            const target = e.target as HTMLInputElement;
            if (target.id === 'date-from') {
                this.dateFrom = target.value ? new Date(target.value + 'T00:00:00') : null;
                this.refreshUI(true);
            } else if (target.id === 'date-to') {
                this.dateTo = target.value ? new Date(target.value + 'T00:00:00') : null;
                this.refreshUI(true);
            }
        });
    }

    private renderDynamicContent(dateInfo: DisplayInfo) {
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

        if (!this.dataViewParsed) {
            this.mainPanelExpanded = true;
            if (mainPanel) mainPanel.classList.remove('collapsed');
            if (btnCollapse) btnCollapse.innerHTML = Visual.SVG_COLLAPSE;
            this.dataViewParsed = true;
        }

        const currentSelections = this.viewMode === 'last' ? this.lastSelections : this.selections;
        const isLast = this.viewMode === 'last';
        const minBound = this.formatDateForInput(dateInfo.minDate);
        const maxBound = this.formatDateForInput(dateInfo.maxDate);

        // --- P4: Check if structural rebuild is needed (DOM diffing) ---
        const dateBoundsKey = `${minBound}|${maxBound}`;
        const itemsUnchanged = this.renderedViewMode === this.viewMode &&
            this.renderedDateBounds === dateBoundsKey &&
            Visual.BLOCK_IDS.every(id => {
                const key = id === 'year' ? 'years' : id === 'quarter' ? 'quarters' : id === 'month' ? 'months' : 'weeks';
                const rendered = this.renderedBlockItems[id];
                const incoming = dateInfo[key];
                return rendered && rendered.length === incoming.length && rendered.every((v, i) => v === incoming[i]);
            });

        if (itemsUnchanged && mainPanel) {
            // ► Fast path: only toggle selected classes + update status (no innerHTML, no reflow)
            mainPanel.querySelectorAll('.block-item').forEach(el => {
                const id = el.getAttribute('data-id')!;
                const val = el.getAttribute('data-val')!;
                el.classList.toggle('selected', currentSelections[id].has(val));
            });
            if (!isLast) {
                const dFromEl = mainPanel.querySelector('#date-from') as HTMLInputElement;
                const dToEl = mainPanel.querySelector('#date-to') as HTMLInputElement;
                if (dFromEl) dFromEl.value = this.dateFrom ? this.formatDateForInput(this.dateFrom) : minBound;
                if (dToEl) dToEl.value = this.dateTo ? this.formatDateForInput(this.dateTo) : maxBound;
            }
            this.updateFilterStatus(currentSelections);
            return;
        }

        // ► Slow path: structural rebuild required
        // Capture current scroll positions to restore after innerHTML
        const scrollPositions: { [key: string]: number } = {};
        if (mainPanel) {
            Visual.BLOCK_IDS.forEach(blockId => {
                const contentNode = mainPanel.querySelector(`#content-${blockId}`);
                if (contentNode) scrollPositions[blockId] = contentNode.scrollLeft;
            });
        }

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

        let html = '';
        html += buildBlock('year', isLast ? 'Year (last)' : 'Year', dateInfo.years);
        html += buildBlock('quarter', isLast ? 'Quarter (last)' : 'Quarter', dateInfo.quarters);
        html += buildBlock('month', isLast ? 'Month (last)' : 'Month', dateInfo.months);
        html += buildBlock('week', isLast ? 'Week (last)' : 'Week', dateInfo.weeks);

        // Date inputs block
        if (isLast) {
            const fromVal = minBound || this.formatDateForInput(new Date());
            const toVal = maxBound || this.formatDateForInput(new Date());
            html += `
                <div class="dates-block dates-disabled">
                    <div class="dates-inputs">
                        <input id="date-from" type="date" value="${fromVal}" disabled />
                        <span class="separator">–</span>
                        <input id="date-to" type="date" value="${toVal}" disabled />
                    </div>
                </div>
            `;
        } else {
            const fromVal = this.dateFrom ? this.formatDateForInput(this.dateFrom) : minBound;
            const toVal = this.dateTo ? this.formatDateForInput(this.dateTo) : maxBound;
            html += `
                <div class="dates-block">
                    <div class="dates-inputs">
                        <input id="date-from" type="date" value="${fromVal}" min="${minBound}" max="${maxBound}" />
                        <span class="separator">–</span>
                        <input id="date-to" type="date" value="${toVal}" min="${minBound}" max="${maxBound}" />
                    </div>
                </div>
            `;
        }

        if (mainPanel) {
            mainPanel.innerHTML = html;

            // Restore scroll positions
            Visual.BLOCK_IDS.forEach(blockId => {
                if (scrollPositions[blockId] !== undefined) {
                    const contentNode = mainPanel.querySelector(`#content-${blockId}`);
                    if (contentNode) contentNode.scrollLeft = scrollPositions[blockId];
                }
            });

            // Bind delegated event listeners once (P4)
            if (!this.delegatedListenersBound) {
                this.bindDelegatedListeners();
                this.delegatedListenersBound = true;
            }
        }

        // Cache rendered state for fast-path diffing
        this.renderedBlockItems = {
            year: dateInfo.years.slice(),
            quarter: dateInfo.quarters.slice(),
            month: dateInfo.months.slice(),
            week: dateInfo.weeks.slice()
        };
        this.renderedViewMode = this.viewMode;
        this.renderedDateBounds = dateBoundsKey;

        this.updateFilterStatus(currentSelections);
    }

    public update(options: VisualUpdateOptions) {
        // Landing page: show when no data is bound, hide when data arrives
        if (this.handleLandingPage(options)) return;

        this.formattingSettings = this.formattingSettingsService.populateFormattingSettingsModel(VisualFormattingSettingsModel, options.dataViews[0]);

        // Dynamically scale UI base font-size to the viewport width
        // Assume 350px width is our "standard" 14px font size layout scaling
        if (options.viewport) {
            const scale = options.viewport.width / 350;
            const baseFontSize = Math.max(6, 14 * scale); // clamp to 6px min
            if (this.containerEl) {
                this.containerEl.style.setProperty('--base-font-size', `${baseFontSize}px`);
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
                                isoDateStr: normalizedDate.toISOString(),
                                dateInt: y * 10000 + (m + 1) * 100 + d.getDate()
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
                        // Bookmark support: restore selection state from saved jsonFilters
                        this.restoreBookmarkFilter(options);
                        this.refreshUI(false); // Render UI without re-applying filter back to host
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

    /** Landing page: shown when no data is bound. Returns true if landing page is active (caller should return early). */
    private handleLandingPage(options: VisualUpdateOptions): boolean {
        const hasData = options.dataViews
            && options.dataViews[0]
            && options.dataViews[0].metadata
            && options.dataViews[0].metadata.columns
            && options.dataViews[0].metadata.columns.length > 0;

        if (!hasData) {
            if (!this.isLandingPageOn) {
                this.isLandingPageOn = true;
                // Hide the normal calendar UI
                if (this.containerEl) this.containerEl.style.display = 'none';
                // Create landing page
                const lp = document.createElement('div');
                lp.className = 'landing-page';
                lp.innerHTML = `
                    <h1 class="landing-title">oneCalendar</h1>
                    <svg class="landing-icon" viewBox="0 0 120 120" xmlns="http://www.w3.org/2000/svg">
                        <rect x="4" y="4" width="112" height="112" rx="24" fill="#fce4ec" stroke="#e57373" stroke-width="3"/>
                        <text x="60" y="72" text-anchor="middle" font-family="'Segoe UI',sans-serif" font-size="48" font-weight="700" fill="#c62828">1C</text>
                    </svg>
                    <p class="landing-instruction">Add a <strong>Date</strong> field to start filtering</p>
                    <p class="landing-developer"><a href="https://www.linkedin.com/in/kirill-bezzubkin-b6860b235" target="_blank" rel="noopener noreferrer"><svg viewBox="0 0 24 24" width="14" height="14" fill="#0A66C2"><path d="M20.447 20.452h-3.554v-5.569c0-1.328-.027-3.037-1.852-3.037-1.853 0-2.136 1.445-2.136 2.939v5.667H9.351V9h3.414v1.561h.046c.477-.9 1.637-1.85 3.37-1.85 3.601 0 4.267 2.37 4.267 5.455v6.286zM5.337 7.433a2.062 2.062 0 01-2.063-2.065 2.064 2.064 0 112.063 2.065zm1.782 13.019H3.555V9h3.564v11.452zM22.225 0H1.771C.792 0 0 .774 0 1.729v20.542C0 23.227.792 24 1.771 24h20.451C23.2 24 24 23.227 24 22.271V1.729C24 .774 23.2 0 22.222 0h.003z"/></svg> Kirill</a></p>
                    <a class="landing-link" href="https://github.com/rusloc/oneCalendar" target="_blank" rel="noopener noreferrer">
                        <svg viewBox="0 0 16 16" width="14" height="14" fill="currentColor"><path d="M8 0C3.58 0 0 3.58 0 8c0 3.54 2.29 6.53 5.47 7.59.4.07.55-.17.55-.38 0-.19-.01-.82-.01-1.49-2.01.37-2.53-.49-2.69-.94-.09-.23-.48-.94-.82-1.13-.28-.15-.68-.52-.01-.53.63-.01 1.08.58 1.23.82.72 1.21 1.87.87 2.33.66.07-.52.28-.87.51-1.07-1.78-.2-3.64-.89-3.64-3.95 0-.87.31-1.59.82-2.15-.08-.2-.36-1.02.08-2.12 0 0 .67-.21 2.2.82.64-.18 1.32-.27 2-.27.68 0 1.36.09 2 .27 1.53-1.04 2.2-.82 2.2-.82.44 1.1.16 1.92.08 2.12.51.56.82 1.27.82 2.15 0 3.07-1.87 3.75-3.65 3.95.29.25.54.73.54 1.48 0 1.07-.01 1.93-.01 2.2 0 .21.15.46.55.38A8.01 8.01 0 0016 8c0-4.42-3.58-8-8-8z"/></svg>
                        GitHub
                    </a>
                `;
                this.target.appendChild(lp);
                this.landingPageEl = lp;
            }
            return true;
        }

        // Data is present — remove landing page if it was showing
        if (this.isLandingPageOn) {
            this.isLandingPageOn = false;
            if (this.landingPageEl) {
                this.landingPageEl.remove();
                this.landingPageEl = null;
            }
            if (this.containerEl) this.containerEl.style.display = '';
        }
        return false;
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