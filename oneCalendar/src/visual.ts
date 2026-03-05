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
}

export class Visual implements IVisual {
    private static readonly BLOCK_IDS = ['year', 'quarter', 'month', 'week'] as const;

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

    private target: HTMLElement;
    private host: IVisualHost;
    private formattingSettings: VisualFormattingSettingsModel;
    private formattingSettingsService: FormattingSettingsService;

    private dataViewParsed: boolean = false;
    private mainPanelExpanded: boolean = true;
    private blocksExpanded: { [key: string]: boolean } = { year: false, quarter: false, month: false, week: false };

    private dataPoints: DateDataPoint[] = [];
    private masterDataPoints: DateDataPoint[] = [];  // Full unfiltered dataset cache
    private currentCategory: powerbi.DataViewCategoryColumn | null = null;
    private selections: { [key: string]: Set<string> } = {
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
                    <button id="btn-reset" class="btn-icon" title="Reset Selections" disabled>${Visual.SVG_RESET}</button>
                    <div id="filter-status" class="filter-status">No filters</div>
                </div>
                <div id="main-panel" class="main-panel collapsed">
                    <div class="empty-state">Please add a Date field.</div>
                </div>
            </div>
        `;
        this.bindStaticEvents();
    }

    private bindStaticEvents() {
        const btnCollapse = this.target.querySelector('#btn-collapse');
        const mainPanel = this.target.querySelector('#main-panel') as HTMLElement;

        btnCollapse?.addEventListener('click', () => {
            if (!this.dataViewParsed) return;
            this.mainPanelExpanded = !this.mainPanelExpanded;
            if (this.mainPanelExpanded) {
                mainPanel.classList.remove('collapsed');
                btnCollapse.innerHTML = Visual.SVG_COLLAPSE;
            } else {
                mainPanel.classList.add('collapsed');
                btnCollapse.innerHTML = Visual.SVG_EXPAND;
            }
        });

        const btnReset = this.target.querySelector('#btn-reset');
        btnReset?.addEventListener('click', () => {
            if (!this.dataViewParsed) return;
            this.selections = {
                year: new Set<string>(),
                quarter: new Set<string>(),
                month: new Set<string>(),
                week: new Set<string>()
            };
            this.lastSelectedIndex = { year: null, quarter: null, month: null, week: null };
            this.dateFrom = null;
            this.dateTo = null;
            // Clear the JSON filter when resetting
            this.host.applyJsonFilter(null, "general", "filter", powerbi.FilterAction.remove);
            this.refreshUI(false); // just redraw, filter already cleared above
        });

        const btnToggleBlocks = this.target.querySelector('#btn-toggle-blocks');
        btnToggleBlocks?.addEventListener('click', () => {
            if (!this.dataViewParsed) return;

            // Check if ANY are collapsed. If any are collapsed, expanding all is intuitive.
            // If all are already expanded, we collapse all.
            const anyCollapsed = Object.values(this.blocksExpanded).some(state => state === false);
            const newState = anyCollapsed;

            this.blocksExpanded = {
                year: newState,
                quarter: newState,
                month: newState,
                week: newState
            };

            this.refreshUI(false);
        });
    }

    private refreshUI(applyFilter: boolean = false) {
        if (!this.dataPoints || this.dataPoints.length === 0) {
            this.clearPanel("No dates available.");
            return;
        }

        const originalDates = this.dataPoints.map(dp => dp.date);
        const rootInfo = this.extractDateInformation(originalDates);
        const availableYears = rootInfo.years;

        // Clean invalid year selections
        for (let y of this.selections.year) {
            if (!availableYears.includes(y)) this.selections.year.delete(y);
        }

        // Filter dates by Year
        const dpAfterYear = this.dataPoints.filter(dp =>
            this.selections.year.size === 0 || this.selections.year.has(String(dp.date.getFullYear()))
        );

        // Quarters are extracted from dates filtered by year
        const qInfo = this.extractDateInformation(dpAfterYear.map(dp => dp.date));
        const availableQuarters = qInfo.quarters;

        // Clean invalid quarter selections
        for (let q of this.selections.quarter) {
            if (!availableQuarters.includes(q)) this.selections.quarter.delete(q);
        }

        // Filter dates by Quarter
        const dpAfterQuarter = dpAfterYear.filter(dp => {
            if (this.selections.quarter.size === 0) return true;
            const q = Math.floor(dp.date.getMonth() / 3) + 1;
            const qLabel = `Q${q} ${dp.date.getFullYear()}`;
            return this.selections.quarter.has(qLabel);
        });

        // Months are extracted from dates filtered by Quarter
        const mInfo = this.extractDateInformation(dpAfterQuarter.map(dp => dp.date));
        const availableMonths = mInfo.months;

        // Clean invalid month selections
        for (let m of this.selections.month) {
            if (!availableMonths.includes(m)) this.selections.month.delete(m);
        }

        // Filter dates by Month
        const dpAfterMonth = dpAfterQuarter.filter(dp => {
            if (this.selections.month.size === 0) return true;
            const mLabel = `${dp.date.toLocaleString('default', { month: 'short' })} ${dp.date.getFullYear()}`;
            return this.selections.month.has(mLabel);
        });

        // Weeks are extracted from dates filtered by Month
        const wInfo = this.extractDateInformation(dpAfterMonth.map(dp => dp.date));
        const availableWeeks = wInfo.weeks;

        // Clean invalid week selections
        for (let w of this.selections.week) {
            if (!availableWeeks.includes(w)) this.selections.week.delete(w);
        }

        // Filter dates by Week
        const dpAfterWeek = dpAfterMonth.filter(dp => {
            if (this.selections.week.size === 0) return true;
            const weekNum = this.getISOWeekNumber(dp.date);
            const wLabel = `W${weekNum}`;
            return this.selections.week.has(wLabel);
        });

        // Finally filter visually by date ranges
        const finalDataPoints = dpAfterWeek.filter(dp => {
            if (this.dateFrom && dp.date < this.dateFrom) return false;
            if (this.dateTo && dp.date > this.dateTo) return false;
            return true;
        });

        // Find visual min max (without input strict bounds if user didn't type them, purely based on week/month bounds)
        const unconstrainedMinDate = dpAfterWeek.length ? new Date(Math.min(...dpAfterWeek.map(dp => dp.date.getTime()))) : null;
        const unconstrainedMaxDate = dpAfterWeek.length ? new Date(Math.max(...dpAfterWeek.map(dp => dp.date.getTime()))) : null;

        const displayInfo = {
            years: availableYears,
            quarters: availableQuarters,
            months: availableMonths,
            weeks: availableWeeks,
            minDate: unconstrainedMinDate,
            maxDate: unconstrainedMaxDate
        };

        this.renderDynamicContent(displayInfo);

        // Apply JSON filter logic (slicer-like behavior)
        if (applyFilter && this.currentCategory) {
            const hasSelections = this.selections.year.size > 0 ||
                this.selections.quarter.size > 0 ||
                this.selections.month.size > 0 ||
                this.selections.week.size > 0 ||
                this.dateFrom !== null ||
                this.dateTo !== null;

            if (hasSelections && finalDataPoints.length > 0) {
                const queryName = this.currentCategory.source.queryName || "Calendar.Date";
                let tableName = queryName;
                if (queryName.includes('.')) {
                    tableName = queryName.substring(0, queryName.indexOf('.'));
                }
                tableName = tableName.replace(/"/g, ""); // strip quotes

                const target = {
                    table: tableName,
                    column: this.currentCategory.source.displayName
                };

                // Build unique date ISO strings for filtering
                const uniqueDates = Array.from(new Set(finalDataPoints.map(dp => {
                    const rDate = new Date(dp.date.getFullYear(), dp.date.getMonth(), dp.date.getDate());
                    return rDate.toISOString();
                })));

                const basicFilter: any = {
                    $schema: "http://powerbi.com/product/schema#basic",
                    target: target,
                    operator: "In",
                    values: uniqueDates,
                    filterType: 1 // BasicFilter type code
                };

                this.host.applyJsonFilter(basicFilter, "general", "filter", powerbi.FilterAction.merge);
            } else {
                this.host.applyJsonFilter(null, "general", "filter", powerbi.FilterAction.remove);
            }
        }
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
        const mainPanel = this.target.querySelector('#main-panel') as HTMLElement;
        const btnCollapse = this.target.querySelector('#btn-collapse') as HTMLButtonElement;
        const btnToggleBlocks = this.target.querySelector('#btn-toggle-blocks') as HTMLButtonElement;
        const btnReset = this.target.querySelector('#btn-reset') as HTMLButtonElement;

        if (btnCollapse) btnCollapse.disabled = false;
        if (btnReset) btnReset.disabled = false;

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

        const container = this.target.querySelector('.calendar-container') as HTMLElement;
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

        const buildBlock = (id: string, label: string, items: string[]) => {
            const isExpanded = this.blocksExpanded[id];

            let itemHtml = '';
            items.forEach((item: string, idx: number) => {
                const isSelected = this.selections[id].has(item);
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
        html += buildBlock('year', 'Year', dateInfo.years);
        html += buildBlock('quarter', 'Quarter', dateInfo.quarters);
        html += buildBlock('month', 'Month', dateInfo.months);
        html += buildBlock('week', 'Week', dateInfo.weeks);

        // Apply visual constraints but populate with active specific selections if made
        const inputMinBound = this.formatDateForInput(dateInfo.minDate);
        const inputMaxBound = this.formatDateForInput(dateInfo.maxDate);

        const currentFromValue = this.dateFrom ? this.formatDateForInput(this.dateFrom) : inputMinBound;
        const currentToValue = this.dateTo ? this.formatDateForInput(this.dateTo) : inputMaxBound;

        html += `
            <div class="dates-block">
                <span class="block-title">Dates</span>
                <div class="dates-inputs">
                    <input id="date-from" type="date" value="${currentFromValue}" min="${inputMinBound}" max="${inputMaxBound}" />
                    <span class="separator">–</span>
                    <input id="date-to" type="date" value="${currentToValue}" min="${inputMinBound}" max="${inputMaxBound}" />
                </div>
            </div>
        `;

        if (mainPanel) {
            mainPanel.innerHTML = html;

            // Update filter status indicator
            const filterStatus = this.target.querySelector('#filter-status');
            if (filterStatus) {
                const activeFilters: string[] = [];
                if (this.selections.year.size > 0) activeFilters.push('Year');
                if (this.selections.quarter.size > 0) activeFilters.push('Qrt');
                if (this.selections.month.size > 0) activeFilters.push('Mon');
                if (this.selections.week.size > 0) activeFilters.push('Week');
                if (this.dateFrom || this.dateTo) activeFilters.push('Date');
                filterStatus.textContent = activeFilters.length > 0 ? 'Filters: ' + activeFilters.join(', ') : 'No filters';
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
        blockItems.forEach(item => {
            item.addEventListener('click', (e: Event) => {
                const mouseEvent = e as MouseEvent;
                const target = mouseEvent.target as HTMLElement;
                const blockId = target.getAttribute('data-id');
                const indexStr = target.getAttribute('data-index');
                const valStr = target.getAttribute('data-val');

                if (!blockId || !indexStr || !valStr) return;
                const idx = parseInt(indexStr);

                if (mouseEvent.shiftKey && this.lastSelectedIndex[blockId] !== null) {
                    const start = Math.min(this.lastSelectedIndex[blockId]!, idx);
                    const end = Math.max(this.lastSelectedIndex[blockId]!, idx);

                    const currentGroupItems = Array.from(this.target.querySelectorAll(`.block-item[data-id="${blockId}"]`));
                    currentGroupItems.forEach((el) => {
                        const elIdx = parseInt(el.getAttribute('data-index')!);
                        const elVal = el.getAttribute('data-val')!;
                        if (elIdx >= start && elIdx <= end) {
                            this.selections[blockId].add(elVal);
                        }
                    });
                } else {
                    if (this.selections[blockId].has(valStr)) {
                        this.selections[blockId].delete(valStr);
                        this.lastSelectedIndex[blockId] = null;
                    } else {
                        // Standard click toggles selection, it cascades naturally
                        this.selections[blockId].add(valStr);
                        this.lastSelectedIndex[blockId] = idx;
                    }
                }

                // Call refresh UI computing final filters, pass true to apply limits
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

                            incomingDataPoints.push({
                                date: d,
                                selectionId: selectionId
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
        const mainPanel = this.target.querySelector('#main-panel') as HTMLElement;
        const btnCollapse = this.target.querySelector('#btn-collapse') as HTMLButtonElement;
        const btnToggleBlocks = this.target.querySelector('#btn-toggle-blocks') as HTMLButtonElement;
        const btnReset = this.target.querySelector('#btn-reset') as HTMLButtonElement;

        if (btnCollapse) {
            btnCollapse.disabled = true;
            btnCollapse.innerHTML = Visual.SVG_EXPAND;
        }
        if (btnToggleBlocks) {
            btnToggleBlocks.disabled = true;
            btnToggleBlocks.innerHTML = Visual.SVG_TOGGLE_EXPAND;
            btnToggleBlocks.removeAttribute('data-expanded');
        }
        if (btnReset) btnReset.disabled = true;
        this.mainPanelExpanded = true;
        if (mainPanel) {
            mainPanel.classList.remove('collapsed');
            mainPanel.innerHTML = `<div class="empty-state">${msg}</div>`;
        }
    }

    public getFormattingModel(): powerbi.visuals.FormattingModel {
        return this.formattingSettingsService.buildFormattingModel(this.formattingSettings);
    }
}