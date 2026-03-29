import powerbi from "powerbi-visuals-api";
import "./../style/visual.less";
import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisual = powerbi.extensibility.visual.IVisual;
export declare class Visual implements IVisual {
    private static readonly BLOCK_IDS;
    private static readonly MONTH_NAMES;
    private static readonly SVG_EXPAND;
    private static readonly SVG_COLLAPSE;
    private static readonly SVG_TOGGLE_EXPAND;
    private static readonly SVG_TOGGLE_COLLAPSE;
    private static readonly SVG_RESET;
    private static readonly SVG_ARROW_DOWN;
    private static readonly SVG_ARROW_UP;
    private static readonly SVG_VIEW_SWITCH;
    private target;
    private host;
    private formattingSettings;
    private formattingSettingsService;
    private dataViewParsed;
    private mainPanelExpanded;
    private blocksExpanded;
    private viewMode;
    private dataPoints;
    private masterDataPoints;
    private currentCategory;
    private mainPanelEl;
    private btnCollapseEl;
    private btnToggleBlocksEl;
    private btnResetEl;
    private btnViewSwitchEl;
    private filterStatusEl;
    private containerEl;
    private selections;
    private lastSelections;
    private lastSelectedIndex;
    private dateFrom;
    private dateTo;
    private renderedBlockItems;
    private renderedViewMode;
    private renderedDateBounds;
    private delegatedListenersBound;
    private lastAppliedFilterKey;
    private isLandingPageOn;
    private landingPageEl;
    constructor(options: VisualConstructorOptions);
    private bindStaticEvents;
    private clearAllSelections;
    private refreshUI;
    private refreshUILast;
    /** Shared filter application — deduplicates logic between Normal & Last modes.
     *  Uses hybrid strategy: AdvancedFilter (min-max range, ~200 bytes) when selected dates
     *  are contiguous; falls back to BasicFilter "In" (full date list) when gaps exist. */
    private applyJsonFilterFromDataPoints;
    /** Bookmark support: restore internal selection state from saved jsonFilters.
     *  Determines the lowest selection level that fully explains the bookmarked dates,
     *  respecting the cascading filter logic (year > quarter > month > week).
     *  Upper levels auto-select only to scope; lower levels stay unselected (filtered, not selected). */
    private restoreBookmarkFilter;
    private getISOWeekNumber;
    private formatDateForInput;
    /** Update the filter status badge in the header */
    private updateFilterStatus;
    /** Bind delegated event listeners on mainPanel once (P4) */
    private bindDelegatedListeners;
    private renderDynamicContent;
    update(options: VisualUpdateOptions): void;
    /** Landing page: shown when no data is bound. Returns true if landing page is active (caller should return early). */
    private handleLandingPage;
    private clearPanel;
    getFormattingModel(): powerbi.visuals.FormattingModel;
}
