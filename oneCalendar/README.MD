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