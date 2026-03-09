/*
 *  Power BI Visualizations
 *
 *  Copyright (c) Microsoft Corporation
 *  All rights reserved.
 *  MIT License
 *
 *  Permission is hereby granted, free of charge, to any person obtaining a copy
 *  of this software and associated documentation files (the ""Software""), to deal
 *  in the Software without restriction, including without limitation the rights
 *  to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 *  copies of the Software, and to permit persons to whom the Software is
 *  furnished to do so, subject to the following conditions:
 *
 *  The above copyright notice and this permission notice shall be included in
 *  all copies or substantial portions of the Software.
 *
 *  THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 *  IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 *  FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 *  AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 *  LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 *  OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
 *  THE SOFTWARE.
 */

"use strict";

import { formattingSettings } from "powerbi-visuals-utils-formattingmodel";

import FormattingSettingsCard = formattingSettings.SimpleCard;
import FormattingSettingsSlice = formattingSettings.Slice;
import FormattingSettingsModel = formattingSettings.Model;

/**
 * Visual Settings Card
 */
class VisualSettingsCard extends FormattingSettingsCard {
    accentColor = new formattingSettings.ColorPicker({
        name: "accentColor",
        displayName: "Accent Color",
        value: { value: "#3b82f6" } // Default pleasant blue
    });

    buttonPosition = new formattingSettings.ItemDropdown({
        name: "buttonPosition",
        displayName: "Button Position",
        items: [
            { value: "left", displayName: "Left" },
            { value: "right", displayName: "Right" }
        ],
        value: { value: "left", displayName: "Left" }
    });

    containerBorderWeight = new formattingSettings.NumUpDown({
        name: "containerBorderWeight",
        displayName: "Container Border Weight",
        value: 1
    });

    containerBorderColor = new formattingSettings.ColorPicker({
        name: "containerBorderColor",
        displayName: "Container Border Color",
        value: { value: "#e5e7eb" }
    });

    datesBgColor = new formattingSettings.ColorPicker({
        name: "datesBgColor",
        displayName: "Dates Block Background",
        value: { value: "#fafafa" }
    });

    name: string = "visualSettings";
    displayName: string = "Visual Settings";
    slices: Array<FormattingSettingsSlice> = [this.accentColor, this.buttonPosition, this.containerBorderWeight, this.containerBorderColor, this.datesBgColor];
}

/**
* visual settings model class
*
*/
export class VisualFormattingSettingsModel extends FormattingSettingsModel {
    // Create formatting settings model formatting cards
    visualSettingsCard = new VisualSettingsCard();

    cards = [this.visualSettingsCard];
}
