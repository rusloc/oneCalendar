import { formattingSettings } from "powerbi-visuals-utils-formattingmodel";
import FormattingSettingsCard = formattingSettings.SimpleCard;
import FormattingSettingsSlice = formattingSettings.Slice;
import FormattingSettingsModel = formattingSettings.Model;
/**
 * Visual Settings Card
 */
declare class VisualSettingsCard extends FormattingSettingsCard {
    accentColor: formattingSettings.ColorPicker;
    buttonPosition: formattingSettings.ItemDropdown;
    containerBorderWeight: formattingSettings.NumUpDown;
    containerBorderColor: formattingSettings.ColorPicker;
    datesBgColor: formattingSettings.ColorPicker;
    name: string;
    displayName: string;
    slices: Array<FormattingSettingsSlice>;
}
/**
* visual settings model class
*
*/
export declare class VisualFormattingSettingsModel extends FormattingSettingsModel {
    visualSettingsCard: VisualSettingsCard;
    cards: VisualSettingsCard[];
}
export {};
