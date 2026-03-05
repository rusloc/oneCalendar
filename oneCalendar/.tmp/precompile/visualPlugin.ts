import { Visual } from "../../src/visual";
import powerbiVisualsApi from "powerbi-visuals-api";
import IVisualPlugin = powerbiVisualsApi.visuals.plugins.IVisualPlugin;
import VisualConstructorOptions = powerbiVisualsApi.extensibility.visual.VisualConstructorOptions;
import DialogConstructorOptions = powerbiVisualsApi.extensibility.visual.DialogConstructorOptions;
var powerbiKey: any = "powerbi";
var powerbi: any = window[powerbiKey];
var oneCalendarD83DAFDDCC8D4309935B4488E6A78767_DEBUG: IVisualPlugin = {
    name: 'oneCalendarD83DAFDDCC8D4309935B4488E6A78767_DEBUG',
    displayName: 'oneCalendar',
    class: 'Visual',
    apiVersion: '5.3.0',
    create: (options?: VisualConstructorOptions) => {
        if (Visual) {
            return new Visual(options);
        }
        throw 'Visual instance not found';
    },
    createModalDialog: (dialogId: string, options: DialogConstructorOptions, initialState: object) => {
        const dialogRegistry = (<any>globalThis).dialogRegistry;
        if (dialogId in dialogRegistry) {
            new dialogRegistry[dialogId](options, initialState);
        }
    },
    custom: true
};
if (typeof powerbi !== "undefined") {
    powerbi.visuals = powerbi.visuals || {};
    powerbi.visuals.plugins = powerbi.visuals.plugins || {};
    powerbi.visuals.plugins["oneCalendarD83DAFDDCC8D4309935B4488E6A78767_DEBUG"] = oneCalendarD83DAFDDCC8D4309935B4488E6A78767_DEBUG;
}
export default oneCalendarD83DAFDDCC8D4309935B4488E6A78767_DEBUG;