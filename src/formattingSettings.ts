import { formattingSettings } from "powerbi-visuals-utils-formattingmodel";
import powerbi from "powerbi-visuals-api";

export class YearRangeCardSettings extends formattingSettings.SimpleCard {
    public textColor = new formattingSettings.ColorPicker({
        name: "textColor",
        displayName: "Text color",
        value: { value: "#000000" }
    });

    public fontSize = new formattingSettings.NumUpDown({
        name: "fontSize",
        displayName: "Font size",
        value: 14,
        options: {
            minValue: {
                value: 8,
                type: powerbi.visuals.ValidatorType.Min
            },
            maxValue: {
                value: 32,
                type: powerbi.visuals.ValidatorType.Max
            }
        }
    });

    public borderRadius = new formattingSettings.NumUpDown({
        name: "borderRadius",
        displayName: "Border radius",
        value: 4,
        options: {
            minValue: {
                value: 0,
                type: powerbi.visuals.ValidatorType.Min
            },
            maxValue: {
                value: 20,
                type: powerbi.visuals.ValidatorType.Max
            }
        }
    });

    public name: string = "yearRange";
    public displayName: string = "Year Range";

    public slices: formattingSettings.Slice[] = [
        this.textColor,
        this.fontSize,
        this.borderRadius
    ];
}

export class YearRangeFormattingSettingsModel extends formattingSettings.Model {
    public yearRangeCard: YearRangeCardSettings = new YearRangeCardSettings();

    public cards: formattingSettings.SimpleCard[] = [
        this.yearRangeCard
    ];
}