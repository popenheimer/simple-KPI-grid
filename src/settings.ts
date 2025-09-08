// src/settings.ts
"use strict";

import { formattingSettings } from "powerbi-visuals-utils-formattingmodel";

export class CardStyle extends formattingSettings.SimpleCard {
    public shapeFill = new formattingSettings.ColorPicker({
        name: "shapeFill",
        displayName: "Fill Color",
        value: { value: "#ffffff" }
    });

    public minWidth = new formattingSettings.NumUpDown({
        name: "minWidth",
        displayName: "Minimum Width",
        value: 50,
        visible: false
    });

    public minHeight = new formattingSettings.NumUpDown({
        name: "minHeight",
        displayName: "Minimum Height",
        value: 50,
        visible: false
    });

    public shapeBorderColor = new formattingSettings.ColorPicker({
        name: "shapeBorderColor",
        displayName: "Border Color",
        value: { value: "#E6E6E6" }
    });

    public shapeBorderWidth = new formattingSettings.NumUpDown({
        name: "shapeBorderWidth",
        displayName: "Border Width",
        value: 1
    });

    public shapeCornerRadius = new formattingSettings.NumUpDown({
        name: "shapeCornerRadius",
        displayName: "Corner Radius",
        value: 5
    });

    name: string = "cardStyle";
    displayName: string = "Card Style";
    slices = [this.shapeFill, this.minWidth, this.minHeight, this.shapeBorderColor, this.shapeBorderWidth, this.shapeCornerRadius];
}

export class MeasureName extends formattingSettings.SimpleCard {
    public measureNameColor = new formattingSettings.ColorPicker({
        name: "measureNameColor",
        displayName: "Color",
        value: { value: "#333333" }
    });

    public innerPadding = new formattingSettings.NumUpDown({
        name: "innerPadding",
        displayName: "Left/Right Padding",
        value: 5
    });

    public measureNameYOffset = new formattingSettings.NumUpDown({
        name: "measureNameYOffset",
        displayName: "Top Padding",
        value: 10
    });

    public measureNameFontSize = new formattingSettings.NumUpDown({
        name: "measureNameFontSize",
        displayName: "Font Size (pt)",
        value: 12
    });

    public measureNameFontFamily = new formattingSettings.FontPicker({
        name: "measureNameFontFamily",
        value: "Segoe UI, sans-serif"
    });

    public measureNameBold = new formattingSettings.ToggleSwitch({
        name: "measureNameBold",
        displayName: "Bold",
        value: false
    });

    public measureNameItalic = new formattingSettings.ToggleSwitch({
        name: "measureNameItalic",
        displayName: "Italic",
        value: false
    });

    public measureNameUnderline = new formattingSettings.ToggleSwitch({
        name: "measureNameUnderline",
        displayName: "Underline",
        value: false
    });

    public measureNameFont = new formattingSettings.FontControl({
        name: "measureNameFont",
        displayName: "Font",
        fontFamily: this.measureNameFontFamily,
        fontSize: this.measureNameFontSize,
        bold: this.measureNameBold,
        italic: this.measureNameItalic,
        underline: this.measureNameUnderline
    });

    name: string = "measureName";
    displayName: string = "Category";
    slices = [this.measureNameColor, this.measureNameYOffset, this.innerPadding, this.measureNameFont];
}

export class MeasureValue extends formattingSettings.SimpleCard {
    public measureValueColor = new formattingSettings.ColorPicker({
        name: "measureValueColor",
        displayName: "Color (if field not used)",
        value: { value: "#333333" }
    });

    public measureValueYOffset = new formattingSettings.NumUpDown({
        name: "measureValueYOffset",
        displayName: "Y Offset",
        value: 10
    });

    public measureValueFontSize = new formattingSettings.NumUpDown({
        name: "measureValueFontSize",
        displayName: "Font Size (pt)",
        value: 22
    });

    public measureValueFontFamily = new formattingSettings.FontPicker({
        name: "measureValueFontFamily",
        value: "Segoe UI, sans-serif"
    });

    public measureValueBold = new formattingSettings.ToggleSwitch({
        name: "measureValueBold",
        displayName: "Bold",
        value: false
    });

    public measureValueItalic = new formattingSettings.ToggleSwitch({
        name: "measureValueItalic",
        displayName: "Italic",
        value: false
    });

    public measureValueUnderline = new formattingSettings.ToggleSwitch({
        name: "measureValueUnderline",
        displayName: "Underline",
        value: false
    });

    public measureValueFont = new formattingSettings.FontControl({
        name: "measureValueFont",
        displayName: "Font",
        fontFamily: this.measureValueFontFamily,
        fontSize: this.measureValueFontSize,
        bold: this.measureValueBold,
        italic: this.measureValueItalic,
        underline: this.measureValueUnderline
    });

    name: string = "measureValue";
    displayName: string = "Value";
    slices = [this.measureValueColor, this.measureValueYOffset, this.measureValueFont];
}

export class ComparisonValue extends formattingSettings.SimpleCard {
    public comparisonValuePrefix = new formattingSettings.TextInput({
        name: "comparisonValuePrefix",
        displayName: "Prefix Text",
        value: "Baseline: ",
        placeholder: "Baseline: "
    });

    public comparisonValueDelta = new formattingSettings.AutoDropdown({
        name: "comparisonValueDelta",
        displayName: "Delta Type",
        value: "none"
    });

    public comparisonValueYOffset = new formattingSettings.NumUpDown({
        name: "comparisonValueYOffset",
        displayName: "Y Offset",
        value: 5
    });

    public comparisonValueColor = new formattingSettings.ColorPicker({
        name: "comparisonValueColor",
        displayName: "Color",
        value: { value: "#333333" }
    });

    public comparisonValueFontSize = new formattingSettings.NumUpDown({
        name: "comparisonValueFontSize",
        displayName: "Font Size (pt)",
        value: 10
    });

    public comparisonValueFontFamily = new formattingSettings.FontPicker({
        name: "comparisonValueFontFamily",
        value: "Segoe UI, sans-serif"
    });

    public comparisonValueBold = new formattingSettings.ToggleSwitch({
        name: "comparisonValueBold",
        displayName: "Bold",
        value: false
    });

    public comparisonValueItalic = new formattingSettings.ToggleSwitch({
        name: "comparisonValueItalic",
        displayName: "Italic",
        value: false
    });

    public comparisonValueUnderline = new formattingSettings.ToggleSwitch({
        name: "comparisonValueUnderline",
        displayName: "Underline",
        value: false
    });

    public comparisonValueFont = new formattingSettings.FontControl({
        name: "comparisonValueFont",
        displayName: "Font",
        fontFamily: this.comparisonValueFontFamily,
        fontSize: this.comparisonValueFontSize,
        bold: this.comparisonValueBold,
        italic: this.comparisonValueItalic,
        underline: this.comparisonValueUnderline
    });

    name: string = "comparisonValue";
    displayName: string = "Comparison Value";
    slices = [this.comparisonValuePrefix, this.comparisonValueDelta, this.comparisonValueYOffset, this.comparisonValueColor, this.comparisonValueFont];
}

export class GridSettings extends formattingSettings.SimpleCard {
    public sizingMode = new formattingSettings.AutoDropdown({
        name: "sizingMode",
        displayName: "Sizing Mode",
        value: "fixed"
    });

    public fixedWidth = new formattingSettings.NumUpDown({
        name: "fixedWidth",
        displayName: "Card Width",
        value: 210
    });

    public fixedHeight = new formattingSettings.NumUpDown({
        name: "fixedHeight",
        displayName: "Card Height",
        value: 130
    });

    public columns = new formattingSettings.NumUpDown({
        name: "columns",
        displayName: "# of Columns (0 for auto)",
        value: 0
    });

    public rows = new formattingSettings.NumUpDown({
        name: "rows",
        displayName: "# of Rows (0 for auto)",
        value: 0
    });

    public margin = new formattingSettings.NumUpDown({
        name: "margin",
        displayName: "Margin Between Cards",
        value: 12
    });

    name: string = "gridSettings";
    displayName: string = "Grid Settings";
    slices = [this.sizingMode, this.fixedWidth, this.fixedHeight, this.columns, this.rows, this.margin];
}

export class StateSettings extends formattingSettings.SimpleCard {
    
    public allowSelections = new formattingSettings.ToggleSwitch({
        name: "allowSelections",
        displayName: "Allow Selections",
        value: false
    });

    public hoverFill = new formattingSettings.ColorPicker({
        name: "hoverFill",
        displayName: "Hover Fill Color",
        value: { value: "#e1e1e1" }
    });

    public activeFill = new formattingSettings.ColorPicker({
        name: "activeFill",
        displayName: "Pressed Fill Color",
        value: { value: "#c3c3c3" }
    });

    public selectedFill = new formattingSettings.ColorPicker({
        name: "selectedFill",
        displayName: "Selected Fill Color",
        value: { value: "#d2d2d2" }
    });

    public selectedDefault = new formattingSettings.ToggleSwitch({
        name: "selectedDefault",
        displayName: "Enable Default Selection",
        value: false
    });

    public selectedDefaultMeasureName = new formattingSettings.TextInput({
        name: "selectedDefaultMeasureName",
        displayName: "Category Name (0 for first)",
        value: "0",
        placeholder: "Enter Category Name",
        visible: this.selectedDefault.value
    });

    name: string = "stateSettings";
    displayName: string = "Allow Selection";
    topLevelSlice: formattingSettings.ToggleSwitch = this.allowSelections;
    slices = [this.hoverFill, this.activeFill, this.selectedFill, this.selectedDefault, this.selectedDefaultMeasureName];
}

export class VisualSettings extends formattingSettings.Model {
    public cardStyle: CardStyle = new CardStyle();
    public measureName: MeasureName = new MeasureName();
    public measureValue: MeasureValue = new MeasureValue();
    public comparisonValue: ComparisonValue = new ComparisonValue();
    public gridSettings: GridSettings = new GridSettings();
    public stateSettings: StateSettings = new StateSettings();
    public cards = [this.gridSettings, this.cardStyle, this.measureName, this.measureValue, this.comparisonValue, this.stateSettings];
}