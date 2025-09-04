// src/visual.ts
"use strict";

import powerbi from "powerbi-visuals-api";
import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisual = powerbi.extensibility.visual.IVisual;
import IVisualHost = powerbi.extensibility.visual.IVisualHost;
import VisualUpdateType = powerbi.VisualUpdateType;
import ISelectionManager = powerbi.extensibility.ISelectionManager;
import ISelectionId = powerbi.visuals.ISelectionId;
import ITooltipService = powerbi.extensibility.ITooltipService;
import VisualTooltipDataItem = powerbi.extensibility.VisualTooltipDataItem;
import EnumerateVisualObjectInstancesOptions = powerbi.EnumerateVisualObjectInstancesOptions;
import VisualObjectInstanceEnumerationObject = powerbi.VisualObjectInstanceEnumerationObject;
import VisualObjectInstance = powerbi.VisualObjectInstance;

import * as d3 from "d3";
import "./../style/visual.less";
type Selection<T extends d3.BaseType> = d3.Selection<T, any, any, any>;

import DataView = powerbi.DataView;
import { valueFormatter, textMeasurementService, interfaces } from "powerbi-visuals-utils-formattingutils";

import { GridSettings, VisualSettings } from "./settings";
import { FormattingSettingsService } from "powerbi-visuals-utils-formattingmodel";

interface KpiDataPoint {
    measureName: string;
    displayValue: string;
    measureValue: number;
    textColor: string;
    comparisonDisplay: string;
    deltaPercent: string;
    deltaAbsolute: string;
    formatString: string;
    comparisonFormat: string;
    selectionId: ISelectionId;
    categoryIndex: number;
    highlighted: boolean;
    selected: boolean;
    tooltipItems: VisualTooltipDataItem[];
}

export class Visual implements IVisual {
    private host: IVisualHost;
    private svg: Selection<SVGSVGElement>;
    private formattingSettings: VisualSettings;
    private formattingSettingsService: FormattingSettingsService;
    private selectionManager: ISelectionManager;
    private isFirstUpdate: boolean = true;
    private tooltipService: ITooltipService;
    private cards: Selection<SVGGElement>;
    private scrollbarWidth: number = 0;
    private isSingleKpi: boolean = false;

    constructor(options: VisualConstructorOptions) {
        this.host = options.host;
        const localizationManager = this.host.createLocalizationManager();
        this.formattingSettingsService = new FormattingSettingsService(localizationManager);
        this.svg = d3.select(options.element).append("svg").classed("simpleRectangleVisual", true);
        this.selectionManager = this.host.createSelectionManager();
        this.tooltipService = this.host.tooltipService;

        // ADDED: Enable scrolling on the host container when content overflows
        d3.select(options.element).classed("scrollable-container", true).style("overflow", "auto");

        // ADDED: Calculate scrollbar dimensions once
        this.scrollbarWidth = this.getScrollbarWidth(options.element);

        this.selectionManager.registerOnSelectCallback((ids: ISelectionId[]) => {
            if (this.cards && this.formattingSettings) {
                this.syncSelectionState(this.cards, this.formattingSettings);
            }
        });
    }

    // ADDED: Helper to measure scrollbar width
    private getScrollbarWidth(element: HTMLElement): number {
        const outer = document.createElement('div');
        outer.style.visibility = 'hidden';
        outer.style.overflow = 'scroll';
        //outer.style.msOverflowStyle = 'scrollbar';  // For IE/Edge
        element.appendChild(outer);

        const inner = document.createElement('div');
        outer.appendChild(inner);

        const scrollbarWidth = outer.offsetWidth - inner.offsetWidth;

        outer.parentNode?.removeChild(outer);  // Clean up

        return scrollbarWidth;
    }

    public update(options: VisualUpdateOptions) {
        if (options.type & (VisualUpdateType.Data | VisualUpdateType.Resize | VisualUpdateType.Style)) {
            this.formattingSettings = this.formattingSettingsService.populateFormattingSettingsModel(VisualSettings, options.dataViews[0]);
        }

        const dataView: DataView = options.dataViews[0];
        const settings: VisualSettings = this.formattingSettings || new VisualSettings();

        this.svg.selectAll("*").remove();

        const dataPoints: KpiDataPoint[] = this.parseDataPoints(dataView, settings);
        this.applyDefaultSelection(dataPoints, settings);

        const vpWidth: number = options.viewport.width;
        const vpHeight: number = options.viewport.height;

        const MAX_CARDS = 1000; // Arbitrary limit to prevent performance issues

        // Handle no data or too many data case
        if (dataPoints.length === 0) {
            // Set SVG to viewport size for no-data message
            this.svg.attr("width", vpWidth - 5).attr("height", vpHeight - 5);
            // Render no-data message
            const fontSize = Math.min(vpWidth, vpHeight) / 18;
            this.svg.append("text")
                .attr("x", vpWidth / 2)
                .attr("y", vpHeight / 2 - 20) // Starting position
                .attr("text-anchor", "middle")
                .style("font-size", `${fontSize}px`)
                .append("tspan")
                .text("No data added.")
                .attr("x", vpWidth / 2)
                .attr("dy", "0.35em") // First line
                .append("tspan")
                .text("Add a Value for a single card")
                .attr("x", vpWidth / 2)
                .attr("dy", "1.2em") // Second line, adjust dy for spacing
                .append("tspan")
                .text("and a Category to display a KPI grid.")
                .attr("x", vpWidth / 2)
                .attr("dy", "1.2em"); // Third line, adjust dy for spacing

            return;
        } else if (dataPoints.length > MAX_CARDS) {
            // Set SVG to viewport size for no-data message
            this.svg.attr("width", vpWidth - 5).attr("height", vpHeight - 5);
            // Render no-data message
            const fontSize = Math.min(vpWidth, vpHeight) / 18;
            this.svg.append("text")
                .attr("x", vpWidth / 2)
                .attr("y", vpHeight / 2 - 20)
                .attr("text-anchor", "middle")
                .style("font-size", `${fontSize}px`)
                .style("fill", "red")  // Use red for emphasis
                .append("tspan")
                .text(`Too many KPIs (${dataPoints.length} items).`)
                .attr("x", vpWidth / 2)
                .attr("dy", "0.35em")
                .append("tspan")
                .text("Apply filters to reduce categories")
                .attr("x", vpWidth / 2)
                .attr("dy", "1.2em")
                .append("tspan")
                .text("for better performance.")
                .attr("x", vpWidth / 2)
                .attr("dy", "1.2em");

            console.warn(`Visual warning: DataPoints exceed limit (${dataPoints.length} > ${MAX_CARDS}). Suggest filtering.`);  // For dev debugging
            return;
        }

        // First-pass calculation (assume no scrollbars)
        let { cardWidth, cardHeight, columns, rows } = this.calculateGridDimensions(dataPoints.length, vpWidth, vpHeight, settings);

        const gridMargin = settings.gridSettings.margin.value || 12;
        let totalWidth = columns * (cardWidth + gridMargin) - gridMargin;
        let totalHeight = rows * (cardHeight + gridMargin) - gridMargin;

        const sizingMode = settings.gridSettings.sizingMode.value;

        // Handle dynamicWidth: Adjust for vertical scrollbar reducing width
        if (sizingMode === "dynamicWidth" && totalHeight > vpHeight && this.scrollbarWidth > 0) {
            const effectiveVpWidth = vpWidth - this.scrollbarWidth;
            ({ cardWidth, cardHeight, columns, rows } = this.calculateGridDimensions(dataPoints.length, effectiveVpWidth, vpHeight, settings));
            totalWidth = columns * (cardWidth + gridMargin) - gridMargin;
            totalHeight = rows * (cardHeight + gridMargin) - gridMargin;
        }

        // Handle dynamicHeight: Adjust for horizontal scrollbar reducing height
        if (sizingMode === "dynamicHeight" && totalWidth > vpWidth && this.scrollbarWidth > 0) {
            const effectiveVpHeight = vpHeight - this.scrollbarWidth - 1;
            ({ cardWidth, cardHeight, columns, rows } = this.calculateGridDimensions(dataPoints.length, vpWidth, effectiveVpHeight, settings));
            totalWidth = columns * (cardWidth + gridMargin) - gridMargin;
            totalHeight = rows * (cardHeight + gridMargin) - gridMargin;
        }
        const hasCategories = dataView.categorical?.categories?.length > 0 || false;
        this.isSingleKpi = !hasCategories && dataPoints.length === 1;

        // For single measure (no categories), force single full-size card, ignoring row/column formatting
        if (this.isSingleKpi) {
            columns = 1;
            rows = 1;
            cardWidth = vpWidth;
            cardHeight = vpHeight;
            totalWidth = cardWidth;
            totalHeight = cardHeight;
            d3.select(this.svg.node().parentNode as HTMLElement).style("overflow", "hidden");
        } else {
            d3.select(this.svg.node().parentNode as HTMLElement).style("overflow", "auto");
        }

        // Set SVG to content size (host div will scroll if needed)
        this.svg.attr("width", totalWidth).attr("height", totalHeight);

        this.cards = this.renderCards(dataPoints, settings, cardWidth, cardHeight, columns, rows, hasCategories);

        if (hasCategories) {
            this.syncSelectionState(this.cards, settings);
        }

        this.svg.on("click", () => this.selectionManager.clear().then(() => {
            this.syncSelectionState(this.cards, settings);
        }));
    }

    private parseDataPoints(dataView: DataView, settings: VisualSettings): KpiDataPoint[] {
        const dataPoints: KpiDataPoint[] = [];
        if (!dataView || !dataView.categorical) return dataPoints;

        const categories = dataView.categorical.categories;
        const values = dataView.categorical.values || [];

        const measureGroup = values.find(v => v.source.roles?.measure);
        const colorGroup = values.find(v => v.source.roles?.textColor);
        const comparisonGroup = values.find(v => v.source.roles?.comparisonValue);
        const tooltipGroups = values.filter(v => v.source.roles?.tooltipFields);

        const hasCategories = categories && categories.length > 0 && categories[0];

        if (hasCategories) {
            const category = categories[0];
            const highlights = measureGroup ? measureGroup.highlights : null;

            for (let i = 0; i < category.values.length; i++) {
                const measureName = category.values[i] as string || "";
                const measureValue = (measureGroup?.values[i] as number ?? 0);
                const formatString = this.getFormatString(measureGroup, i);
                const formatter = valueFormatter.create({ format: formatString });
                const displayValue = formatter.format(measureValue);

                let textColor = "";
                if (colorGroup) {
                    const colorValue = colorGroup.values[i] as string;
                    if (typeof colorValue === "string" && /^#[0-9A-F]{6}$/i.test(colorValue)) {
                        textColor = colorValue;
                    }
                }

                const { comparisonDisplay, deltaPercent, deltaAbsolute } = this.calculateComparison(comparisonGroup, measureValue, i, settings);

                const selectionId = this.host.createSelectionIdBuilder()
                    .withCategory(category, i)
                    .createSelectionId();

                const tooltipItems: VisualTooltipDataItem[] = tooltipGroups.map(group => ({
                    displayName: group.source.displayName || '',
                    value: group.values[i] !== null ? valueFormatter.create({ format: group.source.format }).format(group.values[i]) : '',
                    color: '',
                    header: ''
                }));

                dataPoints.push({
                    measureName,
                    displayValue,
                    measureValue,
                    textColor,
                    comparisonDisplay,
                    deltaPercent,
                    deltaAbsolute,
                    formatString,
                    comparisonFormat: "",
                    selectionId,
                    categoryIndex: i,
                    highlighted: highlights ? highlights[i] !== null : true,
                    selected: false,
                    tooltipItems
                });
            }
        } else if (measureGroup) {
            const i = 0;
            const measureName = measureGroup.source.displayName || "Measure";
            const measureValue = measureGroup.values[i] as number ?? 0;
            const formatString = this.getFormatString(measureGroup, i);
            const formatter = valueFormatter.create({ format: formatString });
            const displayValue = formatter.format(measureValue);

            let textColor = "";
            if (colorGroup) {
                const colorValue = colorGroup.values[i] as string;
                if (/^#[0-9A-F]{6}$/i.test(colorValue)) {
                    textColor = colorValue;
                }
            }

            const { comparisonDisplay, deltaPercent, deltaAbsolute } = this.calculateComparison(comparisonGroup, measureValue, i, settings);

            const selectionId = this.host.createSelectionIdBuilder()
                .withMeasure(measureGroup.source.queryName)
                .createSelectionId();

            const highlights = measureGroup.highlights;
            const highlighted = highlights ? highlights[i] !== null : true;

            const tooltipItems: VisualTooltipDataItem[] = tooltipGroups.map(group => ({
                displayName: group.source.displayName || '',
                value: group.values[i] !== null ? valueFormatter.create({ format: group.source.format }).format(group.values[i]) : '',
                color: '',
                header: ''
            }));

            dataPoints.push({
                measureName,
                displayValue,
                measureValue,
                textColor,
                comparisonDisplay,
                deltaPercent,
                deltaAbsolute,
                formatString,
                comparisonFormat: "",
                selectionId,
                categoryIndex: i,
                highlighted,
                selected: false,
                tooltipItems
            });
        }

        return dataPoints;
    }

    private getFormatString(group: powerbi.DataViewValueColumn | undefined, index: number): string {
        return group ? group.source.format || (group.objects && group.objects[index] && group.objects[index].general && group.objects[index].general.formatString as string) || "" : "";
    }

    private calculateComparison(group: powerbi.DataViewValueColumn | undefined, measureValue: number, index: number, settings: VisualSettings): { comparisonDisplay: string, deltaPercent: string, deltaAbsolute: string } {
        let comparisonValue = group?.values[index] as number ?? 0;
        let comparisonDisplay = "";
        let deltaPercent = "";
        let deltaAbsolute = "";

        if (comparisonValue !== 0) {
            const comparisonFormat = group ? group.source.format || (group.objects && group.objects[index] && group.objects[index].general && group.objects[index].general.formatString as string) || "0.00" : "0.00";
            const compFormatter = valueFormatter.create({ format: comparisonFormat });
            comparisonDisplay = compFormatter.format(comparisonValue);

            const delta = ((measureValue - comparisonValue) / comparisonValue) * 100;
            deltaPercent = ` (${delta.toFixed(1)}%)`;
            deltaAbsolute = ` (${compFormatter.format(measureValue - comparisonValue)})`;
        }

        return { comparisonDisplay, deltaPercent, deltaAbsolute };
    }

    private applyDefaultSelection(dataPoints: KpiDataPoint[], settings: VisualSettings): void {
        if (settings.stateSettings.selectedDefault.value && this.isFirstUpdate && dataPoints.length > 0 && this.selectionManager.getSelectionIds().length === 0) {
            const targetInput = (settings.stateSettings.selectedDefaultMeasureName.value || "").trim().toLowerCase();

            let defaultSelectionId: ISelectionId | null = null;

            if (targetInput === "0") {
                defaultSelectionId = dataPoints[0].selectionId;
            } else if (targetInput) {
                const matchingPoint = dataPoints.find(d => d.measureName.toLowerCase() === targetInput);
                if (matchingPoint) {
                    defaultSelectionId = matchingPoint.selectionId;
                }
            }

            if (defaultSelectionId) {
                this.selectionManager.select(defaultSelectionId, false)
                    .then(() => {
                        if (this.cards) {
                            this.syncSelectionState(this.cards, settings);
                        }
                    })
                    .catch(err => console.error('Default selection error:', err));
            }
            this.isFirstUpdate = false;
        }
    }

    private calculateGridDimensions(numItems: number, vpWidth: number, vpHeight: number, settings: VisualSettings): { cardWidth: number, cardHeight: number, columns: number, rows: number } {
        const gridMargin = settings.gridSettings.margin.value || 10;
        let columns = settings.gridSettings.columns.value || 0;
        let rowsSetting = settings.gridSettings.rows.value || 0;

        let cardWidth = settings.gridSettings.fixedWidth.value;
        let cardHeight = settings.gridSettings.fixedHeight.value;
        const minWidth = settings.cardStyle.minWidth.value || 50;
        const minHeight = settings.cardStyle.minHeight.value || 50;
        const sizingMode = settings.gridSettings.sizingMode.value;

        let rows: number;
        if (sizingMode === "dynamicWidth") {
            cardHeight = Math.max(minHeight, cardHeight);
            columns = columns > 0 ? columns : Math.max(1, Math.floor(vpWidth / (cardWidth + gridMargin)));
            rows = Math.ceil(numItems / columns);
            cardWidth = Math.max(minWidth, (vpWidth - (columns - 1) * gridMargin) / columns);
        } else if (sizingMode === "dynamicHeight") {
            cardWidth = Math.max(minWidth, cardWidth);
            rows = rowsSetting > 0 ? rowsSetting : Math.max(1, Math.floor(vpHeight / (cardHeight + gridMargin)));
            columns = Math.ceil(numItems / rows);
            cardHeight = Math.max(minHeight, (vpHeight - (rows - 1) * gridMargin) / rows);
        } else {
            cardWidth = Math.max(minWidth, cardWidth);
            cardHeight = Math.max(minHeight, cardHeight);
            if (columns <= 0) {
                columns = Math.max(1, Math.floor(vpWidth / (cardWidth + gridMargin)));
            }
            rows = Math.ceil(numItems / columns);
        }

        return { cardWidth, cardHeight, columns, rows };
    }

    private renderCards(dataPoints: KpiDataPoint[], settings: VisualSettings, cardWidth: number, cardHeight: number, columns: number, rows: number, hasCategories: boolean): Selection<SVGGElement> {
        const gridMargin = settings.gridSettings.margin.value || 10;
        const startX = 0;
        const startY = 0;

        const cards = this.svg.selectAll<SVGGElement, KpiDataPoint>("g.card")
            .data(dataPoints, d => d.measureName);

        const enterCards = cards.enter()
            .append<SVGGElement>("g")
            .classed("card", true);

        enterCards.append("rect").classed("rect", true);
        enterCards.append("text").classed("measureNameText", true);
        enterCards.append("text").classed("measureValue", true);
        enterCards.append("text").classed("comparisonValue", true);

        cards.exit().remove();

        const mergedCards = enterCards.merge(cards);

        mergedCards.each((d: KpiDataPoint, index: number, nodes: Array<SVGGElement>) => {
            const group = d3.select(nodes[index]);
            group.style("opacity", d.highlighted ? 1 : 0.5);

            const col = index % columns;
            const row = Math.floor(index / columns);
            const rectX = startX + col * (cardWidth + gridMargin);
            const rectY = startY + row * (cardHeight + gridMargin);

            const rect = group.select("rect.rect")
                .attr("x", rectX)
                .attr("y", rectY)
                .attr("width", cardWidth)
                .attr("height", cardHeight)
                .attr("rx", settings.cardStyle.shapeCornerRadius.value)
                .attr("ry", settings.cardStyle.shapeCornerRadius.value)
                .attr("fill", settings.cardStyle.shapeFill.value.value)
                .attr("stroke", settings.cardStyle.shapeBorderColor.value.value)
                .attr("stroke-width", settings.cardStyle.shapeBorderWidth.value);

            const innerPadding = settings.measureName.innerPadding.value;
            const measureNameYOffset = settings.measureName.measureNameYOffset.value;
            const availableWidth = cardWidth - 2 * innerPadding;

            const measureNameFontSize = settings.measureName.measureNameFontSize.value;
            const measureNameFontFamily = settings.measureName.measureNameFontFamily.value;

            const nameText = group.select<SVGTextElement>("text.measureNameText")
                .attr("x", rectX + cardWidth / 2)
                .attr("y", rectY + measureNameYOffset + measureNameFontSize / 2)
                .attr("dy", "0.35em")
                .attr("text-anchor", "middle")
                .attr("fill", settings.measureName.measureNameColor.value.value)
                .style("font-size", `${measureNameFontSize}pt`)
                .style("font-family", measureNameFontFamily)
                .style("font-weight", settings.measureName.measureNameBold.value ? "bold" : "normal")
                .style("font-style", settings.measureName.measureNameItalic.value ? "italic" : "normal")
                .style("text-decoration", settings.measureName.measureNameUnderline.value ? "underline" : "none")
                .text(d.measureName);

            this.wrapText(nameText, availableWidth, measureNameFontSize, measureNameFontFamily);

            const valueTextColor = d.textColor || settings.measureValue.measureValueColor.value.value;

            const valueY = rectY + cardHeight / 2 + settings.measureValue.measureValueYOffset.value;
            const comparisonY = valueY + settings.measureValue.measureValueFontSize.value + settings.comparisonValue.comparisonValueYOffset.value;
            const comparisonValueDelta = settings.comparisonValue.comparisonValueDelta.value || "none";
            let comparisonDisplayFull = "";

            if (comparisonValueDelta === 'absolute') {
                comparisonDisplayFull = settings.comparisonValue.comparisonValuePrefix.value + d.comparisonDisplay + d.deltaAbsolute;
            } else if (comparisonValueDelta === 'relative') {
                comparisonDisplayFull = settings.comparisonValue.comparisonValuePrefix.value + d.comparisonDisplay + d.deltaPercent;
            } else {
                comparisonDisplayFull = settings.comparisonValue.comparisonValuePrefix.value + d.comparisonDisplay;
            }

            group.select("text.measureValue")
                .attr("x", rectX + cardWidth / 2)
                .attr("y", valueY)
                .attr("dy", "0.35em")
                .attr("text-anchor", "middle")
                .attr("fill", valueTextColor)
                .style("font-size", `${settings.measureValue.measureValueFontSize.value}pt`)
                .style("font-family", settings.measureValue.measureValueFontFamily.value)
                .style("font-weight", settings.measureValue.measureValueBold.value ? "bold" : "normal")
                .style("font-style", settings.measureValue.measureValueItalic.value ? "italic" : "normal")
                .style("text-decoration", settings.measureValue.measureValueUnderline.value ? "underline" : "none")
                .text(d.displayValue)
                .style("display", d.measureValue ? "block" : "none");

            const comparisonText = group.select("text.comparisonValue")
                .attr("x", rectX + cardWidth / 2)
                .attr("y", comparisonY)
                .attr("dy", "0.35em")
                .attr("text-anchor", "middle")
                .attr("fill", settings.comparisonValue.comparisonValueColor.value.value)
                .style("font-size", `${settings.comparisonValue.comparisonValueFontSize.value}pt`)
                .style("font-family", settings.comparisonValue.comparisonValueFontFamily.value)
                .style("font-weight", settings.comparisonValue.comparisonValueBold.value ? "bold" : "normal")
                .style("font-style", settings.comparisonValue.comparisonValueItalic.value ? "italic" : "normal")
                .style("text-decoration", settings.comparisonValue.comparisonValueUnderline.value ? "underline" : "none")
                .text(d.comparisonDisplay ? comparisonDisplayFull : "")
                .style("display", d.comparisonDisplay ? "block" : "none");

            group
                .on("mouseover", () => {
                    if (hasCategories) rect.attr("fill", settings.stateSettings.hoverFill.value.value);
                })
                .on("mouseout", () => {
                    if (hasCategories) rect.attr("fill", this.getBaseFill(d, settings));
                })
                .on("mousedown", () => {
                    if (hasCategories) rect.attr("fill", settings.stateSettings.activeFill.value.value);
                })
                .on("mouseup", () => {
                    if (hasCategories) rect.attr("fill", settings.stateSettings.hoverFill.value.value);
                })
                .on("mousemove", (event) => {
                    if (hasCategories && d.tooltipItems.length > 0) {
                        this.tooltipService.show({
                            dataItems: d.tooltipItems,
                            identities: [d.selectionId],
                            coordinates: [event.clientX, event.clientY],
                            isTouchEvent: false
                        });
                    }
                })
                .on("mouseout.tooltip", () => {
                    this.tooltipService.hide({
                        immediately: true,
                        isTouchEvent: false
                    });
                })
                .on("click", (event) => {
                    if (hasCategories) {
                        this.selectionManager.select(d.selectionId, event.ctrlKey).then(() => {
                            this.syncSelectionState(mergedCards, settings);
                        });
                    }
                    event.stopPropagation();
                });
        });

        return mergedCards;
    }

    private syncSelectionState(mergedCards: Selection<SVGGElement>, settings: VisualSettings): void {
        const ids = <ISelectionId[]>this.selectionManager.getSelectionIds();
        mergedCards.select("rect.rect").attr("fill", (d: KpiDataPoint) => {
            d.selected = ids.some(id => id.getKey() === d.selectionId.getKey());
            return this.getBaseFill(d, settings);
        });
    }

    private getBaseFill(d: KpiDataPoint, settings: VisualSettings): string {
        return d.selected ? settings.stateSettings.selectedFill.value.value : settings.cardStyle.shapeFill.value.value;
    }

    private wrapText(textSelection: Selection<SVGTextElement>, availableWidth: number, fontSize: number, fontFamily: string): void {
        textSelection.each(function () {
            const text = d3.select(this);
            const words = text.text().split(/\s+/).reverse();
            let word: string;
            let line: string[] = [];
            let lineNumber = 0;
            const lineHeight = 1.1;
            const y = text.attr("y");
            const dy = parseFloat(text.attr("dy")) || 0;
            let tspan = text.text(null).append("tspan")
                .attr("x", text.attr("x"))
                .attr("y", y)
                .attr("dy", dy + "em")
                .style("font-size", `${fontSize}pt`)
                .style("font-family", fontFamily);

            while (word = words.pop()) {
                line.push(word);
                tspan.text(line.join(" "));

                const textProperties: interfaces.TextProperties = {
                    text: tspan.text(),
                    fontFamily: fontFamily,
                    fontSize: `${fontSize}pt`,
                    fontWeight: "normal"
                };
                const measuredWidth = textMeasurementService.measureSvgTextWidth(textProperties);

                if (measuredWidth > availableWidth && line.length > 1) {
                    line.pop();
                    tspan.text(line.join(" "));
                    line = [word];
                    tspan = text.append("tspan")
                        .attr("x", text.attr("x"))
                        .attr("y", y)
                        .attr("dy", ++lineNumber * lineHeight + dy + "em")
                        .text(word)
                        .style("font-size", `${fontSize}pt`)
                        .style("font-family", fontFamily);
                }
            }
        });
    }

    public enumerateObjectInstances(options: EnumerateVisualObjectInstancesOptions): VisualObjectInstance[] | VisualObjectInstanceEnumerationObject {
        let objectName = options.objectName;
        let objectEnumeration: VisualObjectInstance[] = [];
        const settings = this.formattingSettings;

        switch (objectName) {
            case 'gridSettings':
                const properties: { [propertyName: string]: any } = {};
                properties['sizingMode'] = settings.gridSettings.sizingMode.value;
                properties['margin'] = settings.gridSettings.margin.value;

                const sizingMode = settings.gridSettings.sizingMode.value;
                if (sizingMode === 'fixed' || sizingMode === 'dynamicHeight') {
                    properties['fixedWidth'] = settings.gridSettings.fixedWidth.value;
                }
                if (sizingMode === 'fixed' || sizingMode === 'dynamicWidth') {
                    properties['fixedHeight'] = settings.gridSettings.fixedHeight.value;
                }
                if (sizingMode === 'fixed' || sizingMode === 'dynamicHeight') {
                    properties['columns'] = settings.gridSettings.columns.value;
                }
                if (sizingMode === 'dynamicWidth') {
                    properties['rows'] = settings.gridSettings.rows.value;
                }

                objectEnumeration.push({
                    objectName: objectName,
                    properties: properties,
                    selector: null
                });
                break;
        }

        return objectEnumeration;
    }

    public getFormattingModel(): powerbi.visuals.FormattingModel {

        // Conditionally toggle visibility based on the current value of the switch
        if (this.formattingSettings?.stateSettings?.selectedDefault?.value) {
            this.formattingSettings.stateSettings.selectedDefaultMeasureName.visible = true;
        } else {
            this.formattingSettings.stateSettings.selectedDefaultMeasureName.visible = false;  // Explicitly hide if false
        }

        // Conditionally toggle visibility based on the current value of the Sizing Mode
        if (this.formattingSettings?.gridSettings?.sizingMode?.value === "dynamicWidth") {
            this.formattingSettings.gridSettings.fixedWidth.visible = false;
            this.formattingSettings.gridSettings.fixedHeight.visible = true;
            this.formattingSettings.gridSettings.rows.visible = false;
            this.formattingSettings.gridSettings.columns.visible = true;

        } else if (this.formattingSettings?.gridSettings?.sizingMode?.value === "dynamicHeight") {
            this.formattingSettings.gridSettings.fixedWidth.visible = true;
            this.formattingSettings.gridSettings.fixedHeight.visible = false;
            this.formattingSettings.gridSettings.rows.visible = true;
            this.formattingSettings.gridSettings.columns.visible = false;
        } else {
            this.formattingSettings.gridSettings.fixedWidth.visible = true;
            this.formattingSettings.gridSettings.fixedHeight.visible = true;
            this.formattingSettings.gridSettings.rows.visible = false;
            this.formattingSettings.gridSettings.columns.visible = true;
        }

        //Conditionally toggle visibility for single KPI
        this.formattingSettings.gridSettings.visible = !this.isSingleKpi;
        this.formattingSettings.stateSettings.visible = !this.isSingleKpi;

        return this.formattingSettingsService.buildFormattingModel(this.formattingSettings);
    }
}