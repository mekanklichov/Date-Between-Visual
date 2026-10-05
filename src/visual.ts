/*
 * Year Range Filter
 *
 * Add ONE field to the visual's "Year" bucket:
 *
 *     Dim_Date[Year]
 *
 * The visual renders two native HTML <select> controls.
 *
 * Left:
 *     Ascending years
 *
 * Right:
 *     Descending years
 *
 * Selecting either year applies an Advanced Filter to the
 * bound Year column, allowing the filter to propagate to
 * other report visuals.
 */

import powerbi from "powerbi-visuals-api";

import IVisual = powerbi.extensibility.visual.IVisual;
import VisualConstructorOptions =
powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions =
powerbi.extensibility.visual.VisualUpdateOptions;
import IVisualHost =
powerbi.extensibility.visual.IVisualHost;
import DataView = powerbi.DataView;
import FilterAction = powerbi.FilterAction;

import {
    FormattingSettingsService
} from "powerbi-visuals-utils-formattingmodel";

import {
    YearRangeFormattingSettingsModel
} from "./formattingSettings";


export class Visual implements IVisual {

    private host!: IVisualHost;

    private root!: HTMLDivElement;
    private startSelect!: HTMLSelectElement;
    private endSelect!: HTMLSelectElement;

    private years: number[] = [];

    private startYear: number | null = null;
    private endYear: number | null = null;

    private target: {
        table: string;
        column: string;
    } | null = null;

    /*
     * Power BI formatting model
     */
    private formattingSettingsService!: FormattingSettingsService;
    private formattingSettings!: YearRangeFormattingSettingsModel;


    constructor(options?: VisualConstructorOptions) {

        if (!options) {
            return;
        }

        this.host = options.host;

        /*
         * Initialize formatting model service.
         */
        this.formattingSettingsService =
            new FormattingSettingsService();

        /*
         * Initialize default formatting settings.
         */
        this.formattingSettings =
            new YearRangeFormattingSettingsModel();


        /*
         * Root container
         */
        this.root = document.createElement("div");
        this.root.className = "year-range-root";


        /*
         * Start year
         */
        this.startSelect = document.createElement("select");
        this.startSelect.className = "year-range-select";
        this.startSelect.setAttribute(
            "aria-label",
            "Start year"
        );


        /*
         * Separator
         */
        const separator = document.createElement("span");

        separator.className = "year-range-separator";
        separator.textContent = "−";
        separator.setAttribute(
            "aria-hidden",
            "true"
        );


        /*
         * End year
         */
        this.endSelect = document.createElement("select");
        this.endSelect.className = "year-range-select";
        this.endSelect.setAttribute(
            "aria-label",
            "End year"
        );


        /*
         * Build DOM
         */
        this.root.appendChild(this.startSelect);
        this.root.appendChild(separator);
        this.root.appendChild(this.endSelect);

        options.element.appendChild(this.root);


        /*
         * Start year change
         */
        this.startSelect.addEventListener(
            "change",
            () => {

                const value = Number(
                    this.startSelect.value
                );

                this.startYear =
                    Number.isFinite(value)
                        ? value
                        : null;

                this.applyRangeFilter();
            }
        );


        /*
         * End year change
         */
        this.endSelect.addEventListener(
            "change",
            () => {

                const value = Number(
                    this.endSelect.value
                );

                this.endYear =
                    Number.isFinite(value)
                        ? value
                        : null;

                this.applyRangeFilter();
            }
        );


        /*
         * Apply initial formatting.
         */
        this.applyFormatting();
    }


    /*
     * ============================================================
     * UPDATE
     * ============================================================
     */

    public update(
        options: VisualUpdateOptions
    ): void {

        /*
         * Update formatting settings first.
         */
        if (
            options.dataViews &&
            options.dataViews.length > 0 &&
            options.dataViews[0]
        ) {
            this.formattingSettings =
                this.formattingSettingsService
                    .populateFormattingSettingsModel(
                        YearRangeFormattingSettingsModel,
                        options.dataViews[0]
                    );
        }


        /*
         * Apply formatting to HTML elements.
         */
        this.applyFormatting();


        /*
         * Get DataView.
         */
        const dataView: DataView | undefined =
            options.dataViews &&
                options.dataViews.length > 0
                ? options.dataViews[0]
                : undefined;


        /*
         * Get Year category.
         */
        const category =
            dataView?.categorical?.categories?.[0];


        /*
         * No data.
         */
        if (!category) {

            this.years = [];
            this.target = null;

            this.renderOptions();

            return;
        }


        /*
         * Extract unique years.
         */
        this.years = Array.from(category.values)
            .map(
                value => Number(value)
            )
            .filter(
                value => Number.isFinite(value)
            )
            .filter(
                (value, index, array) =>
                    array.indexOf(value) === index
            )
            .sort(
                (a, b) => a - b
            );


        /*
         * Determine the Power BI filter target.
         */
        this.target =
            this.getFilterTarget(
                category.source?.queryName
            );


        /*
         * Try to restore the currently applied
         * Advanced Filter.
         */
        const restored =
            this.readAdvancedFilter(
                options.jsonFilters
            );


        if (restored) {

            this.startYear =
                restored.start;

            this.endYear =
                restored.end;

        } else {

            /*
             * Default start year:
             * earliest available year.
             */
            if (
                this.startYear === null ||
                !this.years.includes(
                    this.startYear
                )
            ) {

                this.startYear =
                    this.years.length > 0
                        ? this.years[0]
                        : null;
            }


            /*
             * Default end year:
             * latest available year.
             */
            if (
                this.endYear === null ||
                !this.years.includes(
                    this.endYear
                )
            ) {

                this.endYear =
                    this.years.length > 0
                        ? this.years[
                        this.years.length - 1
                        ]
                        : null;
            }
        }


        /*
         * Keep range logically valid.
         */
        if (
            this.startYear !== null &&
            this.endYear !== null &&
            this.startYear > this.endYear
        ) {

            const temp =
                this.startYear;

            this.startYear =
                this.endYear;

            this.endYear =
                temp;
        }


        /*
         * Render dropdowns.
         */
        this.renderOptions();
    }


    /*
     * ============================================================
     * DROPDOWN OPTIONS
     * ============================================================
     */

    private populateSelect(
        select: HTMLSelectElement,
        years: number[],
        selectedYear?: number | null
    ): void {

        /*
         * Remove existing options.
         */
        while (select.firstChild) {
            select.removeChild(
                select.firstChild
            );
        }


        /*
         * Add exactly one option for each year.
         */
        years.forEach(
            year => {

                const option =
                    document.createElement(
                        "option"
                    );

                option.value =
                    String(year);

                option.textContent =
                    String(year);

                if (
                    selectedYear !== undefined &&
                    selectedYear !== null &&
                    year === selectedYear
                ) {

                    option.selected = true;
                }

                select.appendChild(
                    option
                );
            }
        );
    }


    private renderOptions(): void {

        /*
         * LEFT:
         * ascending
         *
         * 2018
         * 2019
         * 2020
         * ...
         */
        this.populateSelect(
            this.startSelect,
            this.years,
            this.startYear
        );


        /*
         * RIGHT:
         * descending
         *
         * 2026
         * 2025
         * 2024
         * ...
         */
        const yearsDescending =
            [...this.years].reverse();

        this.populateSelect(
            this.endSelect,
            yearsDescending,
            this.endYear
        );


        /*
         * Explicitly restore selected values.
         */
        if (this.startYear !== null) {

            this.startSelect.value =
                String(this.startYear);
        }

        if (this.endYear !== null) {

            this.endSelect.value =
                String(this.endYear);
        }
    }


    /*
     * ============================================================
     * FILTER TARGET
     * ============================================================
     */

    private getFilterTarget(
        queryName?: string
    ): {
        table: string;
        column: string;
    } | null {

        if (!queryName) {
            return null;
        }


        /*
         * Typical Power BI query name:
         *
         * Dim_Date.Year
         */
        const dot =
            queryName.lastIndexOf(".");


        if (
            dot <= 0 ||
            dot === queryName.length - 1
        ) {

            return null;
        }


        return {
            table:
                queryName.substring(
                    0,
                    dot
                ),

            column:
                queryName.substring(
                    dot + 1
                )
        };
    }


    /*
     * ============================================================
     * APPLY POWER BI FILTER
     * ============================================================
     */

    private applyRangeFilter(): void {

        if (
            !this.target ||
            this.startYear === null ||
            this.endYear === null
        ) {

            return;
        }


        /*
         * Make sure start <= end.
         */
        const start =
            Math.min(
                this.startYear,
                this.endYear
            );

        const end =
            Math.max(
                this.startYear,
                this.endYear
            );


        /*
         * Power BI Advanced Filter.
         */
        const filter = {

            $schema:
                "https://powerbi.com/product/schema#advanced",

            target:
                this.target,

            logicalOperator:
                "And",

            conditions: [

                {
                    operator:
                        "GreaterThanOrEqual",

                    value:
                        start
                },

                {
                    operator:
                        "LessThanOrEqual",

                    value:
                        end
                }
            ]
        };


        /*
         * Apply filter to the bound Year column.
         *
         * Because the filter targets Dim_Date[Year],
         * the normal Power BI model relationships will
         * propagate it to other visuals.
         */
        this.host.applyJsonFilter(
            filter,
            "general",
            "filter",
            FilterAction.merge
        );
    }


    /*
     * ============================================================
     * RESTORE FILTER
     * ============================================================
     */

    private readAdvancedFilter(
        jsonFilters:
            powerbi.IFilter[] | undefined
    ): {
        start: number;
        end: number;
    } | null {

        if (
            !jsonFilters ||
            jsonFilters.length === 0
        ) {

            return null;
        }


        /*
         * Find an Advanced Filter containing
         * at least two conditions.
         */
        const filter: any =
            jsonFilters.find(
                (item: any) =>
                    item &&
                    Array.isArray(
                        item.conditions
                    ) &&
                    item.conditions.length >= 2
            );


        if (!filter) {
            return null;
        }


        let start: number | null =
            null;

        let end: number | null =
            null;


        filter.conditions.forEach(
            (condition: any) => {

                if (!condition) {
                    return;
                }


                if (
                    condition.operator ===
                    "GreaterThanOrEqual"
                ) {

                    const value =
                        Number(
                            condition.value
                        );

                    if (
                        Number.isFinite(value)
                    ) {

                        start = value;
                    }
                }


                if (
                    condition.operator ===
                    "LessThanOrEqual"
                ) {

                    const value =
                        Number(
                            condition.value
                        );

                    if (
                        Number.isFinite(value)
                    ) {

                        end = value;
                    }
                }
            }
        );


        if (
            start === null ||
            end === null
        ) {

            return null;
        }


        return {
            start,
            end
        };
    }


    /*
     * ============================================================
     * FORMATTING MODEL
     * ============================================================
     */

    public getFormattingModel():
        powerbi.visuals.FormattingModel {

        return this.formattingSettingsService
            .buildFormattingModel(
                this.formattingSettings
            );
    }


    /*
     * ============================================================
     * APPLY FORMATTING
     * ============================================================
     */

    private applyFormatting(): void {

        /*
         * Safety check.
         */
        if (
            !this.root ||
            !this.startSelect ||
            !this.endSelect
        ) {

            return;
        }


        const settings =
            this.formattingSettings
                ?.yearRangeCard;


        if (!settings) {
            return;
        }


        /*
         * Text color.
         */
        const textColor =
            settings.textColor.value.value;


        /*
         * Font size.
         */
        const fontSize =
            settings.fontSize.value;


        /*
         * Border radius.
         */
        const borderRadius =
            settings.borderRadius.value;


        /*
         * Root layout.
         */
        this.root.style.display =
            "flex";

        this.root.style.alignItems =
            "center";


        /*
         * Apply text color.
         */
        this.startSelect.style.color =
            textColor;

        this.endSelect.style.color =
            textColor;


        /*
         * Apply font size.
         */
        this.startSelect.style.fontSize =
            `${fontSize}px`;

        this.endSelect.style.fontSize =
            `${fontSize}px`;


        /*
         * Apply border radius.
         */
        this.startSelect.style.borderRadius =
            `${borderRadius}px`;

        this.endSelect.style.borderRadius =
            `${borderRadius}px`;


        /*
         * Separator formatting.
         */
        const separator =
            this.root.querySelector(
                ".year-range-separator"
            ) as HTMLSpanElement | null;


        if (separator) {

            separator.style.color =
                textColor;

            separator.style.fontSize =
                `${fontSize}px`;

            separator.style.marginLeft =
                "6px";

            separator.style.marginRight =
                "6px";
        }
    }


    /*
     * ============================================================
     * DESTROY
     * ============================================================
     */

    public destroy(): void {

        /*
         * No external resources to dispose.
         */
    }
}