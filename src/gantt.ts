import "./../style/gantt.less";
import * as moment from 'moment';
import * as momentzone from 'moment-timezone';
import 'moment/locale/nl';
moment.locale('nl');

import { select as d3Select, Selection as d3Selection } from "d3-selection";
import { ScaleTime as timeScale } from "d3-scale";
import {
    timeDay as d3TimeDay,
    timeHour as d3TimeHour,
    timeMinute as d3TimeMinute,
    timeSecond as d3TimeSecond
} from "d3-time";
import { nest as d3Nest } from "d3-collection";
import "d3-transition";

//lodash
import lodashIsEmpty from "lodash.isempty";
import lodashMin from "lodash.min";
import lodashMinBy from "lodash.minby";
import lodashMax from "lodash.max";
import lodashMaxBy from "lodash.maxby";
import lodashGroupBy from "lodash.groupby";
import lodashClone from "lodash.clone";
import { Dictionary as lodashDictionary } from "lodash";

import powerbi from "powerbi-visuals-api";

// powerbi.extensibility.utils.svg
import * as SVGUtil from "powerbi-visuals-utils-svgutils";

// powerbi.extensibility.utils.type
import { pixelConverter as PixelConverter, valueType } from "powerbi-visuals-utils-typeutils";

// powerbi.extensibility.utils.formatting
import { textMeasurementService, valueFormatter as ValueFormatter } from "powerbi-visuals-utils-formattingutils";

// powerbi.extensibility.utils.interactivity
import {
    interactivityBaseService as interactivityService,
    interactivitySelectionService
} from "powerbi-visuals-utils-interactivityutils";

// powerbi.extensibility.utils.tooltip
import {
    createTooltipServiceWrapper,
    ITooltipServiceWrapper,
    TooltipEnabledDataPoint
} from "powerbi-visuals-utils-tooltiputils";

// powerbi.extensibility.utils.color
import { ColorHelper } from "powerbi-visuals-utils-colorutils";

// powerbi.extensibility.utils.chart.legend
import {
    axis as AxisHelper,
    axisInterfaces,
    axisScale,
    legend as LegendModule,
    legendInterfaces,
    OpacityLegendBehavior
} from "powerbi-visuals-utils-chartutils";

// behavior
import { Behavior, BehaviorOptions } from "./behavior";
import {
    DayOffData,
    DaysOffDataForAddition,
    ExtraInformation,
    GanttCalculateScaleAndDomainOptions,
    GanttChartFormatters,
    GanttViewModel,
    GroupedTask,
    Line,
    LinearStop,
    Milestone,
    MilestoneData,
    MilestoneDataPoint,
    MilestonePath,
    Task,
    TaskDaysOff,
    TaskTypeMetadata,
    TaskTypes
} from "./interfaces";
import { DurationHelper } from "./durationHelper";
import { GanttColumns } from "./columns";
import {
    drawCircle,
    drawDiamond,
    drawNotRoundedRectByPath,
    drawRectangle,
    drawRoundedRectByPath,
    hashCode,
    isStringNotNullEmptyOrUndefined,
    isValidDate
} from "./utils";
import { drawCollapseButton, drawExpandButton, } from "./drawButtons";
import { TextProperties } from "powerbi-visuals-utils-formattingutils/lib/src/interfaces";

import { FormattingSettingsService } from "powerbi-visuals-utils-formattingmodel";
import { DateTypeCardSettings, GanttChartSettingsModel } from "./ganttChartSettingsModels";
import { DateType, DurationUnit, GanttRole, LabelForDate, MilestoneShape, ResourceLabelPosition } from "./enums";

// d3
type Selection<T1, T2 = T1> = d3Selection<any, T1, any, T2>;

// powerbi
import DataView = powerbi.DataView;
import IViewport = powerbi.IViewport;
import SortDirection = powerbi.SortDirection;
import DataViewValueColumn = powerbi.DataViewValueColumn;
import DataViewValueColumns = powerbi.DataViewValueColumns;
import DataViewMetadataColumn = powerbi.DataViewMetadataColumn;
import DataViewValueColumnGroup = powerbi.DataViewValueColumnGroup;
import PrimitiveValue = powerbi.PrimitiveValue;
import DataViewCategoryColumn = powerbi.DataViewCategoryColumn;
import DataViewObjectPropertyIdentifier = powerbi.DataViewObjectPropertyIdentifier;
import VisualObjectInstancesToPersist = powerbi.VisualObjectInstancesToPersist;
import IColorPalette = powerbi.extensibility.IColorPalette;
import ILocalizationManager = powerbi.extensibility.ILocalizationManager;
import IVisualEventService = powerbi.extensibility.IVisualEventService;
import VisualTooltipDataItem = powerbi.extensibility.VisualTooltipDataItem;
// powerbi.visuals
import ISelectionIdBuilder = powerbi.visuals.ISelectionIdBuilder;
// powerbi.extensibility.visual
import IVisualHost = powerbi.extensibility.visual.IVisualHost;
import IVisual = powerbi.extensibility.visual.IVisual;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
// powerbi.extensibility.utils.svg
import SVGManipulations = SVGUtil.manipulation;
import ClassAndSelector = SVGUtil.CssConstants.ClassAndSelector;
import createClassAndSelector = SVGUtil.CssConstants.createClassAndSelector;
import IMargin = SVGUtil.IMargin;
// powerbi.extensibility.utils.type
import PrimitiveType = valueType.PrimitiveType;
import ValueType = valueType.ValueType;
// powerbi.extensibility.utils.formatting
import IValueFormatter = ValueFormatter.IValueFormatter;
// powerbi.extensibility.utils.interactivity
import appendClearCatcher = interactivityService.appendClearCatcher;
import IInteractiveBehavior = interactivityService.IInteractiveBehavior;
import IInteractivityService = interactivityService.IInteractivityService;
import createInteractivityService = interactivitySelectionService.createInteractivitySelectionService;
// powerbi.extensibility.utils.chart.legend
import ILegend = legendInterfaces.ILegend;
import LegendPosition = legendInterfaces.LegendPosition;
import LegendData = legendInterfaces.LegendData;
import createLegend = LegendModule.createLegend;
import LegendDataPoint = legendInterfaces.LegendDataPoint;
// powerbi.extensibility.utils.chart
import IAxisProperties = axisInterfaces.IAxisProperties;

const PercentFormat: string = "0.00 %;-0.00 %;0.00 %";
const ScrollMargin: number = 100;
const MillisecondsInASecond: number = 1000;
const MillisecondsInAMinute: number = 60 * MillisecondsInASecond;
const MillisecondsInAHour: number = 60 * MillisecondsInAMinute;
const MillisecondsInADay: number = 24 * MillisecondsInAHour;
const MillisecondsInWeek: number = 4 * MillisecondsInADay;
const MillisecondsInAMonth: number = 30 * MillisecondsInADay;
const MillisecondsInAYear: number = 365 * MillisecondsInADay;
const MillisecondsInAQuarter: number = MillisecondsInAYear / 4;
const PaddingTasks: number = 5;
const DaysInAWeekend: number = 2;
const DaysInAWeek: number = 5;
const DefaultChartLineHeight = 40;
const TaskColumnName: string = "Task";
const ParentColumnName: string = "Parent";
const GanttDurationUnitType = [
    DurationUnit.Second,
    DurationUnit.Minute,
    DurationUnit.Hour,
    DurationUnit.Day,
];

export class SortingOptions {
    isCustomSortingNeeded: boolean;
    sortingDirection: SortDirection;
}



export class Gantt implements IVisual {
    // Track user scroll and last scroll position
    private hasUserScrolled: boolean = false;
    private lastScrollLeft: number = 0;
    private lastScrollTop: number = 0;
    private static ClassName: ClassAndSelector = createClassAndSelector("gantt");
    private static Chart: ClassAndSelector = createClassAndSelector("chart");
    private static ChartLine: ClassAndSelector = createClassAndSelector("chart-line");
    private static Body: ClassAndSelector = createClassAndSelector("gantt-body");
    private static AxisGroup: ClassAndSelector = createClassAndSelector("axis");
    private static Domain: ClassAndSelector = createClassAndSelector("domain");
    private static AxisTick: ClassAndSelector = createClassAndSelector("tick");
    private static Tasks: ClassAndSelector = createClassAndSelector("tasks");
    private static TaskGroup: ClassAndSelector = createClassAndSelector("task-group");
    private static SingleTask: ClassAndSelector = createClassAndSelector("task");
    private static TaskRect: ClassAndSelector = createClassAndSelector("task-rect");
    private static TaskMilestone: ClassAndSelector = createClassAndSelector("task-milestone");
    private static TaskProgress: ClassAndSelector = createClassAndSelector("task-progress");
    private static TaskDaysOff: ClassAndSelector = createClassAndSelector("task-days-off");
    private static TaskResource: ClassAndSelector = createClassAndSelector("task-resource");
    private static TaskLabels: ClassAndSelector = createClassAndSelector("task-labels");
    private static TaskLines: ClassAndSelector = createClassAndSelector("task-lines");
    private static TaskLinesRect: ClassAndSelector = createClassAndSelector("task-lines-rect");
    private static TaskTopLine: ClassAndSelector = createClassAndSelector("task-top-line");
    private static CollapseAll: ClassAndSelector = createClassAndSelector("collapse-all");
    private static CollapseAllArrow: ClassAndSelector = createClassAndSelector("collapse-all-arrow");
    private static Label: ClassAndSelector = createClassAndSelector("label");
    private static LegendItems: ClassAndSelector = createClassAndSelector("legendItem");
    private static LegendTitle: ClassAndSelector = createClassAndSelector("legendTitle");
    private static ClickableArea: ClassAndSelector = createClassAndSelector("clickableArea");
    private currentGroupedTasks: GroupedTask[];

    private viewport: IViewport;
    private colors: IColorPalette;
    private colorHelper: ColorHelper;
    private legend: ILegend;
    private dailyTicks: Date[] = [];

    private textProperties: TextProperties = {
        fontFamily: "wf_segoe-ui_normal",
        fontSize: PixelConverter.toString(9),
    };

    private static LegendPropertyIdentifier: DataViewObjectPropertyIdentifier = {
        objectName: "legend",
        propertyName: "fill"
    };

    private static MilestonesPropertyIdentifier: DataViewObjectPropertyIdentifier = {
        objectName: "milestones",
        propertyName: "fill"
    };

    private static TaskResourcePropertyIdentifier: DataViewObjectPropertyIdentifier = {
        objectName: "taskResource",
        propertyName: "show"
    };

    private static CollapsedTasksPropertyIdentifier: DataViewObjectPropertyIdentifier = {
        objectName: "collapsedTasks",
        propertyName: "list"
    };

    private static CollapsedTasksUpdateIdPropertyIdentifier: DataViewObjectPropertyIdentifier = {
        objectName: "collapsedTasksUpdateId",
        propertyName: "value"
    };

    public static DefaultValues = {
        AxisTickSize: 6,
        BarMargin: 2,
        ResourceWidth: 100,
        TaskColor: "#00B099",
        TaskLineColor: "#ccc",
        CollapseAllColor: "#000",
        PlusMinusColor: "#5F6B6D",
        CollapseAllTextColor: "#aaa",
        MilestoneLineColor: "#ccc",
        TaskCategoryLabelsRectColor: "#fafafa",
        TaskLineWidth: 15,
        IconMargin: 12,
        IconHeight: 16,
        IconWidth: 15,
        ChildTaskLeftMargin: 25,
        ParentTaskLeftMargin: 0,
        DefaultDateType: DateType.Week,
        DateFormatStrings: {
            Second: "HH:mm:ss",
            Minute: "HH:mm",
            Hour: "HH:mm (dd)",
            Day: "MMM dd",
            Week: "MMM dd (Wk W)",
            Month: "MMM yyyy",
            Quarter: "MMM yyyy",
            Year: "yyyy"
        }
    };

    private static DefaultGraphicWidthPercentage: number = 0.78;
    private static ResourceLabelDefaultDivisionCoefficient: number = 1.5;
    private static DefaultTicksLength: number = 50;
    private static DefaultDuration: number = 250;
    private static TaskLineCoordinateX: number = 15;
    private static AxisLabelClip: number = 40;
    private static AxisLabelStrokeWidth: number = 1;
    private static AxisTopMargin: number = 6;
    private static CollapseAllLeftShift: number = 7.5;
    private static ResourceWidthPadding: number = 10;
    private static TaskLabelsMarginTop: number = 15;
    private static CompletionDefault: number = null;
    private static CompletionMax: number = 1;
    private static CompletionMin: number = 0;
    private static CompletionMaxInPercent: number = 100;
    private static MinTasks: number = 1;
    private static ChartLineProportion: number = 1.5;
    private static MilestoneTop: number = 0;
    private static DividerForCalculatingPadding: number = 4;
    private static LabelTopOffsetForPadding: number = 0.5;
    private static TaskBarTopPadding: number = 5;
    private static DividerForCalculatingCenter: number = 2;
    private static SubtasksLeftMargin: number = 10;
    private static NotCompletedTaskOpacity: number = .2;
    private static TaskOpacity: number = 1;
    public static RectRound: number = 7;

    private static TimeScale: timeScale<any, any>;
    private xAxisProperties: IAxisProperties;

    private static get DefaultMargin(): IMargin {
        return {
            top: 68, // Grid margin from the top
            right: 40,
            bottom: 40,
            left: 10
        };
    }

    private formattingSettings: GanttChartSettingsModel;
    private formattingSettingsService: FormattingSettingsService;

    private hasHighlights: boolean;

    private margin: IMargin = Gantt.DefaultMargin;

    private body: Selection<any>;
    private ganttSvg: Selection<any>;
    private viewModel: GanttViewModel;
    private collapseAllGroup: Selection<any>;
    private axisGroup: Selection<any>;
    private chartGroup: Selection<any>;
    private gridGroup: Selection<any>;
    private taskGroup: Selection<any>;
    private lineGroup: Selection<any>;
    private lineGroupWrapper: Selection<any>;
    private clearCatcher: Selection<any>;
    private ganttDiv: Selection<any>;
    private behavior: Behavior;
    private interactivityService: IInteractivityService<Task | LegendDataPoint>;
    private eventService: IVisualEventService;
    private tooltipServiceWrapper: ITooltipServiceWrapper;
    private host: IVisualHost;
    private localizationManager: ILocalizationManager;
    private isInteractiveChart: boolean = false;
    private groupTasksPrevValue: boolean = false;
    private collapsedTasks: string[] = [];
    private collapseAllFlag: "data-is-collapsed";
    private parentLabelOffset: number = 5;
    private groupLabelSize: number = 25;
    private secondExpandAllIconOffset: number = 7;
    private hasNotNullableDates: boolean = false;

    private collapsedTasksUpdateIDs: string[] = [];

    constructor(options: VisualConstructorOptions) {
        this.init(options);
    }

    private init(options: VisualConstructorOptions): void {
        this.host = options.host;
        this.localizationManager = this.host.createLocalizationManager();
        this.formattingSettingsService = new FormattingSettingsService(this.localizationManager);
        this.colors = options.host.colorPalette;
        this.colorHelper = new ColorHelper(this.colors);
        this.body = d3Select(options.element);
        this.tooltipServiceWrapper = createTooltipServiceWrapper(this.host.tooltipService, options.element);
        this.behavior = new Behavior();
        this.interactivityService = createInteractivityService(this.host);
        this.eventService = options.host.eventService;

        this.createViewport(options.element);
    }

    /**
     * Create the viewport area of the gantt chart
     */
    private createViewport(element: HTMLElement): void {
        const axisBackgroundColor: string = this.colorHelper.getThemeColor();
        // create div container to the whole viewport area
        this.ganttDiv = this.body.append("div")
            .classed(Gantt.Body.className, true);

        // create container to the svg area
        this.ganttSvg = this.ganttDiv
            .append("svg")
            .classed(Gantt.ClassName.className, true);

        // create clear catcher
        this.clearCatcher = appendClearCatcher(this.ganttSvg);

        // create chart container
        this.chartGroup = this.ganttSvg
            .append("g")
            .classed(Gantt.Chart.className, true);

        // create grid container
        this.gridGroup = this.chartGroup
            .append("g")
            .classed("grid-group", true);

        // create tasks container
        this.taskGroup = this.chartGroup
            .append("g")
            .classed(Gantt.Tasks.className, true);

        // create axis container
        this.axisGroup = this.ganttSvg
            .append("g")
            .classed(Gantt.AxisGroup.className, true);
        this.axisGroup
            .append("rect")
            .attr("width", "100%")
            .attr("y", "-20")
            .attr("height", "40px")
            .attr("fill", axisBackgroundColor);

        // create task lines container
        this.lineGroup = this.ganttSvg
            .append("g")
            .classed(Gantt.TaskLines.className, true);

        this.lineGroupWrapper = this.lineGroup
            .append("rect")
            .classed(Gantt.TaskLinesRect.className, true)
            // FIX: Use pixel value for height instead of "100%" to ensure gridlines render in Power BI Service
            .attr("height", this.viewport ? this.viewport.height : 500)
            .attr("width", "0")
            .attr("fill", axisBackgroundColor)
            .attr("y", this.margin.top);

        this.lineGroup
            .append("rect")
            .classed(Gantt.TaskTopLine.className, true)
            .attr("width", "100%")
            .attr("height", 1)
            .attr("y", this.margin.top)
            .attr("fill", this.colorHelper.getHighContrastColor("foreground", Gantt.DefaultValues.TaskLineColor));

        this.collapseAllGroup = this.lineGroup
            .append("g")
            .classed(Gantt.CollapseAll.className, true);

        // create legend container
        const interactiveBehavior: IInteractiveBehavior = this.colorHelper.isHighContrast ? new OpacityLegendBehavior() : null;
        this.legend = createLegend(
            element,
            this.isInteractiveChart,
            this.interactivityService,
            true,
            LegendPosition.Top,
            interactiveBehavior);

        this.ganttDiv.on("scroll", (event) => {
            if (this.viewModel) {
                // Mark that user has scrolled
                if (!this.hasUserScrolled) {
                    this.hasUserScrolled = true;
                }
                // Store last scroll positions
                this.lastScrollTop = <number>event.target.scrollTop;
                this.lastScrollLeft = <number>event.target.scrollLeft;

                const taskLabelsWidth: number = this.viewModel.settings.taskLabelsCardSettings.show.value
                    ? this.viewModel.settings.taskLabelsCardSettings.width.value
                    : 0;

                const scrollTop: number = this.lastScrollTop;
                const scrollLeft: number = this.lastScrollLeft;

                this.axisGroup
                    .attr("transform", SVGManipulations.translate(taskLabelsWidth + this.margin.left + Gantt.SubtasksLeftMargin, Gantt.TaskLabelsMarginTop + scrollTop));
                this.lineGroup
                    .attr("transform", SVGManipulations.translate(scrollLeft, 0))
                    .attr("height", 20);

                this.taskResourceRender(
                    this.taskGroup.selectAll(Gantt.SingleTask.selectorName),
                    this.viewModel.settings.taskConfigCardSettings.height.value || DefaultChartLineHeight,
                    this.currentGroupedTasks
                );
            }
        }, false);
    }

    /**
     * Clear the viewport area
     */
    private clearViewport(): void {
        this.ganttDiv
            .style("height", 0)
            .style("width", 0);

        this.body
            .selectAll(Gantt.LegendItems.selectorName)
            .remove();

        this.body
            .selectAll(Gantt.LegendTitle.selectorName)
            .remove();

        this.axisGroup
            .selectAll(Gantt.AxisTick.selectorName)
            .remove();

        this.axisGroup
            .selectAll(Gantt.Domain.selectorName)
            .remove();

        this.collapseAllGroup
            .selectAll(Gantt.CollapseAll.selectorName)
            .remove();

        this.collapseAllGroup
            .selectAll(Gantt.CollapseAllArrow.selectorName)
            .remove();

        this.lineGroup
            .selectAll(Gantt.TaskLabels.selectorName)
            .remove();

        this.lineGroup
            .selectAll(Gantt.Label.selectorName)
            .remove();

        this.chartGroup
            .selectAll(Gantt.ChartLine.selectorName)
            .remove();

        this.chartGroup
            .selectAll(Gantt.TaskGroup.selectorName)
            .remove();

        this.chartGroup
            .selectAll(Gantt.SingleTask.selectorName)
            .remove();
    }

    /**
     * Update div container size to the whole viewport area
     */
    private updateChartSize(): void {
        this.ganttDiv
            .style("height", PixelConverter.toString(this.viewport.height))
            .style("width", PixelConverter.toString(this.viewport.width));
    }

    /**
     * Check if dataView has a given role
     * @param column The dataView headers
     * @param name The role to find
     */
    private static hasRole(column: DataViewMetadataColumn, name: string) {
        return column.roles && column.roles[name];
    }

    /**
     * Get the tooltip info (data display names & formatted values)
     * @param task All task attributes.
     * @param formatters Formatting options for gantt attributes.
     * @param durationUnit Duration unit option
     * @param localizationManager powerbi localization manager
     * @param isEndDateFilled if end date is filled
     * @param roleLegendText customized legend name
     */
    public static getTooltipInfo(
        task: Task,
        formatters: GanttChartFormatters,
        durationUnit: DurationUnit,
        localizationManager: ILocalizationManager,
        isEndDateFilled: boolean,
        roleLegendText?: string): VisualTooltipDataItem[] {

        const tooltipDataArray: VisualTooltipDataItem[] = [];
        if (task.taskType) {
            tooltipDataArray.push({
                displayName: roleLegendText || localizationManager.getDisplayName("Role_Legend"),
                value: task.taskType
            });
        }

        tooltipDataArray.push({
            displayName: localizationManager.getDisplayName("Role_Task"),
            value: task.name
        });

        if (task.start && !isNaN(task.start.getDate())) {
            tooltipDataArray.push({
                displayName: localizationManager.getDisplayName("Role_StartDate"),
                value: formatters.startDateFormatter.format(task.start)
            });
        }

        if (lodashIsEmpty(task.Milestones) && task.end && !isNaN(task.end.getDate())) {
            tooltipDataArray.push({
                displayName: localizationManager.getDisplayName("Role_EndDate"),
                value: formatters.startDateFormatter.format(task.end)
            });
        }

        if (lodashIsEmpty(task.Milestones) && task.duration && !isEndDateFilled) {
            const durationLabel: string = DurationHelper.generateLabelForDuration(task.duration, durationUnit, localizationManager);
            tooltipDataArray.push({
                displayName: localizationManager.getDisplayName("Role_Duration"),
                value: durationLabel
            });
        }

        if (task.completion) {
            tooltipDataArray.push({
                displayName: localizationManager.getDisplayName("Role_Completion"),
                value: formatters.completionFormatter.format(task.completion)
            });
        }

        if (task.resource) {
            tooltipDataArray.push({
                displayName: localizationManager.getDisplayName("Role_Resource"),
                value: task.resource
            });
        }

        if (task.tooltipInfo && task.tooltipInfo.length) {
            tooltipDataArray.push(...task.tooltipInfo);
        }

        task.extraInformation
            .map(tooltip => {
                if (typeof tooltip.value === "string") {
                    return tooltip;
                }

                const value: any = tooltip.value;

                if (isNaN(Date.parse(value)) || typeof value === "number") {
                    tooltip.value = value.toString();
                } else {
                    tooltip.value = formatters.startDateFormatter.format(value);
                }

                return tooltip;
            })
            .forEach(tooltip => tooltipDataArray.push(tooltip));

        tooltipDataArray
            .filter(x => x.value && typeof x.value !== "string")
            .forEach(tooltip => tooltip.value = tooltip.value.toString());

        return tooltipDataArray;
    }

    /**
    * Check if task has data for task
    * @param dataView
    */
    private static isChartHasTask(dataView: DataView): boolean {
        if (dataView?.metadata?.columns) {
            for (const column of dataView.metadata.columns) {
                if (Gantt.hasRole(column, GanttRole.Task)) {
                    return true;
                }
            }
        }

        return false;
    }

    /**
     * Returns the chart formatters
     * @param dataView The data Model
     * @param settings visual settings
     * @param cultureSelector The current user culture
     */
    private static getFormatters(
        dataView: DataView,
        settings: GanttChartSettingsModel,
        cultureSelector: string): GanttChartFormatters {

        if (!dataView?.metadata?.columns) {
            return null;
        }

        let dateFormat: string = "d";
        for (const dvColumn of dataView.metadata.columns) {
            if (Gantt.hasRole(dvColumn, GanttRole.StartDate)) {
                dateFormat = dvColumn.format;
            }
        }

        // Priority of using date format: Format from dvColumn -> Format by culture selector -> Custom Format
        if (cultureSelector) {
            dateFormat = null;
        }

        if (!settings.tooltipConfigCardSettings.dateFormat) {
            settings.tooltipConfigCardSettings.dateFormat.value = dateFormat;
        }

        if (settings.tooltipConfigCardSettings.dateFormat &&
            settings.tooltipConfigCardSettings.dateFormat.value !== dateFormat) {

            dateFormat = settings.tooltipConfigCardSettings.dateFormat.value;
        }

        return <GanttChartFormatters>{
            startDateFormatter: ValueFormatter.create({ format: dateFormat, cultureSelector }),
            completionFormatter: ValueFormatter.create({ format: PercentFormat, value: 1, allowFormatBeautification: true })
        };
    }

    private static createLegend(
        host: IVisualHost,
        colorPalette: IColorPalette,
        settings: GanttChartSettingsModel,
        taskTypes: TaskTypes,
        useDefaultColor: boolean): LegendData {

        const colorHelper = new ColorHelper(colorPalette, Gantt.LegendPropertyIdentifier);
        const legendData: LegendData = {
            fontSize: settings.legendCardSettings.fontSize.value,
            dataPoints: [],
            title: settings.legendCardSettings.showTitle.value ? (settings.legendCardSettings.titleText.value || taskTypes?.typeName) : null,
            labelColor: settings.legendCardSettings.labelColor.value.value
        };

        legendData.dataPoints = taskTypes?.types.map(
            (typeMeta: TaskTypeMetadata): LegendDataPoint => {
                let color: string = settings.taskConfigCardSettings.fill.value.value;


                if (!useDefaultColor && !colorHelper.isHighContrast) {
                    color = colorHelper.getColorForMeasure(typeMeta.columnGroup.objects, typeMeta.name);
                }

                return {
                    label: typeMeta.name?.toString(),
                    color: color,
                    selected: false,
                    identity: host.createSelectionIdBuilder()
                        .withCategory(typeMeta.selectionColumn, 0)
                        .createSelectionId()
                };
            });

        return legendData;
    }

    private static getSortingOptions(dataView: DataView): SortingOptions {
        const sortingOption: SortingOptions = new SortingOptions();

        dataView.metadata.columns.forEach(column => {
            if (column.roles && column.sort && (column.roles[ParentColumnName] || column.roles[TaskColumnName])) {
                sortingOption.isCustomSortingNeeded = true;
                sortingOption.sortingDirection = column.sort;

                return sortingOption;
            }
        });

        return sortingOption;
    }

    private static getMinDurationUnitInMilliseconds(durationUnit: DurationUnit): number {
        switch (durationUnit) {
            case DurationUnit.Hour:
                return MillisecondsInAHour;
            case DurationUnit.Minute:
                return MillisecondsInAMinute;
            case DurationUnit.Second:
                return MillisecondsInASecond;

            default:
                return MillisecondsInADay;
        }
    }

    private static getUniqueMilestones(milestonesDataPoints: MilestoneDataPoint[]) {
        const milestonesWithoutDuplicates: {
            [name: string]: MilestoneDataPoint
        } = {};
        milestonesDataPoints.forEach((milestone: MilestoneDataPoint) => {
            if (milestone.name) {
                milestonesWithoutDuplicates[milestone.name] = milestone;
            }
        });

        return milestonesWithoutDuplicates;
    }

    private static createMilestones(
        dataView: DataView,
        host: IVisualHost): MilestoneData {
        let milestonesIndex = -1;
        for (const index in dataView.categorical.categories) {
            const category = dataView.categorical.categories[index];
            if (category.source.roles.Milestones) {
                milestonesIndex = +index;
            }
        }

        const milestoneData: MilestoneData = {
            dataPoints: []
        };
        const milestonesCategory = dataView.categorical.categories[milestonesIndex];
        const milestones: { value: PrimitiveValue, index: number }[] = [];

        if (milestonesCategory && milestonesCategory.values) {
            milestonesCategory.values.forEach((value: PrimitiveValue, index: number) => milestones.push({ value, index }));
            milestones.forEach((milestone) => {
                const milestoneObjects = milestonesCategory.objects?.[milestone.index];
                const selectionBuilder: ISelectionIdBuilder = host
                    .createSelectionIdBuilder()
                    .withCategory(milestonesCategory, milestone.index);

                const milestoneDataPoint: MilestoneDataPoint = {
                    name: milestone.value as string,
                    identity: selectionBuilder.createSelectionId(),
                    shapeType: milestoneObjects?.milestones?.shapeType ?
                        milestoneObjects.milestones.shapeType as string : MilestoneShape.Rhombus,
                    color: milestoneObjects?.milestones?.fill ?
                        (milestoneObjects.milestones as any).fill.solid.color : Gantt.DefaultValues.TaskColor
                };
                milestoneData.dataPoints.push(milestoneDataPoint);
            });
        }

        return milestoneData;
    }

    /**
     * Create task objects dataView
     * @param dataView The data Model.
     * @param formatters task attributes represented format.
     * @param taskColor Color of task
     * @param settings settings of visual
     * @param colors colors of groped tasks
     * @param host Host object
     * @param taskTypes
     * @param localizationManager powerbi localization manager
     * @param isEndDateFilled
     * @param hasHighlights if any of the tasks has highlights
     */
    private static createTasks(
        dataView: DataView,
        taskTypes: TaskTypes,
        host: IVisualHost,
        formatters: GanttChartFormatters,
        colors: IColorPalette,
        settings: GanttChartSettingsModel,
        taskColor: string,
        localizationManager: ILocalizationManager,
        isEndDateFilled: boolean,
        hasHighlights: boolean): Task[] {
        const categoricalValues: DataViewValueColumns = dataView?.categorical?.values;

        let tasks: Task[] = [];
        const addedParents: string[] = [];
        taskColor = taskColor || Gantt.DefaultValues.TaskColor;

        const values: GanttColumns<any> = GanttColumns.getCategoricalValues(dataView);

        if (!values.Task) {
            return tasks;
        }

        const colorHelper: ColorHelper = new ColorHelper(colors, Gantt.LegendPropertyIdentifier);
        const groupValues: GanttColumns<DataViewValueColumn>[] = GanttColumns.getGroupedValueColumns(dataView);
        const sortingOptions: SortingOptions = Gantt.getSortingOptions(dataView);

        const collapsedTasks: string[] = JSON.parse(settings.collapsedTasksCardSettings.list.value);
        let durationUnit: DurationUnit = <DurationUnit>settings.generalCardSettings.durationUnit.value.value.toString();
        let duration: number = settings.generalCardSettings.durationMin.value;

        let endDate: Date = null;

        values.Task.forEach((categoryValue: PrimitiveValue, index: number) => {
            const selectionBuilder: ISelectionIdBuilder = host
                .createSelectionIdBuilder()
                .withCategory(dataView.categorical.categories[0], index);

            const taskGroupAttributes = this.computeTaskGroupAttributes(taskColor, groupValues, values, index, taskTypes, selectionBuilder, colorHelper, duration, settings, durationUnit);
            const { color, completion, taskType, wasDowngradeDurationUnit, stepDurationTransformation } = taskGroupAttributes;

            duration = taskGroupAttributes.duration;
            durationUnit = taskGroupAttributes.durationUnit;
            endDate = taskGroupAttributes.endDate;

            const {
                taskParentName,
                milestone,
                startDate,
                extraInformation,
                highlight,
                task
            } = this.createTask(values, index, hasHighlights, categoricalValues, color, completion, categoryValue, endDate, duration, taskType, selectionBuilder, wasDowngradeDurationUnit, stepDurationTransformation);

            if (taskParentName) {
                Gantt.addTaskToParentTask(
                    categoryValue,
                    task,
                    tasks,
                    taskParentName,
                    addedParents,
                    collapsedTasks,
                    milestone,
                    startDate,
                    highlight,
                    extraInformation,
                    selectionBuilder,
                );
            }

            tasks.push(task);
        });

        Gantt.downgradeDurationUnitIfNeeded(tasks, durationUnit);

        if (values.Parent) {
            tasks = Gantt.sortTasksWithParents(tasks, sortingOptions);
        }

        this.updateTaskDetails(tasks, durationUnit, settings, duration, dataView, collapsedTasks);

        this.addTooltipInfoForCollapsedTasks(tasks, collapsedTasks, formatters, durationUnit, localizationManager, isEndDateFilled, settings);

        return tasks;
    }

    private static updateTaskDetails(tasks: Task[], durationUnit: DurationUnit, settings: GanttChartSettingsModel, duration: number, dataView: powerbi.DataView, collapsedTasks: string[]) {
        tasks.forEach(task => {
            if (task.children && task.children.length) {
                return;
            }

            if (task.end && task.start && isValidDate(task.end)) {
                const durationInMilliseconds: number = task.end.getTime() - task.start.getTime(),
                    minDurationUnitInMilliseconds: number = Gantt.getMinDurationUnitInMilliseconds(durationUnit);

                task.end = durationInMilliseconds < minDurationUnitInMilliseconds ? Gantt.getEndDate(durationUnit, task.start, task.duration) : task.end;
            } else {
                task.end = isValidDate(task.end) ? task.end : Gantt.getEndDate(durationUnit, task.start, task.duration);
            }

            if (settings.daysOffCardSettings.show.value && duration) {
                let datesDiff: number = 0;
                do {
                    task.daysOffList = Gantt.calculateDaysOff(
                        +settings.daysOffCardSettings.firstDayOfWeek?.value?.value,
                        new Date(task.start.getTime()),
                        new Date(task.end.getTime())
                    );

                    if (task.daysOffList.length) {
                        const isDurationFilled: boolean = dataView.metadata.columns.findIndex(col => Gantt.hasRole(col, GanttRole.Duration)) !== -1;
                        if (isDurationFilled) {
                            const extraDuration = Gantt.calculateExtraDurationDaysOff(task.daysOffList, task.start, task.end, +settings.daysOffCardSettings.firstDayOfWeek.value.value, durationUnit);
                            task.end = Gantt.getEndDate(durationUnit, task.start, task.duration + extraDuration);
                        }

                        const lastDayOffListItem = task.daysOffList[task.daysOffList.length - 1];
                        const lastDayOff: Date = lastDayOffListItem[1] === 1 ? lastDayOffListItem[0]
                            : new Date(lastDayOffListItem[0].getFullYear(), lastDayOffListItem[0].getMonth(), lastDayOffListItem[0].getDate() + 1);
                        datesDiff = Math.ceil((task.end.getTime() - lastDayOff.getTime()) / MillisecondsInADay);
                    }
                } while (task.daysOffList.length && datesDiff - DaysInAWeekend > DaysInAWeek);
            }

            if (task.parent) {
                task.visibility = collapsedTasks.indexOf(task.parent) === -1;
            }
        });
    }

    private static addTooltipInfoForCollapsedTasks(tasks: Task[], collapsedTasks: string[], formatters: GanttChartFormatters, durationUnit: DurationUnit, localizationManager: powerbi.extensibility.ILocalizationManager, isEndDateFilled: boolean, settings: GanttChartSettingsModel) {
        tasks.forEach((task: Task) => {
            if (!task.children || collapsedTasks.includes(task.name)) {
                task.tooltipInfo = Gantt.getTooltipInfo(task, formatters, durationUnit, localizationManager, isEndDateFilled, settings.legendCardSettings.titleText.value);
                if (task.Milestones) {
                    task.Milestones.forEach((milestone) => {
                        const dateFormatted = formatters.startDateFormatter.format(task.start);
                        const dateTypesSettings = settings.dateTypeCardSettings;
                        milestone.tooltipInfo = Gantt.getTooltipForMilestoneLine(dateFormatted, localizationManager, dateTypesSettings, [milestone.type], [milestone.category]);
                    });
                }
            }
        });
    }

    private static createTask(values: GanttColumns<any>, index: number, hasHighlights: boolean, categoricalValues: powerbi.DataViewValueColumns, color: string, completion: number, categoryValue: string | number | Date | boolean, endDate: Date, duration: number, taskType: TaskTypeMetadata, selectionBuilder: powerbi.visuals.ISelectionIdBuilder, wasDowngradeDurationUnit: boolean, stepDurationTransformation: number) {
        const extraInformation: ExtraInformation[] = this.getExtraInformationFromValues(values, index);

        let resource: string = (values.Resource && values.Resource[index] as string) || "";
        extraInformation.forEach(values => {
            resource += " - " + values.value;
        });

        const taskParentName: string = (values.Parent && values.Parent[index] as string) || null;
        const milestone: string = (values.Milestones && !lodashIsEmpty(values.Milestones[index]) && values.Milestones[index]) || null;

        const startDate: Date = (values.StartDate && values.StartDate[index]
            && isValidDate(new Date(values.StartDate[index])) && new Date(values.StartDate[index]))
            || new Date(Date.now());

        let highlight: number = null;
        if (hasHighlights && categoricalValues) {
            const notNullIndex = categoricalValues.findIndex(value => value.highlights && value.values[index] != null);
            if (notNullIndex != -1) highlight = <number>categoricalValues[notNullIndex].highlights[index];
        }

        const task: Task = {
            color,
            completion,
            resource,
            index: null,
            name: categoryValue as string,
            start: startDate,
            end: endDate,
            parent: taskParentName,
            children: null,
            visibility: true,
            duration,
            taskType: taskType && taskType.name,
            description: categoryValue as string,
            tooltipInfo: [],
            selected: false,
            identity: selectionBuilder.createSelectionId(),
            extraInformation,
            daysOffList: [],
            wasDowngradeDurationUnit,
            stepDurationTransformation,
            Milestones: milestone && startDate ? [{
                type: milestone,
                start: startDate,
                tooltipInfo: null,
                category: categoryValue as string
            }] : [],
            highlight: highlight !== null,
            lane: null,
            groupIndex: null
        };

        return { taskParentName, milestone, startDate, extraInformation, highlight, task };
    }

    private static computeTaskGroupAttributes(
        taskColor: string,
        groupValues: GanttColumns<powerbi.DataViewValueColumn>[],
        values: GanttColumns<any>,
        index: number,
        taskTypes: TaskTypes,
        selectionBuilder: powerbi.visuals.ISelectionIdBuilder,
        colorHelper: ColorHelper,
        duration: number,
        settings: GanttChartSettingsModel,
        durationUnit: DurationUnit) {
        let color: string = taskColor;
        let completion: number = 0;
        let taskType: TaskTypeMetadata = null;
        let wasDowngradeDurationUnit: boolean = false;
        let stepDurationTransformation: number = 0;
        let endDate: Date;

        const taskProgressShow: boolean = settings.taskCompletionCardSettings.show.value;

        if (groupValues) {
            groupValues.forEach((group: GanttColumns<DataViewValueColumn>) => {
                let maxCompletionFromTasks: number = lodashMax(values.Completion);
                maxCompletionFromTasks = maxCompletionFromTasks > Gantt.CompletionMax ? Gantt.CompletionMaxInPercent : Gantt.CompletionMax;

                if (group.Duration && group.Duration.values[index] !== null) {
                    taskType =
                        taskTypes.types.find((typeMeta: TaskTypeMetadata) => typeMeta.name === group.Duration.source.groupName);

                    if (taskType) {
                        selectionBuilder.withCategory(taskType.selectionColumn, 0);
                        color = colorHelper.getColorForMeasure(taskType.columnGroup.objects, taskType.name);
                    }

                    duration = (group.Duration.values[index] as number > settings.generalCardSettings.durationMin.value) ? group.Duration.values[index] as number : settings.generalCardSettings.durationMin.value;

                    if (duration && duration % 1 !== 0) {
                        durationUnit = DurationHelper.downgradeDurationUnit(durationUnit, duration);
                        stepDurationTransformation =
                            GanttDurationUnitType.indexOf(<DurationUnit>settings.generalCardSettings.durationUnit.value.value.toString()) - GanttDurationUnitType.indexOf(durationUnit);

                        duration = DurationHelper.transformDuration(duration, durationUnit, stepDurationTransformation);
                        wasDowngradeDurationUnit = true;
                    }

                    completion = ((group.Completion && group.Completion.values[index])
                        && taskProgressShow
                        && Gantt.convertToDecimal(group.Completion.values[index] as number, settings.taskCompletionCardSettings.maxCompletion.value, maxCompletionFromTasks)) || null;

                    if (completion !== null) {
                        if (completion < Gantt.CompletionMin) {
                            completion = Gantt.CompletionMin;
                        }

                        if (completion > Gantt.CompletionMax) {
                            completion = Gantt.CompletionMax;
                        }
                    }

                } else if (group.EndDate && group.EndDate.values[index] !== null) {
                    taskType =
                        taskTypes.types.find((typeMeta: TaskTypeMetadata) => typeMeta.name === group.EndDate.source.groupName);

                    if (taskType) {
                        selectionBuilder.withCategory(taskType.selectionColumn, 0);
                        color = colorHelper.getColorForMeasure(taskType.columnGroup.objects, taskType.name);
                    }

                    endDate = group.EndDate.values[index] ? group.EndDate.values[index] as Date : null;
                    if (typeof (endDate) === "string" || typeof (endDate) === "number") {
                        endDate = new Date(endDate);
                    }

                    completion = ((group.Completion && group.Completion.values[index])
                        && taskProgressShow
                        && Gantt.convertToDecimal(group.Completion.values[index] as number, settings.taskCompletionCardSettings.maxCompletion.value, maxCompletionFromTasks)) || null;

                    if (completion !== null) {
                        if (completion < Gantt.CompletionMin) {
                            completion = Gantt.CompletionMin;
                        }

                        if (completion > Gantt.CompletionMax) {
                            completion = Gantt.CompletionMax;
                        }
                    }
                }
            });
        }

        return {
            duration,
            durationUnit,
            color,
            completion,
            taskType,
            wasDowngradeDurationUnit,
            stepDurationTransformation,
            endDate
        };
    }

    private static addTaskToParentTask(
        categoryValue: PrimitiveValue,
        task: Task,
        tasks: Task[],
        taskParentName: string,
        addedParents: string[],
        collapsedTasks: string[],
        milestone: string,
        startDate: Date,
        highlight: number,
        extraInformation: ExtraInformation[],
        selectionBuilder: ISelectionIdBuilder,
    ) {
        if (addedParents.includes(taskParentName)) {
            const parentTask: Task = tasks.find(x => x.index === 0 && x.name === taskParentName);
            parentTask.children.push(task);
        } else {
            addedParents.push(taskParentName);

            const parentTask: Task = {
                index: 0,
                name: taskParentName,
                start: null,
                duration: null,
                completion: null,
                resource: null,
                end: null,
                parent: null,
                children: [task],
                visibility: true,
                taskType: null,
                description: null,
                color: null,
                tooltipInfo: null,
                extraInformation: collapsedTasks.includes(taskParentName) ? extraInformation : null,
                daysOffList: null,
                wasDowngradeDurationUnit: null,
                selected: null,
                identity: selectionBuilder.createSelectionId(),
                Milestones: milestone && startDate ? [{ type: milestone, start: startDate, tooltipInfo: null, category: categoryValue as string }] : [],
                highlight: highlight !== null,
                lane: null,
                groupIndex: null
            };

            tasks.push(parentTask);
        }
    }

    private static getExtraInformationFromValues(values: GanttColumns<any>, taskIndex: number): ExtraInformation[] {
        const extraInformation: ExtraInformation[] = [];

        if (values.ExtraInformation) {
            const extraInformationKeys: any[] = Object.keys(values.ExtraInformation);
            for (const key of extraInformationKeys) {
                const value: string = values.ExtraInformation[key][taskIndex];
                if (value) {
                    extraInformation.push({
                        displayName: key,
                        value: value
                    });
                }
            }
        }

        return extraInformation;
    }

    public static sortTasksWithParents(tasks: Task[], sortingOptions: SortingOptions): Task[] {
        const sortingFunction = ((a: Task, b: Task) => {
            if (a.name < b.name) {
                return sortingOptions.sortingDirection === SortDirection.Ascending ? -1 : 1;
            }

            if (a.name > b.name) {
                return sortingOptions.sortingDirection === SortDirection.Ascending ? 1 : -1;
            }

            return 0;
        });

        if (sortingOptions.isCustomSortingNeeded) {
            tasks.sort(sortingFunction);
        }

        let index: number = 0;
        tasks.forEach(task => {
            if (!task.index && !task.parent) {
                task.index = index++;

                if (task.children) {
                    if (sortingOptions.isCustomSortingNeeded) {
                        task.children.sort(sortingFunction);
                    }

                    task.children.forEach(subtask => {
                        subtask.index = subtask.index === null ? index++ : subtask.index;
                    });
                }
            }
        });

        const resultTasks: Task[] = [];

        tasks.forEach((task) => {
            resultTasks[task.index] = task;
        });

        return resultTasks;
    }

    /**
     * Calculate days off
     * @param daysOffDataForAddition Temporary days off data for addition new one
     * @param firstDayOfWeek First day of working week. From settings
     * @param date Date for verifying
     * @param extraCondition Extra condition for handle special case for last date
     */
    private static addNextDaysOff(
        daysOffDataForAddition: DaysOffDataForAddition,
        firstDayOfWeek: number,
        date: Date,
        extraCondition: boolean = false): DaysOffDataForAddition {
        daysOffDataForAddition.amountOfLastDaysOff = 1;
        for (let i = DaysInAWeekend; i > 0; i--) {
            const dateForCheck: Date = new Date(date.getTime() + (i * MillisecondsInADay));
            let alreadyInDaysOffList = false;
            daysOffDataForAddition.list.forEach((item) => {
                const itemDate = item[0];
                if (itemDate.getFullYear() === date.getFullYear() && itemDate.getMonth() === date.getMonth() && itemDate.getDate() === date.getDate()) {
                    alreadyInDaysOffList = true;
                }
            });

            const isFirstDaysOfWeek = dateForCheck.getDay() === +firstDayOfWeek;
            const isFirstDayOff = dateForCheck.getDay() === (+firstDayOfWeek + 5) % 7;
            const isSecondDayOff = dateForCheck.getDay() === (+firstDayOfWeek + 6) % 7;
            const isPartlyUsed = !/00:00:00/g.test(dateForCheck.toTimeString());

            if (!alreadyInDaysOffList && isFirstDaysOfWeek && (!extraCondition || (extraCondition && isPartlyUsed))) {
                daysOffDataForAddition.amountOfLastDaysOff = i;
                daysOffDataForAddition.list.push([
                    new Date(date.getFullYear(), date.getMonth(), date.getDate(), 0, 0, 0), i
                ]);
            }

            // Example: some task starts on Saturday 8:30 and ends on Thursday 8:30,
            // so it has extra duration and now will end on next Saturday 8:30
            // --- we need to add days off -- it ends on Monday 8.30
            if (!alreadyInDaysOffList && (isFirstDayOff || isSecondDayOff) && isPartlyUsed) {
                const amount = isFirstDayOff ? 2 : 1;
                daysOffDataForAddition.amountOfLastDaysOff = amount;
                daysOffDataForAddition.list.push([
                    new Date(dateForCheck.getFullYear(), dateForCheck.getMonth(), dateForCheck.getDate(), 0, 0, 0), amount
                ]);
            }
        }

        return daysOffDataForAddition;
    }

    /**
     * Calculates end date from start date and offset for different durationUnits
     * @param durationUnit
     * @param start  Start date
     * @param step An offset
     */
    public static getEndDate(durationUnit: DurationUnit, start: Date, step: number): Date {
        switch (durationUnit) {
            case DurationUnit.Second:
                return d3TimeSecond.offset(start, step);
            case DurationUnit.Minute:
                return d3TimeMinute.offset(start, step);
            case DurationUnit.Hour:
                return d3TimeHour.offset(start, step);
            default:
                return d3TimeDay.offset(start, step);
        }
    }


    private static isDayOff(date: Date, firstDayOfWeek: number): boolean {
        const isFirstDayOff = date.getDay() === (+firstDayOfWeek + 5) % 7;
        const isSecondDayOff = date.getDay() === (+firstDayOfWeek + 6) % 7;

        return isFirstDayOff || isSecondDayOff;
    }

    private static isOneDay(firstDate: Date, secondDate: Date): boolean {
        return firstDate.getMonth() === secondDate.getMonth() && firstDate.getFullYear() === secondDate.getFullYear()
            && firstDate.getDay() === secondDate.getDay();
    }

    /**
     * Calculate days off
     * @param firstDayOfWeek First day of working week. From settings
     * @param fromDate Start of task
     * @param toDate End of task
     */
    private static calculateDaysOff(
        firstDayOfWeek: number,
        fromDate: Date,
        toDate: Date): DayOffData[] {
        const tempDaysOffData: DaysOffDataForAddition = {
            list: [],
            amountOfLastDaysOff: 0
        };

        if (Gantt.isOneDay(fromDate, toDate)) {
            if (!Gantt.isDayOff(fromDate, +firstDayOfWeek)) {
                return tempDaysOffData.list;
            }
        }

        while (fromDate < toDate) {
            Gantt.addNextDaysOff(tempDaysOffData, firstDayOfWeek, fromDate);
            fromDate.setDate(fromDate.getDate() + tempDaysOffData.amountOfLastDaysOff);
        }

        Gantt.addNextDaysOff(tempDaysOffData, firstDayOfWeek, toDate, true);
        return tempDaysOffData.list;
    }

    private static convertMillisecondsToDuration(milliseconds: number, durationUnit: DurationUnit): number {
        switch (durationUnit) {
            case DurationUnit.Hour:
                return milliseconds /= MillisecondsInAHour;
            case DurationUnit.Minute:
                return milliseconds /= MillisecondsInAMinute;
            case DurationUnit.Second:
                return milliseconds /= MillisecondsInASecond;

            default:
                return milliseconds /= MillisecondsInADay;
        }
    }

    private static calculateExtraDurationDaysOff(daysOffList: DayOffData[], startDate: Date, endDate: Date, firstDayOfWeek: number, durationUnit: DurationUnit): number {
        let extraDuration = 0;
        for (let i = 0; i < daysOffList.length; i++) {
            const itemAmount = daysOffList[i][1];
            extraDuration += itemAmount;
            // not to count for neighbour dates
            if (itemAmount === 2 && (i + 1) < daysOffList.length) {
                const itemDate = daysOffList[i][0].getDate();
                const nextDate = daysOffList[i + 1][0].getDate();
                if (itemDate + 1 === nextDate) {
                    i += 2;
                }
            }
        }

        // not to add duration twice
        if (this.isDayOff(startDate, firstDayOfWeek)) {
            const prevDayTimestamp = startDate.getTime();
            const prevDate = new Date(prevDayTimestamp);
            prevDate.setHours(0, 0, 0);

            // in milliseconds
            let alreadyAccountedDuration = startDate.getTime() - prevDate.getTime();
            alreadyAccountedDuration = Gantt.convertMillisecondsToDuration(alreadyAccountedDuration, durationUnit);
            extraDuration = DurationHelper.transformExtraDuration(durationUnit, extraDuration);

            extraDuration -= alreadyAccountedDuration;
        }

        return extraDuration;
    }

    /**
     * Convert the dataView to view model
     * @param dataView The data Model
     * @param host Host object
     * @param colors Color palette
     * @param colorHelper powerbi color helper
     * @param localizationManager localization manager returns localized strings
     */
    public converter(
        dataView: DataView,
        host: IVisualHost,
        colors: IColorPalette,
        colorHelper: ColorHelper,
        localizationManager: ILocalizationManager): GanttViewModel {

        if (dataView?.categorical?.categories?.length === 0 || !Gantt.isChartHasTask(dataView)) {
            return null;
        }

        const settings: GanttChartSettingsModel = this.parseSettings(dataView, colorHelper);

        const taskTypes: TaskTypes = Gantt.getAllTasksTypes(dataView);

        this.hasHighlights = Gantt.hasHighlights(dataView);

        const formatters: GanttChartFormatters = Gantt.getFormatters(dataView, settings, host.locale || null);

        const isDurationFilled: boolean = dataView.metadata.columns.findIndex(col => Gantt.hasRole(col, GanttRole.Duration)) !== -1,
            isEndDateFilled: boolean = dataView.metadata.columns.findIndex(col => Gantt.hasRole(col, GanttRole.EndDate)) !== -1,
            isParentFilled: boolean = dataView.metadata.columns.findIndex(col => Gantt.hasRole(col, GanttRole.Parent)) !== -1,
            isResourcesFilled: boolean = dataView.metadata.columns.findIndex(col => Gantt.hasRole(col, GanttRole.Resource)) !== -1;

        const legendData: LegendData = Gantt.createLegend(host, colors, settings, taskTypes, !isDurationFilled && !isEndDateFilled);
        const milestonesData: MilestoneData = Gantt.createMilestones(dataView, host);

        const taskColor: string = (legendData.dataPoints?.length <= 1) || !isDurationFilled
            ? settings.taskConfigCardSettings.fill.value.value
            : null;

        const tasks: Task[] = Gantt.createTasks(dataView, taskTypes, host, formatters, colors, settings, taskColor, localizationManager, isEndDateFilled, this.hasHighlights);

        // Remove empty legend if tasks isn't exist
        const types = lodashGroupBy(tasks, x => x.taskType);
        legendData.dataPoints = legendData.dataPoints?.filter(x => types[x.label]);

        return {
            dataView,
            settings,
            taskTypes,
            tasks,
            legendData,
            milestonesData,
            isDurationFilled,
            isEndDateFilled: isEndDateFilled,
            isParentFilled,
            isResourcesFilled
        };
    }

    public parseSettings(dataView: DataView, colorHelper: ColorHelper): GanttChartSettingsModel {

        this.formattingSettings = this.formattingSettingsService.populateFormattingSettingsModel(GanttChartSettingsModel, dataView);
        const settings: GanttChartSettingsModel = this.formattingSettings;

        if (!colorHelper) {
            return settings;
        }

        if (settings.taskCompletionCardSettings.maxCompletion.value < Gantt.CompletionMin || settings.taskCompletionCardSettings.maxCompletion.value > Gantt.CompletionMaxInPercent) {
            settings.taskCompletionCardSettings.maxCompletion.value = Gantt.CompletionDefault;
        }

        if (colorHelper.isHighContrast) {
            settings.dateTypeCardSettings.axisColor.value.value = colorHelper.getHighContrastColor("foreground", settings.dateTypeCardSettings.axisColor.value.value);
            settings.dateTypeCardSettings.axisTextColor.value.value = colorHelper.getHighContrastColor("foreground", settings.dateTypeCardSettings.axisColor.value.value);
            settings.dateTypeCardSettings.todayColor.value.value = colorHelper.getHighContrastColor("foreground", settings.dateTypeCardSettings.todayColor.value.value);

            settings.daysOffCardSettings.fill.value.value = colorHelper.getHighContrastColor("foreground", settings.daysOffCardSettings.fill.value.value);
            settings.taskConfigCardSettings.fill.value.value = colorHelper.getHighContrastColor("foreground", settings.taskConfigCardSettings.fill.value.value);
            settings.taskLabelsCardSettings.fill.value.value = colorHelper.getHighContrastColor("foreground", settings.taskLabelsCardSettings.fill.value.value);
            settings.taskResourceCardSettings.fill.value.value = colorHelper.getHighContrastColor("foreground", settings.taskResourceCardSettings.fill.value.value);
            settings.legendCardSettings.labelColor.value.value = colorHelper.getHighContrastColor("foreground", settings.legendCardSettings.labelColor.value.value);
        }

        return settings;
    }

    private static convertToDecimal(value: number, maxCompletionFromSettings: number, maxCompletionFromTasks: number): number {
        if (maxCompletionFromSettings) {
            return value / maxCompletionFromSettings;
        }
        return value / maxCompletionFromTasks;
    }

    /**
    * Gets all unique types from the tasks array
    * @param dataView The data model.
    */
    private static getAllTasksTypes(dataView: DataView): TaskTypes {
        const taskTypes: TaskTypes = {
            typeName: "",
            types: []
        };
        const index: number = dataView.metadata.columns.findIndex(col => GanttRole.Legend in col.roles);

        if (index !== -1) {
            taskTypes.typeName = dataView.metadata.columns[index].displayName;
            const legendMetaCategoryColumn: DataViewMetadataColumn = dataView.metadata.columns[index];
            const values = (dataView?.categorical?.values?.length && dataView.categorical.values) || <DataViewValueColumns>[];

            if (values === undefined || values.length === 0) {
                return;
            }

            const groupValues = values.grouped();
            taskTypes.types = groupValues.map((group: DataViewValueColumnGroup): TaskTypeMetadata => {
                const column: DataViewCategoryColumn = {
                    identity: [group.identity],
                    source: {
                        displayName: null,
                        queryName: legendMetaCategoryColumn.queryName
                    },
                    values: null
                };
                return {
                    name: group.name as string,
                    selectionColumn: column,
                    columnGroup: group
                };
            });
        }

        return taskTypes;
    }

    private static hasHighlights(dataView: DataView): boolean {
        const values = (dataView?.categorical?.values?.length && dataView.categorical.values) || <DataViewValueColumns>[];
        const highlightsExist = values.some(({ highlights }) => highlights?.some(Number.isInteger));
        return !!highlightsExist;
    }

    /**
     * Get legend data, calculate position and draw it
     */
    private renderLegend(): void {
        if (!this.viewModel.legendData?.dataPoints) {
            return;
        }

        const position: string | LegendPosition = this.viewModel.settings.legendCardSettings.show.value
            ? LegendPosition[this.viewModel.settings.legendCardSettings.position.value.value]
            : LegendPosition.None;

        this.legend.changeOrientation(position as LegendPosition);
        this.legend.drawLegend(this.viewModel.legendData, lodashClone(this.viewport));
        LegendModule.positionChartArea(this.ganttDiv, this.legend);

        switch (this.legend.getOrientation()) {
            case LegendPosition.Left:
            case LegendPosition.LeftCenter:
            case LegendPosition.Right:
            case LegendPosition.RightCenter:
                this.viewport.width -= this.legend.getMargins().width;
                break;
            case LegendPosition.Top:
            case LegendPosition.TopCenter:
            case LegendPosition.Bottom:
            case LegendPosition.BottomCenter:
                this.viewport.height -= this.legend.getMargins().height;
                break;
        }
    }

    private scaleAxisLength(axisLength: number): number {
        const fullScreenAxisLength: number = Gantt.DefaultGraphicWidthPercentage * this.viewport.width;
        if (axisLength < fullScreenAxisLength) {
            axisLength = fullScreenAxisLength;
        }

        return axisLength;
    }

    /**
    * Called on data change or resizing
    * @param options The visual option that contains the dataView and the viewport
    */
    public update(options: VisualUpdateOptions): void {
        if (!options || !options.dataViews || !options.dataViews[0]) {
            this.clearViewport();
            return;
        }

        const collapsedTasksUpdateId: any = options.dataViews[0].metadata?.objects?.collapsedTasksUpdateId?.value;

        if (this.collapsedTasksUpdateIDs.includes(collapsedTasksUpdateId)) {
            this.collapsedTasksUpdateIDs = this.collapsedTasksUpdateIDs.filter(id => id !== collapsedTasksUpdateId);
            return;
        }

        this.updateInternal(options);

        // After rendering the chart, handle scroll position
        setTimeout(() => {
            if (!this.ganttDiv || !this.ganttDiv.node) return;
            const ganttDivNode = this.ganttDiv.node();
            if (!ganttDivNode) return;

            if (!this.hasUserScrolled) {
                // Scroll to today on first load
                const scrollLeftToday = this.calculateScrollLeftForToday();
                ganttDivNode.scrollLeft = scrollLeftToday;
                this.lastScrollLeft = scrollLeftToday;
                this.lastScrollTop = 0;
            } else {
                // Restore last scroll position
                ganttDivNode.scrollLeft = this.lastScrollLeft;
                ganttDivNode.scrollTop = this.lastScrollTop;
            }
        }, 0);
    }

    /**
     * Calculate the scrollLeft value needed to bring today into view
     * Returns a pixel value for the scrollLeft property
     */
    private calculateScrollLeftForToday(): number {
        // If no axis or scale, fallback to 0
        if (!this.xAxisProperties || !this.xAxisProperties.scale) return 0;
        // Find the x position of today
        const today = new Date();
        // The scale may be d3.scaleTime or similar
        let x = 0;
        try {
            x = this.xAxisProperties.scale(today);
        } catch {
            x = 0;
        }
        // Optionally, center today in the viewport
        if (this.viewport && this.viewport.width) {
            x = x - this.viewport.width / 2;
        }
        return Math.max(0, x);
    }

    private updateInternal(options: VisualUpdateOptions): void {
        this.viewModel = this.converter(options.dataViews[0], this.host, this.colors, this.colorHelper, this.localizationManager);

        // for duplicated milestone types
        if (this.viewModel && this.viewModel.milestonesData) {
            const newMilestoneData: MilestoneData = this.viewModel.milestonesData;
            const milestonesWithoutDuplicates = Gantt.getUniqueMilestones(newMilestoneData.dataPoints);

            newMilestoneData.dataPoints.forEach((dataPoint: MilestoneDataPoint) => {
                if (dataPoint.name) {
                    const theSameUniqDataPoint: MilestoneDataPoint = milestonesWithoutDuplicates[dataPoint.name];
                    dataPoint.color = theSameUniqDataPoint.color;
                    dataPoint.shapeType = theSameUniqDataPoint.shapeType;
                }
            });

            this.viewModel.milestonesData = newMilestoneData;
        }

        if (!this.viewModel || !this.viewModel.tasks || this.viewModel.tasks.length <= 0) {
            this.clearViewport();
            return;
        }

        this.viewport = lodashClone(options.viewport);
        this.margin = Gantt.DefaultMargin;

        this.eventService.renderingStarted(options);

        this.render();

        this.eventService.renderingFinished(options);
    }

    private render(): void {
        const settings = this.viewModel.settings;

        this.renderLegend();
        this.updateChartSize();

        const visibleTasks = this.viewModel.tasks
            .filter((task: Task) => task.visibility);
        const tasks: Task[] = visibleTasks
            .map((task: Task, i: number) => {
                task.index = i;
                return task;
            });

        if (this.interactivityService) {
            this.interactivityService.applySelectionStateToData(tasks);
        }

        if (tasks.length < Gantt.MinTasks) {
            return;
        }

        this.collapsedTasks = JSON.parse(settings.collapsedTasksCardSettings.list.value);
        const groupTasks = this.viewModel.settings.generalCardSettings.groupTasks.value;
        const groupedTasks: GroupedTask[] = Gantt.getGroupTasks(tasks, groupTasks, this.collapsedTasks);

        groupedTasks.forEach((group, groupIndex) => {
            group.tasks.forEach(task => {
                task.groupIndex = groupIndex;
            });
        });

        // Assign lanes for overlapping tasks in each group
        groupedTasks.forEach(group => Gantt.assignTaskLanes(group.tasks));

        groupedTasks.forEach(group => {
            group.maxLane = Gantt.getMaxLane(group.tasks);
            group.rowHeight = (group.maxLane + 1) * (this.viewModel.settings.taskConfigCardSettings.height.value || DefaultChartLineHeight);
        });

        this.updateCommonTasks(groupedTasks);
        this.updateCommonMilestones(groupedTasks);

        const tasksAfterGrouping: Task[] = groupedTasks.flatMap(t => t.tasks);
        const minDateTask: Task = lodashMinBy(tasksAfterGrouping, (t) => t && t.start);
        const maxDateTask: Task = lodashMaxBy(tasksAfterGrouping, (t) => t && t.end);
        this.hasNotNullableDates = !!minDateTask && !!maxDateTask;

        let axisLength: number = 0;
        if (this.hasNotNullableDates) {
            const startDate: Date = minDateTask.start;
            let endDate: Date = maxDateTask.end;

            if (startDate.toString() === endDate.toString()) {
                endDate = new Date(endDate.valueOf() + (24 * 60 * 60 * 1000));
            }

            const dateTypeMilliseconds: number = Gantt.getDateType(DateType[settings.dateTypeCardSettings.type.value.value]);
            let ticks: number = Math.ceil(Math.round(endDate.valueOf() - startDate.valueOf()) / dateTypeMilliseconds);
            ticks = ticks < 2 ? 2 : ticks;

            axisLength = ticks * Gantt.DefaultTicksLength;
            axisLength = this.scaleAxisLength(axisLength);

            const viewportIn: IViewport = {
                height: this.viewport.height,
                width: axisLength
            };

            const xAxisProperties: IAxisProperties = this.calculateAxes(viewportIn, this.textProperties, startDate, endDate, ticks, false);
            this.xAxisProperties = xAxisProperties;
            Gantt.TimeScale = <timeScale<Date, Date>>xAxisProperties.scale;

            this.renderAxis(xAxisProperties);
        }

        axisLength = this.scaleAxisLength(axisLength);

        // Always set SVG and grid wrapper height before drawing gridlines
        this.setDimension(groupedTasks, axisLength, settings);

        // --- Render weekend backgrounds ---
        if (this.gridGroup && this.dailyTicks && Gantt.TimeScale) {
            // Remove old weekend backgrounds
            this.gridGroup.selectAll('.weekend-background').remove();
            const chartHeight = parseFloat(this.ganttSvg.attr('height') || '0');
            const gridLineHeight = chartHeight > 0 ? chartHeight : (this.viewport ? this.viewport.height : 500);
            // Assume Sunday (0) and Saturday (6) as weekends
            this.dailyTicks.forEach((date) => {
                if (date.getDay() === 0 || date.getDay() === 6) {
                    const x = Gantt.TimeScale(date);
                    // Calculate width: distance to next day or fallback to tick width
                    const nextDay = new Date(date);
                    nextDay.setDate(date.getDate() + 1);
                    const x2 = Gantt.TimeScale(nextDay);
                    const width = (x2 && !isNaN(x2)) ? (x2 - x) : Gantt.DefaultTicksLength;
                    // Draw rect
                    this.gridGroup.append('rect')
                        .attr('class', 'weekend-background')
                        .attr('x', x)
                        .attr('y', 0)
                        .attr('width', width)
                        .attr('height', gridLineHeight)
                        .lower(); // Ensure it's behind grid lines
                }
            });
        }

        // Now render axis/gridlines (SVG height is guaranteed)
        this.renderAxis(this.xAxisProperties);
        this.renderDayLabels();

        this.renderTasks(groupedTasks);
        this.updateTaskLabels(groupedTasks, settings.taskLabelsCardSettings.width.value);
        this.updateElementsPositions(this.margin);
        this.createMilestoneLine(groupedTasks);

        if (this.formattingSettings.generalCardSettings.scrollToCurrentTime.value && this.hasNotNullableDates) {
            this.scrollToMilestoneLine(axisLength);
        }

        this.bindInteractivityService(tasks);
    }

    private bindInteractivityService(tasks: Task[]): void {
        if (this.interactivityService) {
            const behaviorOptions: BehaviorOptions = {
                clearCatcher: this.body,
                taskSelection: this.taskGroup.selectAll(Gantt.SingleTask.selectorName),
                legendSelection: this.body.selectAll(Gantt.LegendItems.selectorName),
                subTasksCollapse: {
                    selection: this.body.selectAll(Gantt.ClickableArea.selectorName),
                    callback: this.subTasksCollapseCb.bind(this)
                },
                allSubtasksCollapse: {
                    selection: this.body
                        .selectAll(Gantt.CollapseAll.selectorName),
                    callback: this.subTasksCollapseAll.bind(this)
                },
                interactivityService: this.interactivityService,
                behavior: this.behavior,
                dataPoints: tasks
            };

            this.interactivityService.bind(behaviorOptions);

            this.behavior.renderSelection(this.hasHighlights);
        }
    }

    private static getDateType(dateType: DateType): number {
        switch (dateType) {
            case DateType.Second:
                return MillisecondsInASecond;

            case DateType.Minute:
                return MillisecondsInAMinute;

            case DateType.Hour:
                return MillisecondsInAHour;

            case DateType.Day:
                return MillisecondsInADay;

            case DateType.Week:
                return MillisecondsInWeek;

            case DateType.Month:
                return MillisecondsInAMonth;

            case DateType.Quarter:
                return MillisecondsInAQuarter;

            case DateType.Year:
                return MillisecondsInAYear;

            default:
                return MillisecondsInWeek;
        }
    }

    private calculateAxes(
        viewportIn: IViewport,
        textProperties: TextProperties,
        startDate: Date,
        endDate: Date,
        ticksCount: number,
        scrollbarVisible: boolean): IAxisProperties {

        // Adjust startDate to the nearest Monday
        const dayOfWeek = startDate.getDay();
        const daysFromMonday = (dayOfWeek + 6) % 7;
        startDate = new Date(startDate.getTime() - (daysFromMonday + 7) * MillisecondsInADay); // one week earlier

        // Generate ticks for every Monday
        const ticks: Date[] = [];
        const currentTick = new Date(startDate);
        while (currentTick <= endDate) {
            ticks.push(new Date(currentTick));
            currentTick.setDate(currentTick.getDate() + 7); // Move to the next Monday
        }

        this.dailyTicks = [];
        const currentDay = new Date(startDate);
        while (currentDay <= endDate) {
            this.dailyTicks.push(new Date(currentDay));
            currentDay.setDate(currentDay.getDate() + 1);
        }

        const dataTypeDatetime: ValueType = ValueType.fromPrimitiveTypeAndCategory(PrimitiveType.Date);
        const category: DataViewMetadataColumn = {
            displayName: this.localizationManager.getDisplayName("Role_StartDate"),
            queryName: GanttRole.StartDate,
            type: dataTypeDatetime,
            index: 0
        };

        const visualOptions: GanttCalculateScaleAndDomainOptions = {
            viewport: viewportIn,
            margin: this.margin,
            forcedXDomain: [startDate, endDate],
            forceMerge: false,
            showCategoryAxisLabel: false,
            showValueAxisLabel: false,
            categoryAxisScaleType: axisScale.linear,
            valueAxisScaleType: null,
            valueAxisDisplayUnits: 0,
            categoryAxisDisplayUnits: 0,
            trimOrdinalDataOnOverflow: false,
            forcedTickCount: ticks.length
        };

        const width: number = viewportIn.width;
        const axes: IAxisProperties = this.calculateAxesProperties(viewportIn, visualOptions, category);

        // Override ticks with custom Monday ticks
        axes.axis.ticks(ticks.length).tickValues(ticks);

        axes.willLabelsFit = AxisHelper.LabelLayoutStrategy.willLabelsFit(
            axes,
            width,
            textMeasurementService.measureSvgTextWidth,
            textProperties);

        // If labels do not fit, and we are not scrolling, try word breaking
        axes.willLabelsWordBreak = (!axes.willLabelsFit && !scrollbarVisible) && AxisHelper.LabelLayoutStrategy.willLabelsWordBreak(
            axes, this.margin, width, textMeasurementService.measureSvgTextWidth,
            textMeasurementService.estimateSvgTextHeight, textMeasurementService.getTailoredTextOrDefault,
            textProperties);

        return axes;
    }

    /**
 * Assigns a 'lane' index to each task in a group so overlapping tasks are stacked.
 * Modifies tasks in-place.
 */
    private static assignTaskLanes(tasks: Task[]): void {
        // Sort by start time
        const sorted = tasks.slice().sort((a, b) => a.start.getTime() - b.start.getTime());
        const lanes: Task[][] = [];

        sorted.forEach(task => {
            let placed = false;
            for (let i = 0; i < lanes.length; i++) {
                // Check if last task in this lane ends before this task starts
                const last = lanes[i][lanes[i].length - 1];
                if (!last.end || last.end.getTime() <= task.start.getTime()) {
                    lanes[i].push(task);
                    task.lane = i;
                    placed = true;
                    break;
                }
            }
            if (!placed) {
                // New lane
                task.lane = lanes.length;
                lanes.push([task]);
            }
        });
    }

    private getGroupCumulativeY(groupedTasks: GroupedTask[], groupIndex: number): number {
        let y = 0;
        for (let i = 0; i < groupIndex; i++) {
            y += groupedTasks[i].rowHeight + this.getResourceLabelTopMargin();
        }
        return y;
    }

    private calculateAxesProperties(
        viewportIn: IViewport,
        options: GanttCalculateScaleAndDomainOptions,
        metaDataColumn: DataViewMetadataColumn): IAxisProperties {

        const dateType: DateType = DateType[this.viewModel.settings.dateTypeCardSettings.type.value.value];
        const cultureSelector: string = this.host.locale;
        const customDateFormatter = (date: Date): string => {
            if (dateType === DateType.Week || dateType === DateType.Day) {
                const weekNumber = moment(date).isoWeek(); // Get ISO week number
                return `${moment(date).format("DD MMM")} (Wk ${weekNumber})`;
            }
            return moment(date).format(Gantt.DefaultValues.DateFormatStrings[dateType]);
        };

        const xAxisDateFormatter: IValueFormatter = ValueFormatter.create({
            format: Gantt.DefaultValues.DateFormatStrings[dateType],
            cultureSelector
        });
        xAxisDateFormatter.format = customDateFormatter;

        const xAxisProperties: IAxisProperties = AxisHelper.createAxis({
            pixelSpan: viewportIn.width,
            dataDomain: options.forcedXDomain,
            metaDataColumn: metaDataColumn,
            formatString: Gantt.DefaultValues.DateFormatStrings[dateType],
            outerPadding: 5,
            isScalar: true,
            isVertical: false,
            forcedTickCount: options.forcedTickCount,
            useTickIntervalForDisplayUnits: true,
            isCategoryAxis: true,
            getValueFn: (index) => {
                return xAxisDateFormatter.format(new Date(index));
            },
            scaleType: options.categoryAxisScaleType,
            axisDisplayUnits: options.categoryAxisDisplayUnits,
        });

        xAxisProperties.axisLabel = metaDataColumn.displayName;
        return xAxisProperties;
    }

    private setDimension(
        groupedTasks: GroupedTask[],
        axisLength: number,
        settings: GanttChartSettingsModel): void {

        const fullResourceLabelMargin = groupedTasks.length * this.getResourceLabelTopMargin();
        let widthBeforeConversion = this.margin.left + settings.taskLabelsCardSettings.width.value + axisLength;

        if (settings.taskResourceCardSettings.show.value && settings.taskResourceCardSettings.position.value.value === ResourceLabelPosition.Right) {
            widthBeforeConversion += Gantt.DefaultValues.ResourceWidth;
        } else {
            widthBeforeConversion += Gantt.DefaultValues.ResourceWidth / 2;
        }

        const height = Math.round(
            Gantt.TaskBarTopPadding +
            groupedTasks.reduce((sum, g) => sum + g.rowHeight, 0) +
            this.margin.top +
            fullResourceLabelMargin
        );
        const width = Math.round(widthBeforeConversion);

        this.ganttSvg
            .attr("height", height)
            .attr("width", width);
        // Ensure the grid background/wrapper matches SVG height for gridlines
        if (this.lineGroupWrapper) {
            this.lineGroupWrapper.attr("height", height);
        }
    }

    private static getGroupTasks(tasks: Task[], groupTasks: boolean, collapsedTasks: string[]): GroupedTask[] {
        if (groupTasks) {
            const groupedTasks: lodashDictionary<Task[]> = lodashGroupBy(tasks,
                x => (x.parent ? `${x.parent}.${x.name}` : x.name));

            const result: GroupedTask[] = [];
            const taskKeys: string[] = Object.keys(groupedTasks);
            const alreadyReviewedKeys: string[] = [];

            taskKeys.forEach((key: string) => {
                const isKeyAlreadyReviewed = alreadyReviewedKeys.includes(key);
                if (!isKeyAlreadyReviewed) {
                    let name: string = key;
                    if (groupedTasks[key] && groupedTasks[key].length && groupedTasks[key][0].parent && key.indexOf(groupedTasks[key][0].parent) !== -1) {
                        name = key.substr(groupedTasks[key][0].parent.length + 1, key.length);
                    }

                    // add current task
                    const taskRecord = <GroupedTask>{
                        name,
                        tasks: groupedTasks[key]
                    };
                    result.push(taskRecord);
                    alreadyReviewedKeys.push(key);

                    // see all the children and add them
                    groupedTasks[key].forEach((task: Task) => {
                        if (task.children && !collapsedTasks.includes(task.name)) {
                            task.children.forEach((childrenTask: Task) => {
                                const childrenFullName = `${name}.${childrenTask.name}`;
                                const isChildrenKeyAlreadyReviewed = alreadyReviewedKeys.includes(childrenFullName);

                                if (!isChildrenKeyAlreadyReviewed) {
                                    const childrenRecord = <GroupedTask>{
                                        name: childrenTask.name,
                                        tasks: groupedTasks[childrenFullName]
                                    };
                                    result.push(childrenRecord);
                                    alreadyReviewedKeys.push(childrenFullName);
                                }
                            });
                        }
                    });
                }
            });

            result.forEach((x, i) => {
                x.tasks.forEach(t => t.index = i);
                x.index = i;
            });

            return result;
        }

        return tasks.map(x => <GroupedTask>{
            name: x.name,
            index: x.index,
            tasks: [x]
        });
    }

    private renderAxis(xAxisProperties: IAxisProperties, duration: number = Gantt.DefaultDuration): void {
        const axisColor: string = this.viewModel.settings.dateTypeCardSettings.axisColor.value.value;
        const axisTextColor: string = this.viewModel.settings.dateTypeCardSettings.axisTextColor.value.value;

        const xAxis = xAxisProperties.axis;
        this.axisGroup.call(xAxis.tickSizeOuter(xAxisProperties.outerPadding));

        this.axisGroup
            .transition()
            .duration(duration)
            .call(xAxis);

        this.axisGroup
            .selectAll("path")
            .style("stroke", axisColor);

        this.axisGroup
            .selectAll(".tick line")
            .style("stroke", (timestamp: number) => this.setTickColor(timestamp, axisColor));

        this.axisGroup
            .selectAll(".tick text")
            .style("fill", (timestamp: number) => this.setTickColor(timestamp, axisTextColor));

        // --- Add vertical grid lines ---
        // Remove old grid lines first
        this.gridGroup.selectAll(".vertical-grid-line").remove();

        // Get tick positions
        const tickNodes = this.axisGroup.selectAll(".tick").nodes();
        const chartHeight = parseFloat(this.ganttSvg.attr("height") || "0");
        // Fallback: if chartHeight is 0, use viewport height as a backup
        const gridLineHeight = chartHeight > 0 ? chartHeight : (this.viewport ? this.viewport.height : 500);

        tickNodes.forEach((tick: SVGGElement) => {
            const transform = tick.getAttribute("transform");
            if (transform) {
                // Extract the x position from the transform string
                const match = /translate\(([^,]+),/.exec(transform);
                if (match) {
                    const x = parseFloat(match[1]);
                    this.gridGroup.append("line")
                        .attr("class", "vertical-grid-line")
                        .attr("x1", x)
                        .attr("y1", 0)
                        .attr("x2", x)
                        .attr("y2", gridLineHeight)
                        .attr("stroke", "#e0e0e0")
                        .attr("stroke-width", 1)
                        .attr("shape-rendering", "crispEdges");
                }
            }
        });

        this.dailyTicks.forEach((date: Date) => {
            const x = Gantt.TimeScale(date);
            this.gridGroup.append("line")
                .attr("class", "vertical-grid-line daily-grid-line")
                .attr("x1", x)
                .attr("y1", 0)
                .attr("x2", x)
                .attr("y2", chartHeight)
                .attr("stroke", "#e0e0e0")
                .attr("stroke-width", 0.5)
                .attr("shape-rendering", "crispEdges");
        });
    }

    private setTickColor(
        timestamp: number,
        defaultColor: string): string {
        const tickTime = new Date(timestamp);
        const firstDayOfWeek: string = this.viewModel.settings.daysOffCardSettings.firstDayOfWeek?.value?.value.toString();
        const color: string = this.viewModel.settings.daysOffCardSettings.fill.value.value;
        if (this.viewModel.settings.daysOffCardSettings.show.value) {
            const dateForCheck: Date = new Date(tickTime.getTime());
            for (let i = 0; i <= DaysInAWeekend; i++) {
                if (dateForCheck.getDay() === +firstDayOfWeek) {
                    return !i
                        ? defaultColor
                        : color;
                }
                dateForCheck.setDate(dateForCheck.getDate() + 1);
            }
        }

        return defaultColor;
    }

    /**
    * Update task labels and add its tooltips
    * @param tasks All tasks array
    * @param width The task label width
    */
    // eslint-disable-next-line max-lines-per-function
    private updateTaskLabels(
        tasks: GroupedTask[],
        width: number
    ): void {
        let axisLabel: Selection<any>;
        const taskLabelsShow: boolean = this.viewModel.settings.taskLabelsCardSettings.show.value;
        const displayGridLines: boolean = this.viewModel.settings.generalCardSettings.displayGridLines.value;
        const taskLabelsWidth: number = this.viewModel.settings.taskLabelsCardSettings.width.value;
        const taskConfigHeight: number = this.viewModel.settings.taskConfigCardSettings.height.value || DefaultChartLineHeight;
        const categoriesAreaBackgroundColor: string = this.colorHelper.getThemeColor();
        const isHighContrast: boolean = this.colorHelper.isHighContrast;

        // Only one label/gridline per group
        const labelRows = tasks.map((group, groupIndex) => ({
            group,
            groupIndex
        }));

        this.updateCollapseAllGroup(categoriesAreaBackgroundColor, taskLabelsShow);

        if (taskLabelsShow) {
            this.lineGroupWrapper
                .attr("width", taskLabelsWidth)
                .attr("fill", isHighContrast ? categoriesAreaBackgroundColor : Gantt.DefaultValues.TaskCategoryLabelsRectColor)
                .attr("stroke", this.colorHelper.getHighContrastColor("foreground", Gantt.DefaultValues.TaskLineColor))
                .attr("stroke-width", 1);

            this.lineGroup
                .selectAll(Gantt.Label.selectorName)
                .remove();

            axisLabel = this.lineGroup
                .selectAll(Gantt.Label.selectorName)
                .data(labelRows);

            const axisLabelGroup = axisLabel
                .enter()
                .append("g")
                .merge(axisLabel);

            axisLabelGroup.classed(Gantt.Label.className, true)
                .attr("transform", (d) =>
                    SVGManipulations.translate(
                        0,
                        this.margin.top +
                        this.getGroupCumulativeY(tasks, d.groupIndex)
                    )
                );

            // Main label text
            axisLabelGroup
                .append("text")
                .attr("x", Gantt.TaskLineCoordinateX)
                .attr("y", taskConfigHeight / 2)
                .attr("stroke-width", Gantt.AxisLabelStrokeWidth)
                .attr("class", "task-labels")
                .text((d) => d.group.name)
                .call(AxisHelper.LabelLayoutStrategy.clip, width - Gantt.AxisLabelClip, textMeasurementService.svgEllipsis)
                .append("title")
                .text((d) => d.group.name);

            // Grid line for each group
            axisLabelGroup
                .append("rect")
                .attr("y", 0)
                .attr("width", () => displayGridLines ? this.viewport.width : 0)
                .attr("height", 1)
                .attr("fill", this.colorHelper.getHighContrastColor("foreground", Gantt.DefaultValues.TaskLineColor));

            axisLabel
                .exit()
                .remove();
        } else {
            this.lineGroupWrapper
                .attr("width", 0)
                .attr("fill", "transparent");

            this.lineGroup
                .selectAll(Gantt.Label.selectorName)
                .remove();
        }
    }

    private updateCollapseAllGroup(categoriesAreaBackgroundColor: string, taskLabelShow: boolean) {
        this.collapseAllGroup
            .selectAll("svg")
            .remove();

        this.collapseAllGroup
            .selectAll("rect")
            .remove();

        this.collapseAllGroup
            .selectAll("text")
            .remove();

        if (this.viewModel.isParentFilled) {
            const categoryLabelsWidth: number = this.viewModel.settings.taskLabelsCardSettings.show.value
                ? this.viewModel.settings.taskLabelsCardSettings.width.value
                : 0;

            this.collapseAllGroup
                .append("rect")
                .attr("width", categoryLabelsWidth)
                .attr("height", 2 * Gantt.TaskLabelsMarginTop)
                .attr("fill", categoriesAreaBackgroundColor);

            const expandCollapseButton = this.collapseAllGroup
                .append("svg")
                .classed(Gantt.CollapseAllArrow.className, true)
                .attr("viewBox", "0 0 48 48")
                .attr("width", this.groupLabelSize)
                .attr("height", this.groupLabelSize)
                .attr("x", Gantt.CollapseAllLeftShift + this.xAxisProperties.outerPadding || 0)
                .attr("y", this.secondExpandAllIconOffset)
                .attr(this.collapseAllFlag, (this.collapsedTasks.length ? "1" : "0"));

            expandCollapseButton
                .append("rect")
                .attr("width", this.groupLabelSize)
                .attr("height", this.groupLabelSize)
                .attr("x", 0)
                .attr("y", this.secondExpandAllIconOffset)
                .attr("fill", "transparent");

            const buttonExpandCollapseColor = this.colorHelper.getHighContrastColor("foreground", Gantt.DefaultValues.CollapseAllColor);
            if (this.collapsedTasks.length) {
                drawExpandButton(expandCollapseButton, buttonExpandCollapseColor);
            } else {
                drawCollapseButton(expandCollapseButton, buttonExpandCollapseColor);
            }

            if (taskLabelShow) {
                this.collapseAllGroup
                    .append("text")
                    .attr("x", this.secondExpandAllIconOffset + this.groupLabelSize)
                    .attr("y", this.groupLabelSize)
                    .attr("font-size", "12px")
                    .attr("fill", this.colorHelper.getHighContrastColor("foreground", Gantt.DefaultValues.CollapseAllTextColor))
                    .text(this.collapsedTasks.length ? this.localizationManager.getDisplayName("Visual_Expand_All") : this.localizationManager.getDisplayName("Visual_Collapse_All"));
            }
        }
    }

    /**
     * callback for subtasks click event
     * @param taskClicked Grouped clicked task
     */
    private subTasksCollapseCb(taskClicked: GroupedTask): void {
        const taskIsChild: boolean = taskClicked.tasks[0].parent && !taskClicked.tasks[0].children;
        const taskWithoutParentAndChildren: boolean = !taskClicked.tasks[0].parent && !taskClicked.tasks[0].children;
        if (taskIsChild || taskWithoutParentAndChildren) {
            return;
        }

        const taskClickedParent: string = taskClicked.tasks[0].parent || taskClicked.tasks[0].name;
        this.viewModel.tasks.forEach((task: Task) => {
            if (task.parent === taskClickedParent &&
                task.parent.length >= taskClickedParent.length) {
                const index: number = this.collapsedTasks.indexOf(task.parent);
                if (task.visibility) {
                    this.collapsedTasks.push(task.parent);
                } else {
                    if (taskClickedParent === task.parent) {
                        this.collapsedTasks.splice(index, 1);
                    }
                }
            }
        });

        // eslint-disable-next-line
        const newId = crypto?.randomUUID() || Math.random().toString();
        this.collapsedTasksUpdateIDs.push(newId);

        this.setJsonFiltersValues(this.collapsedTasks, newId);
    }

    /**
     * callback for subtasks collapse all click event
     */
    private subTasksCollapseAll(): void {
        const collapsedAllSelector = this.collapseAllGroup.select(Gantt.CollapseAllArrow.selectorName);
        const isCollapsed: string = collapsedAllSelector.attr(this.collapseAllFlag);
        const buttonExpandCollapseColor = this.colorHelper.getHighContrastColor("foreground", Gantt.DefaultValues.CollapseAllColor);

        collapsedAllSelector.selectAll("path").remove();
        if (isCollapsed === "1") {
            this.collapsedTasks = [];
            collapsedAllSelector.attr(this.collapseAllFlag, "0");
            drawCollapseButton(collapsedAllSelector, buttonExpandCollapseColor);

        } else {
            collapsedAllSelector.attr(this.collapseAllFlag, "1");
            drawExpandButton(collapsedAllSelector, buttonExpandCollapseColor);
            this.viewModel.tasks.forEach((task: Task) => {
                if (task.parent) {
                    if (task.visibility) {
                        this.collapsedTasks.push(task.parent);
                    }
                }
            });
        }

        // eslint-disable-next-line
        const newId = crypto?.randomUUID() || Math.random().toString();
        this.collapsedTasksUpdateIDs.push(newId);

        this.setJsonFiltersValues(this.collapsedTasks, newId);
    }

    private setJsonFiltersValues(collapsedValues: string[], collapsedTasksUpdateId: string) {
        this.host.persistProperties(<VisualObjectInstancesToPersist>{
            merge: [{
                objectName: "collapsedTasks",
                selector: null,
                properties: {
                    list: JSON.stringify(collapsedValues)
                }
            }, {
                objectName: "collapsedTasksUpdateId",
                selector: null,
                properties: {
                    value: JSON.stringify(collapsedTasksUpdateId)
                }
            }]
        });
    }

    /**
     * Render tasks
     * @param groupedTasks Grouped tasks
     */
    private renderTasks(groupedTasks: GroupedTask[]): void {
        const taskConfigHeight: number = this.viewModel.settings.taskConfigCardSettings.height.value || DefaultChartLineHeight;
        const generalBarsRoundedCorners: boolean = this.viewModel.settings.generalCardSettings.barsRoundedCorners.value;
        const taskGroupSelection: Selection<any> = this.taskGroup
            .selectAll(Gantt.TaskGroup.selectorName)
            .data(groupedTasks);

        this.currentGroupedTasks = groupedTasks;

        taskGroupSelection
            .exit()
            .remove();

        // render task group container
        const taskGroupSelectionMerged = taskGroupSelection
            .enter()
            .append("g")
            .merge(taskGroupSelection);

        taskGroupSelectionMerged.classed(Gantt.TaskGroup.className, true);

        const taskSelection: Selection<Task> = this.taskSelectionRectRender(taskGroupSelectionMerged);
        this.taskMainRectRender(taskSelection, taskConfigHeight, generalBarsRoundedCorners);
        this.MilestonesRender(taskSelection, taskConfigHeight, groupedTasks);
        this.taskProgressRender(taskSelection);
        this.taskDaysOffRender(taskSelection, taskConfigHeight, groupedTasks);
        this.taskResourceRender(taskSelection, taskConfigHeight, groupedTasks);

        this.renderTooltip(taskSelection);
    }


    /**
     * Change task structure to be able for
     * Rendering common tasks when all the children of current parent are collapsed
     * used only the Grouping mode is OFF
     * @param groupedTasks Grouped tasks
     */
    private updateCommonTasks(groupedTasks: GroupedTask[]): void {
        if (!this.viewModel.settings.generalCardSettings.groupTasks.value) {
            groupedTasks.forEach((groupedTask: GroupedTask) => {
                const currentTaskName: string = groupedTask.name;
                if (this.collapsedTasks.includes(currentTaskName)) {
                    const firstTask: Task = groupedTask.tasks && groupedTask.tasks[0];
                    const tasks = groupedTask.tasks;
                    tasks.forEach((task: Task) => {
                        if (task.children) {
                            const childrenColors = task.children.map((child: Task) => child.color).filter((color) => color);
                            const minChildDateStart = lodashMin(task.children.map((child: Task) => child.start).filter((dateStart) => dateStart));
                            const maxChildDateEnd = lodashMax(task.children.map((child: Task) => child.end).filter((dateStart) => dateStart));
                            firstTask.color = !firstTask.color && task.children ? childrenColors[0] : firstTask.color;
                            firstTask.start = lodashMin([firstTask.start, minChildDateStart]);
                            firstTask.end = <any>lodashMax([firstTask.end, maxChildDateEnd]);
                        }
                    });

                    groupedTask.tasks = firstTask && [firstTask] || [];
                }
            });
        }
    }

    /**
     * Change task structure to be able for
     * Rendering common milestone when all the children of current parent are collapsed
     * used only the Grouping mode is OFF
     * @param groupedTasks Grouped tasks
     */
    private updateCommonMilestones(groupedTasks: GroupedTask[]): void {
        groupedTasks.forEach((groupedTask: GroupedTask) => {
            const currentTaskName: string = groupedTask.name;
            if (this.collapsedTasks.includes(currentTaskName)) {

                const lastTask: Task = groupedTask.tasks && groupedTask.tasks[groupedTask.tasks.length - 1];
                const tasks = groupedTask.tasks;
                tasks.forEach((task: Task) => {
                    if (task.children) {
                        task.children.map((child: Task) => {
                            if (!lodashIsEmpty(child.Milestones)) {
                                lastTask.Milestones = lastTask.Milestones.concat(child.Milestones);
                            }
                        });
                    }
                });
            }
        });
    }

    /**
     * Render task progress rect
     * @param taskGroupSelection Task Group Selection
     */
    private taskSelectionRectRender(taskGroupSelection: Selection<any>) {
        const taskSelection: Selection<Task> = taskGroupSelection
            .selectAll(Gantt.SingleTask.selectorName)
            .data((d: GroupedTask) => d.tasks);

        taskSelection
            .exit()
            .remove();

        const taskSelectionMerged = taskSelection
            .enter()
            .append("g")
            .merge(taskSelection);

        taskSelectionMerged.classed(Gantt.SingleTask.className, true);

        return taskSelectionMerged;
    }

    /**
     * @param task
     */
    private getTaskRectWidth(task: Task): number {
        const taskIsCollapsed = this.collapsedTasks.includes(task.name);

        if (this.hasNotNullableDates && (taskIsCollapsed || lodashIsEmpty(task.Milestones))) {
            const { start, end } = Gantt.normalizeTaskDates(task);
            return Gantt.taskDurationToWidth(start, end);
        }
        return 0;
    }

    /**
     *
     * @param task
     * @param taskConfigHeight
     * @param barsRoundedCorners are bars with rounded corners
     */
    private drawTaskRect(task: Task, taskConfigHeight: number, barsRoundedCorners: boolean, groupIndex: number, groupedTasks: GroupedTask[]): string {
        const { start, end } = Gantt.normalizeTaskDates(task);
        const x = this.hasNotNullableDates ? Gantt.TimeScale(start) : 0;
        const lane = task.lane || 0;
        const y = Gantt.TaskBarTopPadding + this.getGroupCumulativeY(groupedTasks, groupIndex)
            + lane * (Gantt.getBarHeight(taskConfigHeight) + 2);
        const width = Gantt.taskDurationToWidth(start, end);
        const height = Gantt.getBarHeight(taskConfigHeight);
        const radius = Gantt.RectRound;

        if (barsRoundedCorners && width >= 2 * radius) {
            return drawRoundedRectByPath(x, y, width, height, radius);
        }
        return drawNotRoundedRectByPath(x, y, width, height);
    }

    /**
     * Render task progress rect
     * @param taskSelection Task Selection
     * @param taskConfigHeight Task heights from settings
     * @param barsRoundedCorners are bars with rounded corners
     */
    private taskMainRectRender(
        taskSelection: Selection<Task>,
        taskConfigHeight: number,
        barsRoundedCorners: boolean): void {
        const highContrastModeTaskRectStroke: number = 1;

        const taskRect: Selection<Task> = taskSelection
            .selectAll(Gantt.TaskRect.selectorName)
            .data((d: Task) => [d]);

        const taskRectMerged = taskRect
            .enter()
            .append("path")
            .merge(taskRect);

        taskRectMerged.classed(Gantt.TaskRect.className, true);

        let index = 0, groupedTaskIndex = 0;
        taskRectMerged
            .attr("d", (task: Task) =>
                this.drawTaskRect(task, taskConfigHeight, barsRoundedCorners, task.groupIndex, this.currentGroupedTasks))
            .attr("width", (task: Task) => this.getTaskRectWidth(task))
            .style("fill", (task: Task) => {
                // logic used for grouped tasks, when there are several bars related to one category
                if (index === task.index) {
                    groupedTaskIndex++;
                } else {
                    groupedTaskIndex = 0;
                    index = task.index;
                }

                const url = `${task.index}-${groupedTaskIndex}-${isStringNotNullEmptyOrUndefined(task.taskType) ? task.taskType.toString() : "taskType"}`;
                const encodedUrl = `task${hashCode(url)}`;

                return `url(#${encodedUrl})`;
            });

        if (this.colorHelper.isHighContrast) {
            taskRectMerged
                .style("stroke", (task: Task) => this.colorHelper.getHighContrastColor("foreground", task.color))
                .style("stroke-width", highContrastModeTaskRectStroke);
        }

        taskRect
            .exit()
            .remove();
    }

    /**
     *
     * @param milestoneType milestone type
     */
    private getMilestoneColor(milestoneType: string): string {
        const milestone: MilestoneDataPoint = this.viewModel.milestonesData.dataPoints.filter((dataPoint: MilestoneDataPoint) => dataPoint.name === milestoneType)[0];

        return this.colorHelper.getHighContrastColor("foreground", milestone.color);
    }

    private getMilestonePath(milestoneType: string, taskConfigHeight: number): string {
        let shape: string;
        const convertedHeight: number = Gantt.getBarHeight(taskConfigHeight);
        const milestone: MilestoneDataPoint = this.viewModel.milestonesData.dataPoints.filter((dataPoint: MilestoneDataPoint) => dataPoint.name === milestoneType)[0];
        switch (milestone.shapeType) {
            case MilestoneShape.Rhombus:
                shape = drawDiamond(convertedHeight);
                break;
            case MilestoneShape.Square:
                shape = drawRectangle(convertedHeight);
                break;
            case MilestoneShape.Circle:
                shape = drawCircle(convertedHeight);
        }

        return shape;
    }

    /**
     * Render milestones
     * @param taskSelection Task Selection
     * @param taskConfigHeight Task heights from settings
     */
    private MilestonesRender(taskSelection: Selection<Task>, taskConfigHeight: number, groupedTasks: GroupedTask[]): void {
        const taskMilestones: Selection<any> = taskSelection
            .selectAll(Gantt.TaskMilestone.selectorName)
            .data((d: Task) => {
                const nestedByDate = d3Nest().key((d: Milestone) => d.start.toDateString()).entries(d.Milestones);
                const updatedMilestones: MilestonePath[] = nestedByDate.map((nestedObj) => {
                    const oneDateMilestones = nestedObj.values;
                    // if there is 2 or more milestones for concrete date => draw only one milestone for concrete date, but with tooltip for all of them
                    const currentMilestone = [...oneDateMilestones].pop();
                    const allTooltipInfo = oneDateMilestones.map((milestone: MilestonePath) => milestone.tooltipInfo);
                    currentMilestone.tooltipInfo = allTooltipInfo.reduce((a, b) => a.concat(b), []);

                    return {
                        type: currentMilestone.type,
                        start: currentMilestone.start,
                        taskID: d.index,
                        tooltipInfo: currentMilestone.tooltipInfo
                    };
                });

                return [{
                    key: d.index, values: <MilestonePath[]>updatedMilestones
                }];
            });


        taskMilestones
            .exit()
            .remove();

        const taskMilestonesAppend = taskMilestones
            .enter()
            .append("g");

        const taskMilestonesMerged = taskMilestonesAppend
            .merge(taskMilestones);

        taskMilestonesMerged.classed(Gantt.TaskMilestone.className, true);

        const transformForMilestone = (groupIndex: number, groupedTasks: GroupedTask[], lane: number, start: Date, taskConfigHeight: number) => {
            return SVGManipulations.translate(
                Gantt.TimeScale(start) - Gantt.getBarHeight(taskConfigHeight) / 4,
                Gantt.TaskBarTopPadding + this.getGroupCumulativeY(groupedTasks, groupIndex) + lane * (Gantt.getBarHeight(taskConfigHeight) + 2)
            );
        };
        const taskMilestonesSelection = taskMilestonesMerged.selectAll("path");
        const taskMilestonesSelectionData = taskMilestonesSelection.data(milestonesData => <MilestonePath[]>milestonesData.values);

        // add milestones: for collapsed task may be several milestones of its children, in usual case - just 1 milestone
        const taskMilestonesSelectionAppend = taskMilestonesSelectionData.enter()
            .append("path");

        taskMilestonesSelectionData
            .exit()
            .remove();

        const taskMilestonesSelectionMerged = taskMilestonesSelectionAppend
            .merge(<any>taskMilestonesSelection);

        if (this.hasNotNullableDates) {
            taskMilestonesSelectionMerged
                .attr("d", (data: MilestonePath) => this.getMilestonePath(data.type, taskConfigHeight))
                .attr("transform", (data: MilestonePath) => {
                    // Use groupIndex directly from the task
                    const groupIndex = (() => {
                        // Find the task by index in groupedTasks
                        for (let i = 0; i < groupedTasks.length; i++) {
                            if (groupedTasks[i].tasks.some(t => t.index === data.taskID)) {
                                const task = groupedTasks[i].tasks.find(t => t.index === data.taskID);
                                return task.groupIndex;
                            }
                        }
                        return 0;
                    })();
                    const group = groupedTasks[groupIndex];
                    const task = group.tasks.find(t => t.index === data.taskID);
                    return transformForMilestone(groupIndex, groupedTasks, task.lane || 0, data.start, taskConfigHeight);
                })
                .attr("fill", (data: MilestonePath) => this.getMilestoneColor(data.type));
        }

        this.renderTooltip(taskMilestonesSelectionMerged);
    }

    /**
     * Render days off rects
     * @param taskSelection Task Selection
     * @param taskConfigHeight Task heights from settings
     */
    private taskDaysOffRender(taskSelection: Selection<Task>, taskConfigHeight: number, groupedTasks: GroupedTask[]): void {
        const taskDaysOffColor: string = this.viewModel.settings.daysOffCardSettings.fill.value.value;
        const taskDaysOffShow: boolean = this.viewModel.settings.daysOffCardSettings.show.value;

        taskSelection
            .selectAll(Gantt.TaskDaysOff.selectorName)
            .remove();

        if (taskDaysOffShow) {
            const tasksDaysOff: Selection<TaskDaysOff, Task> = taskSelection
                .selectAll(Gantt.TaskDaysOff.selectorName)
                .data((d: Task) => {
                    const tasksDaysOff: TaskDaysOff[] = [];

                    if (!d.children && d.daysOffList) {
                        for (let i = 0; i < d.daysOffList.length; i++) {
                            const currentDaysOffItem: DayOffData = d.daysOffList[i];
                            const startOfLastDay: Date = new Date(+d.end);
                            startOfLastDay.setHours(0, 0, 0);
                            if (currentDaysOffItem[0].getTime() < startOfLastDay.getTime()) {
                                tasksDaysOff.push({
                                    id: d.index,
                                    daysOff: d.daysOffList[i]
                                });
                            }
                        }
                    }

                    return tasksDaysOff;
                });

            const tasksDaysOffMerged = tasksDaysOff
                .enter()
                .append("path")
                .merge(tasksDaysOff);

            tasksDaysOffMerged.classed(Gantt.TaskDaysOff.className, true);

            const getTaskRectDaysOffWidth = (task: TaskDaysOff) => {
                let width = 0;

                if (this.hasNotNullableDates) {
                    const startDate: Date = task.daysOff[0];
                    const startTime: number = startDate.getTime();
                    const endDate: Date = new Date(startTime + (task.daysOff[1] * MillisecondsInADay));

                    width = Gantt.taskDurationToWidth(startDate, endDate);
                }

                return width;
            };

            const drawTaskRectDaysOff = (task: TaskDaysOff, groupIndex: number, groupedTasks: GroupedTask[]) => {
                let x = this.hasNotNullableDates ? Gantt.TimeScale(task.daysOff[0]) : 0;
                // Find the task to get its lane and groupIndex
                let lane = 0;
                let actualGroupIndex = groupIndex;
                for (let i = 0; i < groupedTasks.length; i++) {
                    const found = groupedTasks[i].tasks.find(t => t.index === task.id);
                    if (found) {
                        lane = found.lane || 0;
                        actualGroupIndex = found.groupIndex;
                        break;
                    }
                }
                const y = this.getGroupCumulativeY(groupedTasks, actualGroupIndex)
                    + lane * (Gantt.getBarHeight(taskConfigHeight) + 2) + Gantt.TaskBarTopPadding;
                const height = Gantt.getBarHeight(taskConfigHeight);
                const radius = this.viewModel.settings.generalCardSettings.barsRoundedCorners.value ? Gantt.RectRound : 0;
                const width = getTaskRectDaysOffWidth(task);

                if (width < radius) {
                    x = x - width / 2;
                }

                if (this.formattingSettings.generalCardSettings.barsRoundedCorners.value && width >= 2 * radius) {
                    return drawRoundedRectByPath(x, y, width, height, radius);
                }

                return drawNotRoundedRectByPath(x, y, width, height);
            };

            tasksDaysOffMerged
                .attr("d", (task: TaskDaysOff) => {
                    // Use groupIndex from the task
                    let groupIndex = 0;
                    for (let i = 0; i < groupedTasks.length; i++) {
                        if (groupedTasks[i].tasks.some(t => t.index === task.id)) {
                            const found = groupedTasks[i].tasks.find(t => t.index === task.id);
                            groupIndex = found.groupIndex;
                            break;
                        }
                    }
                    return drawTaskRectDaysOff(task, groupIndex, groupedTasks);
                })
                .style("fill", taskDaysOffColor)
                .attr("width", (task: TaskDaysOff) => getTaskRectDaysOffWidth(task));

            tasksDaysOff
                .exit()
                .remove();
        }
    }

    private renderDayLabels(): void {
        if (!this.dailyTicks || !Gantt.TimeScale) return;

        // Remove old day labels and capacity labels
        this.axisGroup.selectAll(".day-label").remove();
        this.axisGroup.selectAll(".capacity-label").remove();

        // Calculate Y positions (margins) task labels
        const weekAxisHeight = 24;
        const capacityLabelYOffset = weekAxisHeight + 12; // above day label
        const dayLabelYOffset = weekAxisHeight + 24;

        // Gather all unique resources
        const allTasks = this.viewModel.tasks || [];
        const allResourcesSet = new Set<string>();
        allTasks.forEach(task => {
            if (task.name) allResourcesSet.add(task.name);
        });
        const totalResources = allResourcesSet.size;

        // For each day, calculate occupied resources
        const capacityData = this.dailyTicks.map((date) => {
            const occupiedResources = new Set<string>();
            allTasks.forEach(task => {
                if (!task.name) return;
                // Check if task is active on this day
                const { start, end } = Gantt.normalizeTaskDates(task);
                if (date >= start && date <= end) {
                    occupiedResources.add(task.name);
                }
            });
            return {
                date,
                occupied: occupiedResources.size,
                total: totalResources
            };
        });

        // Render capacity labels
        this.axisGroup.selectAll(".capacity-label")
            .data(capacityData)
            .enter()
            .append("text")
            .attr("class", "capacity-label")
            .attr("x", (d) => Gantt.TimeScale(d.date) + Gantt.DefaultTicksLength / 2)
            .attr("y", capacityLabelYOffset)
            .attr("text-anchor", "middle")
            .attr("font-size", "10px")
            .attr("fill", "#444")
            .text((d) => `${d.occupied} / ${d.total}`);

        // Render each day label as before
        this.axisGroup.selectAll(".day-label")
            .data(this.dailyTicks)
            .enter()
            .append("text")
            .attr("class", "day-label")
            .attr("x", (d: Date) => Gantt.TimeScale(d) + Gantt.DefaultTicksLength / 2)
            .attr("y", dayLabelYOffset)
            .attr("text-anchor", "middle")
            .attr("font-size", "10px")
            .attr("fill", "#888")
            .text((d: Date) => d.toLocaleDateString(undefined, { weekday: "short" }).slice(0, 2));
    }

    /**
     * Render task progress rect
     * @param taskSelection Task Selection
     */
    private taskProgressRender(
        taskSelection: Selection<Task>): void {
        const taskProgressShow: boolean = this.viewModel.settings.taskCompletionCardSettings.show.value;

        let index = 0, groupedTaskIndex = 0;
        const taskProgress: Selection<any> = taskSelection
            .selectAll(Gantt.TaskProgress.selectorName)
            .data((d: Task) => {
                const taskProgressPercentage = this.getDaysOffTaskProgressPercent(d);
                // logic used for grouped tasks, when there are several bars related to one category
                if (index === d.index) {
                    groupedTaskIndex++;
                } else {
                    groupedTaskIndex = 0;
                    index = d.index;
                }

                const url = `${d.index}-${groupedTaskIndex}-${isStringNotNullEmptyOrUndefined(d.taskType) ? d.taskType.toString() : "taskType"}`;
                const encodedUrl = `task${hashCode(url)}`;

                return [{
                    key: encodedUrl, values: <LinearStop[]>[
                        { completion: 0, color: d.color },
                        { completion: taskProgressPercentage, color: d.color },
                        { completion: taskProgressPercentage, color: d.color },
                        { completion: 1, color: d.color }
                    ]
                }];
            });

        const taskProgressMerged = taskProgress
            .enter()
            .append("linearGradient")
            .merge(taskProgress);

        taskProgressMerged.classed(Gantt.TaskProgress.className, true);

        taskProgressMerged
            .attr("id", (data) => data.key);

        const stopsSelection = taskProgressMerged.selectAll("stop");
        const stopsSelectionData = stopsSelection.data(gradient => <LinearStop[]>gradient.values);

        // draw 4 stops: 1st and 2d stops are for completed rect part; 3d and 4th ones -  for main rect
        stopsSelectionData.enter()
            .append("stop")
            .merge(<any>stopsSelection)
            .attr("offset", (data: LinearStop) => `${data.completion * 100}%`)
            .attr("stop-color", (data: LinearStop) => this.colorHelper.getHighContrastColor("foreground", data.color))
            .attr("stop-opacity", (_: LinearStop, index: number) => (index > 1) && taskProgressShow ? Gantt.NotCompletedTaskOpacity : Gantt.TaskOpacity);

        taskProgress
            .exit()
            .remove();
    }

    /**
     * Render task resource labels
     * @param taskSelection Task Selection
     * @param taskConfigHeight Task heights from settings
     */
    private taskResourceRender(
        taskSelection: Selection<Task>,
        taskConfigHeight: number,
        groupedTasks: GroupedTask[]
    ): void {

        const groupTasks: boolean = this.viewModel.settings.generalCardSettings.groupTasks.value;
        let newLabelPosition: ResourceLabelPosition | null = null;
        if (groupTasks && !this.groupTasksPrevValue) {
            newLabelPosition = ResourceLabelPosition.Inside;
        }

        if (!groupTasks && this.groupTasksPrevValue) {
            newLabelPosition = ResourceLabelPosition.Right;
        }

        if (newLabelPosition) {
            this.host.persistProperties(<VisualObjectInstancesToPersist>{
                merge: [{
                    objectName: "taskResource",
                    selector: null,
                    properties: { position: newLabelPosition }
                }]
            });

            this.viewModel.settings.taskResourceCardSettings.position.value.value = newLabelPosition;
            newLabelPosition = null;
        }

        this.groupTasksPrevValue = groupTasks;

        const isResourcesFilled: boolean = this.viewModel.isResourcesFilled;
        const taskResourceShow: boolean = this.viewModel.settings.taskResourceCardSettings.show.value;
        const taskResourceColor: string = this.viewModel.settings.taskResourceCardSettings.fill.value.value;
        const taskResourceFontSize: number = this.viewModel.settings.taskResourceCardSettings.fontSize.value;
        const taskResourcePosition: ResourceLabelPosition = ResourceLabelPosition[this.viewModel.settings.taskResourceCardSettings.position.value.value];
        const taskResourceFullText: boolean = this.viewModel.settings.taskResourceCardSettings.fullText.value;
        const taskResourceWidthByTask: boolean = this.viewModel.settings.taskResourceCardSettings.widthByTask.value;
        const isGroupedByTaskName: boolean = this.viewModel.settings.generalCardSettings.groupTasks.value;

        if (isResourcesFilled && taskResourceShow) {
            const taskResource: Selection<Task> = taskSelection
                .selectAll(Gantt.TaskResource.selectorName)
                .data((d: Task) => [d]);

            const taskResourceMerged = taskResource
                .enter()
                .append("text")
                .merge(taskResource);

            taskResourceMerged.classed(Gantt.TaskResource.className, true);

            taskResourceMerged
                .attr("x", (task: Task) => this.getResourceLabelXCoordinate(task, taskConfigHeight, taskResourceFontSize, taskResourcePosition))
                .attr("y", (task: Task) => {
                    const groupIndex = task.groupIndex;
                    const lane = task.lane || 0;
                    return Gantt.TaskBarTopPadding
                        + this.getGroupCumulativeY(groupedTasks, groupIndex)
                        + Gantt.getResourceLabelYOffset(taskConfigHeight, taskResourceFontSize, taskResourcePosition)
                        + lane * (Gantt.getBarHeight(taskConfigHeight) + 2);
                })
                .text((task: Task) => lodashIsEmpty(task.Milestones) && task.resource || "")
                .style("fill", taskResourceColor)
                .style("font-size", PixelConverter.fromPoint(taskResourceFontSize))
                .style("alignment-baseline", taskResourcePosition === ResourceLabelPosition.Inside ? "central" : "auto");

            const hasNotNullableDates: boolean = this.hasNotNullableDates;
            const defaultWidth: number = Gantt.DefaultValues.ResourceWidth - Gantt.ResourceWidthPadding;

            if (taskResourceWidthByTask) {
                taskResourceMerged
                    .each(function (task: Task) {
                        const width: number = hasNotNullableDates ? Gantt.taskDurationToWidth(task.start, task.end) : 0;
                        AxisHelper.LabelLayoutStrategy.clip(d3Select(this), width, textMeasurementService.svgEllipsis);
                    });
            } else if (isGroupedByTaskName) {
                taskResourceMerged
                    .each(function (task: Task, outerIndex: number) {
                        const sameRowNextTaskStart: Date = Gantt.getSameRowNextTaskStartDate(task, outerIndex, taskResourceMerged);

                        if (sameRowNextTaskStart) {
                            let width: number = 0;
                            if (hasNotNullableDates) {
                                const startDate: Date = taskResourcePosition === ResourceLabelPosition.Top ? task.start : task.end;
                                width = Gantt.taskDurationToWidth(startDate, sameRowNextTaskStart);
                            }

                            AxisHelper.LabelLayoutStrategy.clip(d3Select(this), width, textMeasurementService.svgEllipsis);
                        } else {
                            if (!taskResourceFullText) {
                                AxisHelper.LabelLayoutStrategy.clip(d3Select(this), defaultWidth, textMeasurementService.svgEllipsis);
                            }
                        }
                    });
            } else if (!taskResourceFullText) {
                taskResourceMerged
                    .each(function () {
                        AxisHelper.LabelLayoutStrategy.clip(d3Select(this), defaultWidth, textMeasurementService.svgEllipsis);
                    });
            }

            taskResource
                .exit()
                .remove();
        } else {
            taskSelection
                .selectAll(Gantt.TaskResource.selectorName)
                .remove();
        }

        this.renderStickyResourceLabelsInBars(
            groupedTasks, taskConfigHeight, taskResourceFontSize,
            taskResourceColor, taskResourcePosition, this.lastScrollLeft
        );
    }

    private static getMaxLane(tasks: Task[]): number {
        return Math.max(0, ...tasks.map(t => t.lane || 0));
    }

    private static getSameRowNextTaskStartDate(task: Task, index: number, selection: Selection<Task>) {
        let sameRowNextTaskStart: Date;

        selection
            .each(function (x: Task, i: number) {
                if (index !== i &&
                    x.index === task.index &&
                    x.start >= task.start &&
                    (!sameRowNextTaskStart || sameRowNextTaskStart < x.start)) {

                    sameRowNextTaskStart = x.start;
                }
            });

        return sameRowNextTaskStart;
    }

    private getBarXEndCoordinate(
        task: Task
    ): number {
        // Calculate the bar's start X using the time scale
        const { start, end } = Gantt.normalizeTaskDates(task);
        const barX = Gantt.TimeScale(start);
        const barWidth = Gantt.taskDurationToWidth(start, end);
        return barX + barWidth;
    }

    private renderStickyResourceLabelsInBars(
        groupedTasks: GroupedTask[],
        taskConfigHeight: number,
        taskResourceFontSize: number,
        taskResourceColor: string,
        taskResourcePosition: ResourceLabelPosition,
        scrollLeft: number
    ): void {
        const stickyClass = "sticky-resource-label";
        this.taskGroup.selectAll(`.${stickyClass}`).remove();

        const visibleLeft = scrollLeft;

        groupedTasks.forEach(group => {
            group.tasks.forEach(task => {
                const barEndX = this.getBarXEndCoordinate(task);
                const barStartX = Gantt.TimeScale(task.start); // <-- Add this line

                const labelText = lodashIsEmpty(task.Milestones) && task.resource || "";

                // Show sticky if the bar's left edge is left of the visible area and the bar is still visible
                const shouldShowSticky = barStartX < visibleLeft && barEndX > visibleLeft;

                // Hide the original label if sticky is shown
                if (shouldShowSticky) {
                    // Add sticky label at the left edge of the visible area
                    const stickyLabel = this.taskGroup.append("text")
                        .attr("class", `${Gantt.TaskResource.className} ${stickyClass}`)
                        .attr("x", visibleLeft + 2)
                        .attr("y", Gantt.TaskBarTopPadding
                            + this.getGroupCumulativeY(groupedTasks, task.groupIndex)
                            + Gantt.getResourceLabelYOffset(taskConfigHeight, taskResourceFontSize, taskResourcePosition)
                            + (task.lane || 0) * (Gantt.getBarHeight(taskConfigHeight) + 2))
                        .text(labelText)
                        .style("fill", taskResourceColor)
                        .style("font-size", PixelConverter.fromPoint(taskResourceFontSize))
                        .style("alignment-baseline", taskResourcePosition === ResourceLabelPosition.Inside ? "central" : "auto");

                    // Clip sticky label to visible width
                    const visibleWidth = barEndX - visibleLeft;
                    AxisHelper.LabelLayoutStrategy.clip(
                        stickyLabel,
                        visibleWidth,
                        textMeasurementService.svgEllipsis
                    );

                    // Hide the original label by setting its opacity to 0
                    this.taskGroup.selectAll(`.${Gantt.TaskResource.className}`)
                        .filter(function (d) { return d === task; })
                        .style("opacity", 0);
                } else {
                    // Ensure the original label is visible
                    this.taskGroup.selectAll(`.${Gantt.TaskResource.className}`)
                        .filter(function (d) { return d === task; })
                        .style("opacity", 1);
                }
            });
        });
    }

    private static normalizeTaskDates(task: Task): { start: Date, end: Date } {
        // Use Europe/Amsterdam for Dutch time
        const start = momentzone.tz(task.start, "Europe/Amsterdam").startOf('day').add(5, 'h').toDate();
        const end = momentzone.tz(task.end, "Europe/Amsterdam").endOf('day').add(5, 'h').toDate();
        return { start, end };
    }

    private static getResourceLabelYOffset(
        taskConfigHeight: number,
        taskResourceFontSize: number,
        taskResourcePosition: ResourceLabelPosition): number {
        const barHeight: number = Gantt.getBarHeight(taskConfigHeight);
        switch (taskResourcePosition) {
            case ResourceLabelPosition.Right:
                return (barHeight / Gantt.DividerForCalculatingCenter) + (taskResourceFontSize / Gantt.DividerForCalculatingCenter);
            case ResourceLabelPosition.Top:
                return -(taskResourceFontSize / Gantt.DividerForCalculatingPadding) + Gantt.LabelTopOffsetForPadding;
            case ResourceLabelPosition.Inside:
                return -(taskResourceFontSize / Gantt.DividerForCalculatingPadding) + Gantt.LabelTopOffsetForPadding + barHeight / Gantt.ResourceLabelDefaultDivisionCoefficient;
        }
    }

    private getResourceLabelXCoordinate(
        task: Task,
        taskConfigHeight: number,
        taskResourceFontSize: number,
        taskResourcePosition: ResourceLabelPosition): number {
        if (!this.hasNotNullableDates) {
            return 0;
        }

        const barHeight: number = Gantt.getBarHeight(taskConfigHeight);
        const leftPadding = -10;

        switch (taskResourcePosition) {
            case ResourceLabelPosition.Right:
                return (Gantt.TimeScale(task.end) + (taskResourceFontSize / 2) + Gantt.RectRound + leftPadding) || 0;
            case ResourceLabelPosition.Top:
                return (Gantt.TimeScale(task.start) + Gantt.RectRound + leftPadding) || 0;
            case ResourceLabelPosition.Inside:
                return (Gantt.TimeScale(task.start) + barHeight / (2 * Gantt.ResourceLabelDefaultDivisionCoefficient) + Gantt.RectRound + leftPadding) || 0;
        }
    }

    /**
     * Returns the matching Y coordinate for a given task index
     * @param taskIndex Task Number
     */
    private getTaskLabelCoordinateY(taskIndex: number, groupedTasks: GroupedTask[]): number {
        return this.getGroupCumulativeY(groupedTasks, taskIndex);
    }

    /**
    * Get completion percent when days off feature is on
    * @param task All task attributes
    */
    private getDaysOffTaskProgressPercent(task: Task) {
        if (this.viewModel.settings.daysOffCardSettings.show.value) {
            if (task.daysOffList && task.daysOffList.length && task.duration && task.completion) {
                let durationUnit: DurationUnit = <DurationUnit>this.viewModel.settings.generalCardSettings.durationUnit.value.value.toString();
                if (task.wasDowngradeDurationUnit) {
                    durationUnit = DurationHelper.downgradeDurationUnit(durationUnit, task.duration);
                }
                const startTime: number = task.start.getTime();
                const progressLength: number = (task.end.getTime() - startTime) * task.completion;
                const currentProgressTime: number = new Date(startTime + progressLength).getTime();

                const daysOffFiltered: DayOffData[] = task.daysOffList
                    .filter((date) => startTime <= date[0].getTime() && date[0].getTime() <= currentProgressTime);

                const extraDuration: number = Gantt.calculateExtraDurationDaysOff(daysOffFiltered, task.end, task.start, +this.viewModel.settings.daysOffCardSettings.firstDayOfWeek.value.value, durationUnit);
                const extraDurationPercentage = extraDuration / task.duration;
                return task.completion + extraDurationPercentage;
            }
        }

        return task.completion;
    }

    /**
    * Get bar y coordinate
    * @param lineNumber Line number that represents the task number
    * @param lineHeight Height of task line
    */
    private static getBarYCoordinate(lineNumber: number, lineHeight: number): number {
        return (lineHeight * lineNumber) + PaddingTasks;
    }

    /**
     * Get bar height
     * @param lineHeight The height of line
     */
    private static getBarHeight(lineHeight: number): number {
        return lineHeight / Gantt.ChartLineProportion;
    }

    /**
     * Get the margin that added to task rects and task category labels
     *
     * depends on resource label position and resource label font size
     */
    private getResourceLabelTopMargin(): number {
        const isResourcesFilled: boolean = this.viewModel.isResourcesFilled;
        const taskResourceShow: boolean = this.viewModel.settings.taskResourceCardSettings.show.value;
        const taskResourceFontSize: number = this.viewModel.settings.taskResourceCardSettings.fontSize.value;
        const taskResourcePosition: ResourceLabelPosition = ResourceLabelPosition[this.viewModel.settings.taskResourceCardSettings.position.value.value];

        let margin: number = 0;
        if (isResourcesFilled && taskResourceShow && taskResourcePosition === ResourceLabelPosition.Top) {
            margin = Number(taskResourceFontSize) + Gantt.LabelTopOffsetForPadding;
        }

        return margin;
    }

    /**
     * convert task duration to width in the timescale
     * @param start The start of task to convert
     * @param end The end of task to convert
     */
    private static taskDurationToWidth(
        start: Date,
        end: Date): number {
        return Gantt.TimeScale(end) - Gantt.TimeScale(start);
    }

    private static getTooltipForMilestoneLine(
        formattedDate: string,
        localizationManager: ILocalizationManager,
        dateTypeSettings: DateTypeCardSettings,
        milestoneTitle: string[] | LabelForDate[], milestoneCategoryName?: string[]): VisualTooltipDataItem[] {
        const result: VisualTooltipDataItem[] = [];

        for (let i = 0; i < milestoneTitle.length; i++) {
            if (!milestoneTitle[i]) {
                switch (dateTypeSettings.type.value.value) {
                    case DateType.Second:
                    case DateType.Minute:
                    case DateType.Hour:
                        milestoneTitle[i] = localizationManager.getDisplayName("Visual_Label_Now");
                        break;
                    default:
                        milestoneTitle[i] = localizationManager.getDisplayName("Visual_Label_Today");
                }
            }

            if (milestoneCategoryName) {
                result.push({
                    displayName: localizationManager.getDisplayName("Visual_Milestone_Name"),
                    value: milestoneCategoryName[i]
                });
            }

            result.push({
                displayName: <string>milestoneTitle[i],
                value: formattedDate
            });
        }

        return result;
    }

    /**
    * Create vertical dotted line that represent milestone in the time axis (by default it shows not time)
    * @param tasks All tasks array
    * @param milestoneTitle
    * @param timestamp the milestone to be shown in the time axis (default Date.now())
    */
    private createMilestoneLine(
        tasks: GroupedTask[],
        timestamp: number = Date.now(),
        milestoneTitle?: string): void {
        if (!this.hasNotNullableDates) {
            return;
        }

        const milestoneDates = [new Date(timestamp)];
        tasks.forEach((task: GroupedTask) => {
            const subtasks: Task[] = task.tasks;
            subtasks.forEach((task: Task) => {
                if (!lodashIsEmpty(task.Milestones)) {
                    task.Milestones.forEach((milestone) => {
                        if (!milestoneDates.includes(milestone.start)) {
                            milestoneDates.push(milestone.start);
                        }
                    });
                }
            });
        });

        const line: Line[] = [];
        const dateTypeSettings: DateTypeCardSettings = this.viewModel.settings.dateTypeCardSettings;
        milestoneDates.forEach((date: Date) => {
            const title = date === Gantt.TimeScale(timestamp) ? milestoneTitle : "Milestone";
            const lineOptions = {
                x1: Gantt.TimeScale(date),
                y1: Gantt.MilestoneTop,
                x2: Gantt.TimeScale(date),
                // y2: this.getMilestoneLineLengthFromGroupedTasks(tasks),
                y2: this.getMilestoneLineLengthFromGroupedTasks(),
                tooltipInfo: Gantt.getTooltipForMilestoneLine(date.toLocaleDateString(), this.localizationManager, dateTypeSettings, [title])
            };
            line.push(lineOptions);
        });

        const chartLineSelection: Selection<Line> = this.chartGroup
            .selectAll(Gantt.ChartLine.selectorName)
            .data(line);

        const chartLineSelectionMerged = chartLineSelection
            .enter()
            .append("line")
            .merge(chartLineSelection);

        chartLineSelectionMerged
            .classed(Gantt.ChartLine.className, true)
            .classed("today-line", (line: Line) => {
                // Mark the line for today (timestamp)
                return line.x1 === Gantt.TimeScale(timestamp);
            });

        chartLineSelectionMerged
            .attr("x1", (line: Line) => line.x1)
            .attr("y1", (line: Line) => line.y1)
            .attr("x2", (line: Line) => line.x2)
            .attr("y2", (line: Line) => line.y2)
            .style("stroke", (line: Line) =>
                line.x1 === Gantt.TimeScale(timestamp)
                    ? null // Let CSS handle today-line
                    : this.colorHelper.getHighContrastColor("foreground", Gantt.DefaultValues.MilestoneLineColor)
            );

        this.renderTooltip(chartLineSelectionMerged);

        chartLineSelection
            .exit()
            .remove();
    }

    // private getMilestoneLineLengthFromGroupedTasks(groupedTasks: GroupedTask[]): number {
    private getMilestoneLineLengthFromGroupedTasks(): number {
        // Sum all row heights and add top margin
        // return groupedTasks.reduce((sum, g) => sum + g.rowHeight, 0) + this.margin.top;

        // Match the SVG height exactly
        return parseFloat(this.ganttSvg.attr("height") || "0");
    }

    private scrollToMilestoneLine(axisLength: number,
        timestamp: number = Date.now()): void {

        let scrollValue = Gantt.TimeScale(new Date(timestamp));
        scrollValue -= scrollValue > ScrollMargin
            ? ScrollMargin
            : 0;

        if (axisLength > scrollValue) {
            (this.body.node() as SVGSVGElement)
                .querySelector(Gantt.Body.selectorName).scrollLeft = scrollValue;
        }
    }

    private renderTooltip(selection: Selection<Line | Task | MilestonePath>): void {
        this.tooltipServiceWrapper.addTooltip(
            selection,
            (tooltipEvent: TooltipEnabledDataPoint) => tooltipEvent.tooltipInfo);
    }

    private updateElementsPositions(margin: IMargin): void {
        const settings: GanttChartSettingsModel = this.viewModel.settings;
        const taskLabelsWidth: number = settings.taskLabelsCardSettings.show.value
            ? settings.taskLabelsCardSettings.width.value
            : 0;

        let translateXValue: number = taskLabelsWidth + margin.left + Gantt.SubtasksLeftMargin;
        this.chartGroup
            .attr("transform", SVGManipulations.translate(translateXValue, margin.top));

        const translateYValue: number = Gantt.TaskLabelsMarginTop + (this.ganttDiv.node() as SVGSVGElement).scrollTop;
        this.axisGroup
            .attr("transform", SVGManipulations.translate(translateXValue, translateYValue));

        translateXValue = (this.ganttDiv.node() as SVGSVGElement).scrollLeft;
        this.lineGroup
            .attr("transform", SVGManipulations.translate(translateXValue, 0));
        this.collapseAllGroup
            .attr("transform", SVGManipulations.translate(0, margin.top / 4 + Gantt.AxisTopMargin));
    }

    private getMilestoneLineLength(numOfTasks: number): number {
        return numOfTasks * ((this.viewModel.settings.taskConfigCardSettings.height.value || DefaultChartLineHeight) + (1 + numOfTasks) * this.getResourceLabelTopMargin() / 2);
    }

    public static downgradeDurationUnitIfNeeded(tasks: Task[], durationUnit: DurationUnit) {
        const downgradedDurationUnitTasks = tasks.filter(t => t.wasDowngradeDurationUnit);

        if (downgradedDurationUnitTasks.length) {
            let maxStepDurationTransformation: number = 0;
            downgradedDurationUnitTasks.forEach(x => maxStepDurationTransformation = x.stepDurationTransformation > maxStepDurationTransformation ? x.stepDurationTransformation : maxStepDurationTransformation);

            tasks.filter(x => x.stepDurationTransformation !== maxStepDurationTransformation).forEach(task => {
                task.duration = DurationHelper.transformDuration(task.duration, durationUnit, maxStepDurationTransformation);
                task.stepDurationTransformation = maxStepDurationTransformation;
                task.wasDowngradeDurationUnit = true;
            });
        }
    }

    public getFormattingModel(): powerbi.visuals.FormattingModel {
        this.filterSettingsCards();
        this.formattingSettings.setLocalizedOptions(this.localizationManager);
        return this.formattingSettingsService.buildFormattingModel(this.formattingSettings);
    }

    public filterSettingsCards() {
        const settings: GanttChartSettingsModel = this.formattingSettings;

        settings.cards.forEach(element => {
            switch (element.name) {
                case Gantt.MilestonesPropertyIdentifier.objectName: {
                    if (this.viewModel && !this.viewModel.isDurationFilled && !this.viewModel.isEndDateFilled) {
                        return;
                    }

                    const dataPoints: MilestoneDataPoint[] = this.viewModel && this.viewModel.milestonesData.dataPoints;
                    if (!dataPoints || !dataPoints.length) {
                        settings.milestonesCardSettings.visible = false;
                        return;
                    }

                    const milestonesWithoutDuplicates = Gantt.getUniqueMilestones(dataPoints);

                    settings.populateMilestones(milestonesWithoutDuplicates);
                    break;
                }

                case Gantt.LegendPropertyIdentifier.objectName: {
                    if (this.viewModel && !this.viewModel.isDurationFilled && !this.viewModel.isEndDateFilled) {
                        return;
                    }

                    const dataPoints: LegendDataPoint[] = this.viewModel && this.viewModel.legendData.dataPoints;
                    if (!dataPoints || !dataPoints.length) {
                        return;
                    }

                    settings.populateLegend(dataPoints, this.localizationManager);
                    break;
                }

                case Gantt.TaskResourcePropertyIdentifier.objectName:
                    if (!this.viewModel.isResourcesFilled) {
                        settings.taskResourceCardSettings.visible = false;
                    }
                    break;
            }
        });
    }
}