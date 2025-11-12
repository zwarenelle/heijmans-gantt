import powerbi from "powerbi-visuals-api";

import DataView = powerbi.DataView;
import IViewport = powerbi.IViewport;
import DataViewCategoryColumn = powerbi.DataViewCategoryColumn;
import DataViewValueColumnGroup = powerbi.DataViewValueColumnGroup;
import ISelectionId = powerbi.visuals.ISelectionId;
import VisualTooltipDataItem = powerbi.extensibility.VisualTooltipDataItem;

import { interactivitySelectionService as interactivityService } from "powerbi-visuals-utils-interactivityutils";
import SelectableDataPoint = interactivityService.SelectableDataPoint;

import { valueFormatter as vf } from "powerbi-visuals-utils-formattingutils";
import IValueFormatter = vf.IValueFormatter;

import { legendInterfaces } from "powerbi-visuals-utils-chartutils";
import LegendData = legendInterfaces.LegendData;

import * as SVGUtil from "powerbi-visuals-utils-svgutils";
import IMargin = SVGUtil.IMargin;

import { GanttChartSettingsModel } from "./ganttChartSettingsModels";

export type DayOffData = [Date, number];

export interface DaysOffDataForAddition {
    list: DayOffData[];
    amountOfLastDaysOff: number;
}

export interface TaskDaysOff {
    id: number;
    daysOff: DayOffData;
}

export interface ExtraInformation {
    displayName: string;
    value: string;
}

export interface Task extends SelectableDataPoint {
    index: number;
    name: string;
    start: Date;
    duration: number;
    completion: number;
    resource: string;
    end: Date;
    parent: string;
    children: Task[];
    visibility: boolean;
    taskType: string;
    description: string;
    color: string;
    tooltipInfo: VisualTooltipDataItem[];
    extraInformation: ExtraInformation[];
    daysOffList: DayOffData[];
    wasDowngradeDurationUnit: boolean;
    stepDurationTransformation?: number;
    highlight?: boolean;
    Milestones?: Milestone[];
    lane: number;
    groupIndex: number;
    id: string;
}

export interface GroupedTask {
    index: number;
    name: string;
    tasks: Task[];
    rowHeight: number;
    maxLane: number;
}

export interface GanttChartFormatters {
    startDateFormatter: IValueFormatter;
    completionFormatter: IValueFormatter;
}

export interface GanttViewModel {
    dataView: DataView;
    settings: GanttChartSettingsModel;
    tasks: Task[];
    legendData: LegendData;
    milestonesData: MilestoneData;
    taskTypes: TaskTypes;
    isDurationFilled: boolean;
    isEndDateFilled: boolean;
    isParentFilled: boolean;
    isResourcesFilled: boolean;
}

export interface TaskTypes { /*TODO: change to more proper name*/
    typeName: string;
    types: TaskTypeMetadata[];
}

export interface TaskTypeMetadata {
    name: string;
    columnGroup: DataViewValueColumnGroup;
    selectionColumn: DataViewCategoryColumn;
}

export interface GanttCalculateScaleAndDomainOptions {
    viewport: IViewport;
    margin: IMargin;
    showCategoryAxisLabel: boolean;
    showValueAxisLabel: boolean;
    forceMerge: boolean;
    categoryAxisScaleType: string;
    valueAxisScaleType: string;
    trimOrdinalDataOnOverflow: boolean;
    forcedTickCount?: number;
    forcedYDomain?: any[];
    forcedXDomain?: any[];
    ensureXDomain?: any;
    ensureYDomain?: any;
    categoryAxisDisplayUnits?: number;
    categoryAxisPrecision?: number;
    valueAxisDisplayUnits?: number;
    valueAxisPrecision?: number;
}

export interface Line {
    x1: number;
    y1: number;
    x2: number;
    y2: number;
    tooltipInfo: VisualTooltipDataItem[];
}

export interface LinearStop {
    completion: number;
    color: string;
}

export interface Milestone {
    type: string;
    category?: string;
    start: Date;
    tooltipInfo: VisualTooltipDataItem[];
}

export interface MilestonePath extends Milestone {
    taskID: number;
}

export interface MilestoneDataPoint {
    name: string;
    shapeType: string;
    color: string;
    identity: ISelectionId;
}

export interface MilestoneData {
    dataPoints: MilestoneDataPoint[];
}
