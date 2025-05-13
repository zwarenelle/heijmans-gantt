import { Selection as d3Selection } from "d3-selection";

type Selection<T1, T2 = T1> = d3Selection<any, T1, any, T2>;

import { interactivityBaseService as interactivityService } from "powerbi-visuals-utils-interactivityutils";
import IInteractiveBehavior = interactivityService.IInteractiveBehavior;
import IInteractivityService = interactivityService.IInteractivityService;
import ISelectionHandler = interactivityService.ISelectionHandler;

import { Task, GroupedTask } from "./interfaces";
import { IBehaviorOptions } from "powerbi-visuals-utils-interactivityutils/lib/interactivityBaseService";

export const DimmedOpacity: number = 0.4;
export const DefaultOpacity: number = 1.0;

export function getFillOpacity(
    selected: boolean,
    highlight: boolean,
    hasSelection: boolean,
    hasPartialHighlights: boolean
): number {
    if ((hasPartialHighlights && !highlight) || (hasSelection && !selected)) {
        return DimmedOpacity;
    }

    return DefaultOpacity;
}

export interface BehaviorOptions extends IBehaviorOptions<Task> {
    clearCatcher: Selection<any>;
    taskSelection: Selection<Task>;
    legendSelection: Selection<any>;
    interactivityService: IInteractivityService<Task>;
    subTasksCollapse: {
        selection: Selection<any>;
        callback: (groupedTask: GroupedTask) => void;
    };
    allSubtasksCollapse: {
        selection: Selection<any>;
        callback: () => void;
    };
}

export class Behavior implements IInteractiveBehavior {
    private options: BehaviorOptions;
    private selectionHandler: ISelectionHandler;

    public bindEvents(options: BehaviorOptions, selectionHandler: ISelectionHandler) {
        this.options = options;
        this.selectionHandler = selectionHandler;
        const clearCatcher = options.clearCatcher;

        this.bindContextMenu();

        options.taskSelection.on("click", (event: MouseEvent, dataPoint: Task) => {
            selectionHandler.handleSelection(dataPoint, event.ctrlKey || event.metaKey);

            event.stopPropagation();
        });

        options.legendSelection.on("click", (event: MouseEvent, d: any) => {
            if (d.selected) {
                selectionHandler.handleClearSelection();
                return;
            }

            selectionHandler.handleSelection(d, event.ctrlKey || event.metaKey);
            event.stopPropagation();

            const selectedType: string = d.tooltipInfo;
            options.taskSelection.each((d: Task) => {
                if (d.taskType === selectedType && d.parent && !d.selected) {
                    selectionHandler.handleSelection(d, event.ctrlKey || event.metaKey);
                }
            });
        });

        options.subTasksCollapse.selection.on("click", (event: MouseEvent, d: GroupedTask) => {
            if (!d.tasks.map(task => task.children).flat().length) {
                return;
            }

            event.stopPropagation();
            options.subTasksCollapse.callback(d);
        });

        options.allSubtasksCollapse.selection.on("click", (event: MouseEvent) => {
            event.stopPropagation();
            options.allSubtasksCollapse.callback();
        });

        clearCatcher.on("click", () => {
            selectionHandler.handleClearSelection();
        });
    }

    public renderSelection(hasSelection: boolean) {
        const {
            taskSelection,
            interactivityService,
        } = this.options;

        const hasHighlights: boolean = interactivityService.hasSelection();

        taskSelection.style("opacity", (dataPoint: Task) => {
            return getFillOpacity(
                dataPoint.selected,
                dataPoint.highlight,
                !dataPoint.highlight && hasSelection,
                !dataPoint.selected && hasHighlights
            );
        });
    }

    private bindContextMenu(): void {
        this.options.taskSelection.on("contextmenu", (event: MouseEvent, task: Task) => {
            if (event) {
                this.selectionHandler.handleContextMenu(
                    task,
                    {
                        x: event.clientX,
                        y: event.clientY
                    });
                event.preventDefault();
                event.stopPropagation();
            }
        });

        this.options.legendSelection.on("contextmenu", (event: MouseEvent, legend: any) => {
            if (event) {
                this.selectionHandler.handleContextMenu(
                    legend,
                    {
                        x: event.clientX,
                        y: event.clientY
                    });
                event.preventDefault();
                event.stopPropagation();
            }
        });

        this.options.clearCatcher.on("contextmenu", (event: MouseEvent) => {
            if (event) {
                this.selectionHandler.handleContextMenu(
                    null,
                    {
                        x: event.clientX,
                        y: event.clientY
                    });
                event.preventDefault();
            }
        });
    }
}
