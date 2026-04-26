import powerbi from "powerbi-visuals-api";
import IVisual = powerbi.extensibility.IVisual;
import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import EnumerateVisualObjectInstancesOptions = powerbi.EnumerateVisualObjectInstancesOptions;
import VisualObjectInstanceEnumeration = powerbi.VisualObjectInstanceEnumeration;
import DataView = powerbi.DataView;
import ISelectionManager = powerbi.extensibility.ISelectionManager;
import ISelectionIdBuilder = powerbi.extensibility.ISelectionIdBuilder;

import * as d3 from "d3";
type Selection<T1, T2 = any> = d3.Selection<any, T1, any, T2>;

import {
    VisualSettings,
    DefaultSettings,
    getSettings,
    getDashArray
} from "./settings";

interface DataPoint {
    value: number;
    category: string;
    categoryIndex: number;
    phase: string;
    isOOC: boolean;
    oocRule: string;
    identity: powerbi.extensibility.ISelectionId;
}

interface PhaseStats {
    phase: string;
    mean: number;
    stdDev: number;
    ucl: number;
    lcl: number;
    sigma1Upper: number;
    sigma1Lower: number;
    sigma2Upper: number;
    sigma2Lower: number;
    startIndex: number;
    endIndex: number;
    count: number;
}

interface ChartStats {
    mean: number;
    stdDev: number;
    ucl: number;
    lcl: number;
    min: number;
    max: number;
    range: number;
    n: number;
    oocCount: number;
    oocPercent: number;
    cp: number | null;
    cpk: number | null;
    pp: number | null;
    ppk: number | null;
}

const MARGIN = { top: 40, right: 30, bottom: 50, left: 60 };

export class Visual implements IVisual {
    private target: HTMLElement;
    private settings: VisualSettings;
    private selectionManager: ISelectionManager;
    private selectionIdBuilder: ISelectionIdBuilder;
    private svg: Selection<any>;
    private chartGroup: Selection<any>;
    private dataPoints: DataPoint[];
    private phaseStats: PhaseStats[];
    private overallStats: ChartStats;

    constructor(options: VisualConstructorOptions) {
        this.target = options.element;
        this.selectionManager = options.host.createSelectionManager();
        this.selectionIdBuilder = options.host.createSelectionIdBuilder();

        this.svg = d3.select(this.target)
            .append("svg")
            .classed("controlChart", true);

        this.chartGroup = this.svg.append("g").classed("chartGroup", true);
    }

    public update(options: VisualUpdateOptions): void {
        const dataView: DataView = options.dataViews[0];
        if (!dataView || !dataView.categorical) {
            this.clearChart();
            return;
        }

        this.settings = getSettings(dataView.metadata ? dataView.metadata.objects : {});
        this.dataPoints = this.parseData(dataView);
        this.phaseStats = this.calculatePhaseStats();
        this.overallStats = this.calculateOverallStats();
        this.applyOOCRules();

        const width = options.viewport.width - MARGIN.left - MARGIN.right;
        const height = options.viewport.height - MARGIN.top - MARGIN.bottom;

        if (width <= 0 || height <= 0) return;

        this.svg.attr("width", options.viewport.width).attr("height", options.viewport.height);
        this.chartGroup.attr("transform", `translate(${MARGIN.left},${MARGIN.top})`);

        this.clearChart();

        const xScale = this.createXScale(width);
        const yScale = this.createYScale(height);

        this.renderZones(xScale, yScale, height);
        this.renderAxes(xScale, yScale, width, height);
        this.renderPhaseSeparators(xScale, height);
        this.renderControlLines(xScale, yScale, width);
        this.renderSpecLimits(xScale, yScale, width);
        this.renderTargetLine(xScale, yScale, width);
        this.renderDataLine(xScale, yScale);
        this.renderDataPoints(xScale, yScale);
        this.renderLineLabels(xScale, yScale, width);
        this.renderStatisticsPanel(width, height);
    }

    private parseData(dataView: DataView): DataPoint[] {
        const categorical = dataView.categorical;
        const categories = categorical.categories;
        const values = categorical.values;

        if (!values || values.length === 0) return [];

        const measureValues = values[0] ? values[0].values : [];

        let sampleValues: (string | null)[] = [];
        let phaseValues: (string | null)[] = [];

        if (categories && categories.length > 0) {
            if (categories[0] && categories[0].source && categories[0].source.roles && categories[0].source.roles["sample"]) {
                sampleValues = categories[0].values as (string | null)[];
            }
            const phaseIdx = categories.findIndex(c => c.source && c.source.roles && c.source.roles["phase"]);
            if (phaseIdx >= 0) {
                phaseValues = categories[phaseIdx].values as (string | null)[];
            }
        }

        const dataPoints: DataPoint[] = [];
        const count = measureValues.length;

        for (let i = 0; i < count; i++) {
            const val = measureValues[i] as number;
            if (val === null || val === undefined || isNaN(val)) continue;

            const category = (sampleValues[i] != null && sampleValues[i] !== "")
                ? String(sampleValues[i])
                : String(i + 1);

            const phase = (phaseValues[i] != null && phaseValues[i] !== "")
                ? String(phaseValues[i])
                : "All";

            const identity = this.selectionIdBuilder
                .withCategory(categories[0], i)
                .createSelectionId();

            dataPoints.push({
                value: val,
                category,
                categoryIndex: i,
                phase,
                isOOC: false,
                oocRule: "",
                identity
            });
        }

        return dataPoints;
    }

    private calculatePhaseStats(): PhaseStats[] {
        if (!this.dataPoints || this.dataPoints.length === 0) return [];

        const phaseMap = new Map<string, { points: DataPoint[], indices: number[] }>();
        this.dataPoints.forEach((dp, idx) => {
            if (!phaseMap.has(dp.phase)) {
                phaseMap.set(dp.phase, { points: [], indices: [] });
            }
            dp.categoryIndex = idx;
            phaseMap.get(dp.phase).points.push(dp);
            phaseMap.get(dp.phase).indices.push(idx);
        });

        const stats: PhaseStats[] = [];

        phaseMap.forEach((data, phase) => {
            const values = data.points.map(p => p.value);
            const mean = d3.mean(values);
            const stdDev = this.calculateStdDev(values);

            stats.push({
                phase,
                mean,
                stdDev,
                ucl: mean + 3 * stdDev,
                lcl: mean - 3 * stdDev,
                sigma1Upper: mean + stdDev,
                sigma1Lower: mean - stdDev,
                sigma2Upper: mean + 2 * stdDev,
                sigma2Lower: mean - 2 * stdDev,
                startIndex: data.indices[0],
                endIndex: data.indices[data.indices.length - 1],
                count: values.length
            });
        });

        return stats;
    }

    private calculateOverallStats(): ChartStats {
        const values = this.dataPoints.map(p => p.value);
        const mean = d3.mean(values);
        const stdDev = this.calculateStdDev(values);

        const oocCount = this.dataPoints.filter(p => p.isOOC).length;

        const usl = this.getEffectiveUSL();
        const lsl = this.getEffectiveLSL();

        let cp: number | null = null;
        let cpk: number | null = null;
        let pp: number | null = null;
        let ppk: number | null = null;

        if (lsl !== null && usl !== null && stdDev > 0) {
            cp = (usl - lsl) / (6 * stdDev);
            cpk = Math.min((usl - mean) / (3 * stdDev), (mean - lsl) / (3 * stdDev));
            const overallStdDev = this.calculateOverallStdDev(values);
            if (overallStdDev > 0) {
                pp = (usl - lsl) / (6 * overallStdDev);
                ppk = Math.min((usl - mean) / (3 * overallStdDev), (mean - lsl) / (3 * overallStdDev));
            }
        }

        return {
            mean,
            stdDev,
            ucl: mean + 3 * stdDev,
            lcl: mean - 3 * stdDev,
            min: d3.min(values),
            max: d3.max(values),
            range: d3.max(values) - d3.min(values),
            n: values.length,
            oocCount,
            oocPercent: values.length > 0 ? (oocCount / values.length) * 100 : 0,
            cp,
            cpk,
            pp,
            ppk
        };
    }

    private calculateStdDev(values: number[]): number {
        if (values.length < 2) return 0;
        const n = values.length;
        const mean = d3.mean(values);
        const sqDiffs = values.map(v => Math.pow(v - mean, 2));
        return Math.sqrt(d3.sum(sqDiffs) / (n - 1));
    }

    private calculateOverallStdDev(values: number[]): number {
        if (values.length < 2) return 0;
        const n = values.length;
        const mean = d3.mean(values);
        const sqDiffs = values.map(v => Math.pow(v - mean, 2));
        return Math.sqrt(d3.sum(sqDiffs) / n);
    }

    private getEffectiveLSL(): number | null {
        if (this.settings.specLimits.show && this.settings.specLimits.lslValue !== 0) {
            return this.settings.specLimits.lslValue;
        }
        return null;
    }

    private getEffectiveUSL(): number | null {
        if (this.settings.specLimits.show && this.settings.specLimits.uslValue !== 0) {
            return this.settings.specLimits.uslValue;
        }
        return null;
    }

    private getEffectiveTarget(): number | null {
        if (this.settings.targetLine.show && this.settings.targetLine.targetValue !== 0) {
            return this.settings.targetLine.targetValue;
        }
        return null;
    }

    private applyOOCRules(): void {
        this.dataPoints.forEach((dp) => {
            dp.isOOC = false;
            dp.oocRule = "";
        });

        const usePerPhase = this.settings.phaseControl.recalculatePerPhase && this.phaseStats.length > 1;

        if (usePerPhase) {
            this.phaseStats.forEach((stats) => {
                const phasePoints = this.dataPoints.slice(stats.startIndex, stats.endIndex + 1);
                this.applyOOCRulesToSubset(phasePoints, stats.mean, stats.stdDev, stats.ucl, stats.lcl);
            });
        } else {
            this.applyOOCRulesToSubset(
                this.dataPoints,
                this.overallStats.mean,
                this.overallStats.stdDev,
                this.overallStats.ucl,
                this.overallStats.lcl
            );
        }

        this.overallStats.oocCount = this.dataPoints.filter(p => p.isOOC).length;
        this.overallStats.oocPercent = this.dataPoints.length > 0
            ? (this.overallStats.oocCount / this.dataPoints.length) * 100
            : 0;
    }

    private applyOOCRulesToSubset(
        points: DataPoint[],
        mean: number,
        stdDev: number,
        ucl: number,
        lcl: number
    ): void {
        const rules = this.settings.oocRules;

        if (rules.rule1) {
            points.forEach((dp) => {
                if (dp.value > ucl || dp.value < lcl) {
                    dp.isOOC = true;
                    dp.oocRule = "Rule 1";
                }
            });
        }

        if (rules.rule2) {
            const windowSize = 9;
            for (let i = windowSize - 1; i < points.length; i++) {
                const window = points.slice(i - windowSize + 1, i + 1);
                const allAbove = window.every(p => p.value > mean);
                const allBelow = window.every(p => p.value < mean);
                if (allAbove || allBelow) {
                    window.forEach(p => { p.isOOC = true; p.oocRule = "Rule 2"; });
                }
            }
        }

        if (rules.rule3) {
            const windowSize = 6;
            for (let i = windowSize - 1; i < points.length; i++) {
                const window = points.slice(i - windowSize + 1, i + 1);
                let increasing = true;
                let decreasing = true;
                for (let j = 1; j < window.length; j++) {
                    if (window[j].value <= window[j - 1].value) increasing = false;
                    if (window[j].value >= window[j - 1].value) decreasing = false;
                }
                if (increasing || decreasing) {
                    window.forEach(p => { p.isOOC = true; p.oocRule = "Rule 3"; });
                }
            }
        }

        if (rules.rule4) {
            const windowSize = 14;
            for (let i = windowSize - 1; i < points.length; i++) {
                const window = points.slice(i - windowSize + 1, i + 1);
                let alternating = true;
                for (let j = 1; j < window.length; j++) {
                    const diff = window[j].value - window[j - 1].value;
                    const prevDiff = window[j - 1].value - (j >= 2 ? window[j - 2].value : window[j - 1].value);
                    if (diff * prevDiff >= 0) {
                        alternating = false;
                        break;
                    }
                }
                if (alternating) {
                    window.forEach(p => { p.isOOC = true; p.oocRule = "Rule 4"; });
                }
            }
        }

        if (rules.rule5) {
            const windowSize = 3;
            const threshold = mean + 2 * stdDev;
            const thresholdLow = mean - 2 * stdDev;
            for (let i = windowSize - 1; i < points.length; i++) {
                const window = points.slice(i - windowSize + 1, i + 1);
                const beyondUpper = window.filter(p => p.value > threshold).length;
                const beyondLower = window.filter(p => p.value < thresholdLow).length;
                if (beyondUpper >= 2 || beyondLower >= 2) {
                    window.forEach(p => { p.isOOC = true; p.oocRule = "Rule 5"; });
                }
            }
        }

        if (rules.rule6) {
            const windowSize = 5;
            const threshold = mean + stdDev;
            const thresholdLow = mean - stdDev;
            for (let i = windowSize - 1; i < points.length; i++) {
                const window = points.slice(i - windowSize + 1, i + 1);
                const beyondUpper = window.filter(p => p.value > threshold).length;
                const beyondLower = window.filter(p => p.value < thresholdLow).length;
                if (beyondUpper >= 4 || beyondLower >= 4) {
                    window.forEach(p => { p.isOOC = true; p.oocRule = "Rule 6"; });
                }
            }
        }
    }

    private createXScale(width: number): d3.ScalePoint<string> {
        return d3.scalePoint<string>()
            .domain(this.dataPoints.map(d => d.category))
            .range([0, width])
            .padding(0.5);
    }

    private createYScale(height: number): d3.ScaleLinear<number, number> {
        const values = this.dataPoints.map(d => d.value);
        let minVal = d3.min(values);
        let maxVal = d3.max(values);

        const lsl = this.getEffectiveLSL();
        const usl = this.getEffectiveUSL();
        const target = this.getEffectiveTarget();

        if (lsl !== null) minVal = Math.min(minVal, lsl);
        if (usl !== null) maxVal = Math.max(maxVal, usl);
        if (target !== null) {
            minVal = Math.min(minVal, target);
            maxVal = Math.max(maxVal, target);
        }

        if (this.settings.controlLines.show && this.settings.controlLines.show3Sigma) {
            minVal = Math.min(minVal, this.overallStats.lcl);
            maxVal = Math.max(maxVal, this.overallStats.ucl);
        }

        const padding = (maxVal - minVal) * 0.1 || 1;
        return d3.scaleLinear()
            .domain([minVal - padding, maxVal + padding])
            .range([height, 0]);
    }

    private renderZones(xScale: d3.ScalePoint<string>, yScale: d3.ScaleLinear<number, number>, height: number): void {
        if (!this.settings.zones.show) return;

        const mean = this.overallStats.mean;
        const stdDev = this.overallStats.stdDev;
        const opacity = this.settings.zones.zoneOpacity;

        const xRange = xScale.range();
        const zoneWidth = xRange[1] - xRange[0] + 40;

        const zoneData = [
            { y1: mean + 2 * stdDev, y2: mean + 3 * stdDev, color: this.settings.zones.zoneAColor },
            { y1: mean + 1 * stdDev, y2: mean + 2 * stdDev, color: this.settings.zones.zoneBColor },
            { y1: mean, y2: mean + 1 * stdDev, color: this.settings.zones.zoneCColor },
            { y1: mean - 1 * stdDev, y2: mean, color: this.settings.zones.zoneCColor },
            { y1: mean - 2 * stdDev, y2: mean - 1 * stdDev, color: this.settings.zones.zoneBColor },
            { y1: mean - 3 * stdDev, y2: mean - 2 * stdDev, color: this.settings.zones.zoneAColor }
        ];

        this.chartGroup.selectAll(".zone")
            .data(zoneData)
            .enter()
            .append("rect")
            .classed("zone", true)
            .attr("x", -20)
            .attr("width", zoneWidth + 40)
            .attr("y", d => yScale(d.y2))
            .attr("height", d => Math.abs(yScale(d.y1) - yScale(d.y2)))
            .attr("fill", d => d.color)
            .attr("opacity", opacity);
    }

    private renderAxes(xScale: d3.ScalePoint<string>, yScale: d3.ScaleLinear<number, number>, width: number, height: number): void {
        const axisSettings = this.settings.axis;

        if (axisSettings.showXAxis) {
            const tickInterval = Math.max(1, Math.floor(this.dataPoints.length / 15));
            const xAxis = d3.axisBottom(xScale)
                .tickValues(
                    this.dataPoints
                        .filter((_, i) => i % tickInterval === 0)
                        .map(d => d.category)
                );

            this.chartGroup.append("g")
                .classed("x-axis", true)
                .attr("transform", `translate(0,${height})`)
                .call(xAxis)
                .selectAll("text")
                .attr("fill", axisSettings.axisColor)
                .style("font-size", `${axisSettings.axisFontSize}px`)
                .attr("transform", "rotate(-45)")
                .style("text-anchor", "end");

            this.chartGroup.select(".x-axis .domain")
                .attr("stroke", axisSettings.axisColor);
            this.chartGroup.select(".x-axis .tick line")
                .attr("stroke", axisSettings.axisColor);

            const xLabelText = axisSettings.xLabel || "Sample";
            this.chartGroup.append("text")
                .attr("x", width / 2)
                .attr("y", height + 45)
                .attr("text-anchor", "middle")
                .attr("fill", axisSettings.axisColor)
                .style("font-size", `${axisSettings.axisFontSize + 2}px`)
                .text(xLabelText);
        }

        if (axisSettings.showYAxis) {
            const yAxis = d3.axisLeft(yScale).ticks(8);

            this.chartGroup.append("g")
                .classed("y-axis", true)
                .call(yAxis)
                .selectAll("text")
                .attr("fill", axisSettings.axisColor)
                .style("font-size", `${axisSettings.axisFontSize}px`);

            this.chartGroup.select(".y-axis .domain")
                .attr("stroke", axisSettings.axisColor);
            this.chartGroup.select(".y-axis .tick line")
                .attr("stroke", axisSettings.axisColor);

            if (axisSettings.yLabel) {
                this.chartGroup.append("text")
                    .attr("transform", "rotate(-90)")
                    .attr("y", -50)
                    .attr("x", -height / 2)
                    .attr("text-anchor", "middle")
                    .attr("fill", axisSettings.axisColor)
                    .style("font-size", `${axisSettings.axisFontSize + 2}px`)
                    .text(axisSettings.yLabel);
            }
        }
    }

    private renderPhaseSeparators(xScale: d3.ScalePoint<string>, height: number): void {
        if (!this.settings.phaseControl.show || this.phaseStats.length <= 1) return;

        const settings = this.settings.phaseControl;

        for (let i = 1; i < this.phaseStats.length; i++) {
            const prevEnd = this.phaseStats[i - 1].endIndex;
            const x = xScale(this.dataPoints[prevEnd].category);

            this.chartGroup.append("line")
                .classed("phase-separator", true)
                .attr("x1", x)
                .attr("x2", x)
                .attr("y1", 0)
                .attr("y2", height)
                .attr("stroke", settings.separatorColor)
                .attr("stroke-width", settings.separatorWidth)
                .attr("stroke-dasharray", getDashArray(settings.separatorStyle));

            if (settings.showPhaseLabels) {
                this.chartGroup.append("text")
                    .classed("phase-label", true)
                    .attr("x", x)
                    .attr("y", -5)
                    .attr("text-anchor", "middle")
                    .attr("fill", settings.separatorColor)
                    .style("font-size", `${settings.phaseLabelFontSize}px`)
                    .text(this.phaseStats[i].phase);
            }
        }
    }

    private renderControlLines(xScale: d3.ScalePoint<string>, yScale: d3.ScaleLinear<number, number>, width: number): void {
        if (!this.settings.controlLines.show) return;

        const settings = this.settings.controlLines;
        const xRange = xScale.range();
        const x1 = xRange[0] - 20;
        const x2 = xRange[1] + 20;

        if (this.phaseStats.length <= 1 || !this.settings.phaseControl.recalculatePerPhase) {
            const mean = this.overallStats.mean;
            const stdDev = this.overallStats.stdDev;

            if (settings.show3Sigma) {
                this.drawHorizontalLine(x1, x2, yScale(mean), settings.clColor, settings.clStyle, settings.clWidth, "cl");
                this.drawHorizontalLine(x1, x2, yScale(mean + 3 * stdDev), settings.uclColor, settings.uclStyle, settings.uclWidth, "ucl");
                this.drawHorizontalLine(x1, x2, yScale(mean - 3 * stdDev), settings.lclColor, settings.lclStyle, settings.lclWidth, "lcl");
            }

            if (settings.show2Sigma) {
                this.drawHorizontalLine(x1, x2, yScale(mean + 2 * stdDev), settings.sigma2Color, "dotted", 1, "sigma2u");
                this.drawHorizontalLine(x1, x2, yScale(mean - 2 * stdDev), settings.sigma2Color, "dotted", 1, "sigma2l");
            }

            if (settings.show1Sigma) {
                this.drawHorizontalLine(x1, x2, yScale(mean + stdDev), settings.sigma1Color, "dotted", 1, "sigma1u");
                this.drawHorizontalLine(x1, x2, yScale(mean - stdDev), settings.sigma1Color, "dotted", 1, "sigma1l");
            }
        } else {
            this.phaseStats.forEach((stats) => {
                const startX = xScale(this.dataPoints[stats.startIndex].category) - 20;
                const endX = xScale(this.dataPoints[stats.endIndex].category) + 20;

                if (settings.show3Sigma) {
                    this.drawHorizontalLine(startX, endX, yScale(stats.mean), settings.clColor, settings.clStyle, settings.clWidth, "cl");
                    this.drawHorizontalLine(startX, endX, yScale(stats.ucl), settings.uclColor, settings.uclStyle, settings.uclWidth, "ucl");
                    this.drawHorizontalLine(startX, endX, yScale(stats.lcl), settings.lclColor, settings.lclStyle, settings.lclWidth, "lcl");
                }

                if (settings.show2Sigma) {
                    this.drawHorizontalLine(startX, endX, yScale(stats.sigma2Upper), settings.sigma2Color, "dotted", 1, "sigma2u");
                    this.drawHorizontalLine(startX, endX, yScale(stats.sigma2Lower), settings.sigma2Color, "dotted", 1, "sigma2l");
                }

                if (settings.show1Sigma) {
                    this.drawHorizontalLine(startX, endX, yScale(stats.sigma1Upper), settings.sigma1Color, "dotted", 1, "sigma1u");
                    this.drawHorizontalLine(startX, endX, yScale(stats.sigma1Lower), settings.sigma1Color, "dotted", 1, "sigma1l");
                }
            });
        }
    }

    private drawHorizontalLine(x1: number, x2: number, y: number, color: string, style: string, width: number, className: string): void {
        this.chartGroup.append("line")
            .classed(className, true)
            .attr("x1", x1)
            .attr("x2", x2)
            .attr("y1", y)
            .attr("y2", y)
            .attr("stroke", color)
            .attr("stroke-width", width)
            .attr("stroke-dasharray", getDashArray(style));
    }

    private renderSpecLimits(xScale: d3.ScalePoint<string>, yScale: d3.ScaleLinear<number, number>, width: number): void {
        if (!this.settings.specLimits.show) return;

        const settings = this.settings.specLimits;
        const xRange = xScale.range();
        const x1 = xRange[0] - 20;
        const x2 = xRange[1] + 20;

        if (settings.uslValue !== 0) {
            this.drawHorizontalLine(x1, x2, yScale(settings.uslValue), settings.uslColor, settings.uslStyle, settings.uslWidth, "usl");
        }

        if (settings.lslValue !== 0) {
            this.drawHorizontalLine(x1, x2, yScale(settings.lslValue), settings.lslColor, settings.lslStyle, settings.lslWidth, "lsl");
        }
    }

    private renderTargetLine(xScale: d3.ScalePoint<string>, yScale: d3.ScaleLinear<number, number>, width: number): void {
        if (!this.settings.targetLine.show) return;

        const settings = this.settings.targetLine;
        if (settings.targetValue === 0) return;

        const xRange = xScale.range();
        const x1 = xRange[0] - 20;
        const x2 = xRange[1] + 20;

        this.drawHorizontalLine(x1, x2, yScale(settings.targetValue), settings.color, settings.style, settings.width, "target");
    }

    private renderDataLine(xScale: d3.ScalePoint<string>, yScale: d3.ScaleLinear<number, number>): void {
        if (!this.settings.dataPoints.showLine) return;

        const line = d3.line<DataPoint>()
            .x(d => xScale(d.category))
            .y(d => yScale(d.value))
            .curve(d3.curveLinear);

        this.chartGroup.append("path")
            .datum(this.dataPoints)
            .attr("class", "data-line")
            .attr("d", line)
            .attr("fill", "none")
            .attr("stroke", this.settings.dataPoints.lineColor)
            .attr("stroke-width", this.settings.dataPoints.lineWidth);
    }

    private renderDataPoints(xScale: d3.ScalePoint<string>, yScale: d3.ScaleLinear<number, number>): void {
        if (!this.settings.dataPoints.show) return;

        const settings = this.settings.dataPoints;

        this.chartGroup.selectAll(".data-point")
            .data(this.dataPoints)
            .enter()
            .append("circle")
            .classed("data-point", true)
            .attr("cx", d => xScale(d.category))
            .attr("cy", d => yScale(d.value))
            .attr("r", settings.pointSize)
            .attr("fill", d => d.isOOC ? settings.outControlColor : settings.inControlColor)
            .attr("stroke", d => d.isOOC ? "#000" : "none")
            .attr("stroke-width", d => d.isOOC ? 1 : 0)
            .on("click", (event, d) => {
                this.selectionManager.select(d.identity, true);
                event.stopPropagation();
            });
    }

    private renderLineLabels(xScale: d3.ScalePoint<string>, yScale: d3.ScaleLinear<number, number>, width: number): void {
        if (!this.settings.lineLabels.show) return;

        const settings = this.settings.lineLabels;
        const xRange = xScale.range();
        const labelX = xRange[1] + 5;

        if (this.phaseStats.length <= 1 || !this.settings.phaseControl.recalculatePerPhase) {
            const mean = this.overallStats.mean;
            const stdDev = this.overallStats.stdDev;

            this.addLineLabel(labelX, yScale(mean), "CL", mean, settings);
            this.addLineLabel(labelX, yScale(mean + 3 * stdDev), "UCL", mean + 3 * stdDev, settings);
            this.addLineLabel(labelX, yScale(mean - 3 * stdDev), "LCL", mean - 3 * stdDev, settings);
            this.addLineLabel(labelX, yScale(mean + 2 * stdDev), "+2σ", mean + 2 * stdDev, settings);
            this.addLineLabel(labelX, yScale(mean - 2 * stdDev), "-2σ", mean - 2 * stdDev, settings);
            this.addLineLabel(labelX, yScale(mean + stdDev), "+1σ", mean + stdDev, settings);
            this.addLineLabel(labelX, yScale(mean - stdDev), "-1σ", mean - stdDev, settings);
        } else {
            this.phaseStats.forEach((stats) => {
                const midIdx = Math.floor((stats.startIndex + stats.endIndex) / 2);
                const lx = xScale(this.dataPoints[midIdx].category) + 5;

                this.addLineLabel(lx, yScale(stats.mean), "CL", stats.mean, settings);
                this.addLineLabel(lx, yScale(stats.ucl), "UCL", stats.ucl, settings);
                this.addLineLabel(lx, yScale(stats.lcl), "LCL", stats.lcl, settings);
            });
        }

        const usl = this.getEffectiveUSL();
        const lsl = this.getEffectiveLSL();
        if (usl !== null) this.addLineLabel(labelX, yScale(usl), "USL", usl, settings);
        if (lsl !== null) this.addLineLabel(labelX, yScale(lsl), "LSL", lsl, settings);

        const target = this.getEffectiveTarget();
        if (target !== null) this.addLineLabel(labelX, yScale(target), "T", target, settings);
    }

    private addLineLabel(x: number, y: number, label: string, value: number, settings: { fontSize: number, showValues: boolean }): void {
        const text = settings.showValues ? `${label}: ${value.toFixed(2)}` : label;

        this.chartGroup.append("text")
            .classed("line-label", true)
            .attr("x", x)
            .attr("y", y + 4)
            .attr("fill", "#333")
            .style("font-size", `${settings.fontSize}px`)
            .text(text);
    }

    private renderStatisticsPanel(width: number, height: number): void {
        if (!this.settings.statistics.show) return;

        const stats = this.settings.statistics;
        const lines: string[] = [];

        if (stats.showN) lines.push(`N = ${this.overallStats.n}`);
        if (stats.showMean) lines.push(`Mean = ${this.overallStats.mean.toFixed(4)}`);
        if (stats.showStdDev) lines.push(`σ = ${this.overallStats.stdDev.toFixed(4)}`);
        if (stats.showRange) lines.push(`Range = ${this.overallStats.range.toFixed(4)}`);
        if (stats.showCp && this.overallStats.cp !== null) lines.push(`Cp = ${this.overallStats.cp.toFixed(3)}`);
        if (stats.showCpk && this.overallStats.cpk !== null) lines.push(`Cpk = ${this.overallStats.cpk.toFixed(3)}`);
        if (stats.showPp && this.overallStats.pp !== null) lines.push(`Pp = ${this.overallStats.pp.toFixed(3)}`);
        if (stats.showPpk && this.overallStats.ppk !== null) lines.push(`Ppk = ${this.overallStats.ppk.toFixed(3)}`);
        if (stats.showOOCCount) lines.push(`OOC = ${this.overallStats.oocCount}`);
        if (stats.showOOCPercent) lines.push(`OOC% = ${this.overallStats.oocPercent.toFixed(1)}%`);

        if (lines.length === 0) return;

        const padding = 8;
        const lineHeight = stats.fontSize + 4;
        const boxWidth = 140;
        const boxHeight = lines.length * lineHeight + padding * 2;

        let x: number;
        let y: number;

        switch (stats.position) {
            case "topLeft":
                x = 5;
                y = 5;
                break;
            case "bottomLeft":
                x = 5;
                y = height - boxHeight - 5;
                break;
            case "bottomRight":
                x = width - boxWidth - 5;
                y = height - boxHeight - 5;
                break;
            default:
                x = width - boxWidth - 5;
                y = 5;
        }

        this.chartGroup.append("rect")
            .classed("stats-box", true)
            .attr("x", x)
            .attr("y", y)
            .attr("width", boxWidth)
            .attr("height", boxHeight)
            .attr("fill", "rgba(255,255,255,0.9)")
            .attr("stroke", "#ccc")
            .attr("stroke-width", 1)
            .attr("rx", 3);

        lines.forEach((line, i) => {
            this.chartGroup.append("text")
                .classed("stats-text", true)
                .attr("x", x + padding)
                .attr("y", y + padding + (i + 1) * lineHeight)
                .attr("fill", "#333")
                .style("font-size", `${stats.fontSize}px`)
                .text(line);
        });
    }

    private clearChart(): void {
        this.chartGroup.selectAll("*").remove();
    }

    public enumerateObjectInstances(options: EnumerateVisualObjectInstancesOptions): VisualObjectInstanceEnumeration {
        const settings = this.settings || DefaultSettings;
        const objectName = options.objectName;
        const instances: powerbi.VisualObjectInstance[] = [];

        switch (objectName) {
            case "controlLines":
                instances.push({
                    objectName,
                    properties: {
                        show: settings.controlLines.show,
                        show1Sigma: settings.controlLines.show1Sigma,
                        show2Sigma: settings.controlLines.show2Sigma,
                        show3Sigma: settings.controlLines.show3Sigma,
                        clColor: settings.controlLines.clColor,
                        clStyle: settings.controlLines.clStyle,
                        clWidth: settings.controlLines.clWidth,
                        uclColor: settings.controlLines.uclColor,
                        uclStyle: settings.controlLines.uclStyle,
                        uclWidth: settings.controlLines.uclWidth,
                        lclColor: settings.controlLines.lclColor,
                        lclStyle: settings.controlLines.lclStyle,
                        lclWidth: settings.controlLines.lclWidth,
                        sigma1Color: settings.controlLines.sigma1Color,
                        sigma2Color: settings.controlLines.sigma2Color
                    },
                    selector: null
                });
                break;
            case "specLimits":
                instances.push({
                    objectName,
                    properties: {
                        show: settings.specLimits.show,
                        uslValue: settings.specLimits.uslValue,
                        uslColor: settings.specLimits.uslColor,
                        uslStyle: settings.specLimits.uslStyle,
                        uslWidth: settings.specLimits.uslWidth,
                        lslValue: settings.specLimits.lslValue,
                        lslColor: settings.specLimits.lslColor,
                        lslStyle: settings.specLimits.lslStyle,
                        lslWidth: settings.specLimits.lslWidth
                    },
                    selector: null
                });
                break;
            case "targetLine":
                instances.push({
                    objectName,
                    properties: {
                        show: settings.targetLine.show,
                        targetValue: settings.targetLine.targetValue,
                        color: settings.targetLine.color,
                        style: settings.targetLine.style,
                        width: settings.targetLine.width
                    },
                    selector: null
                });
                break;
            case "dataPoints":
                instances.push({
                    objectName,
                    properties: {
                        show: settings.dataPoints.show,
                        inControlColor: settings.dataPoints.inControlColor,
                        outControlColor: settings.dataPoints.outControlColor,
                        pointSize: settings.dataPoints.pointSize,
                        showLine: settings.dataPoints.showLine,
                        lineColor: settings.dataPoints.lineColor,
                        lineWidth: settings.dataPoints.lineWidth
                    },
                    selector: null
                });
                break;
            case "phaseControl":
                instances.push({
                    objectName,
                    properties: {
                        show: settings.phaseControl.show,
                        separatorColor: settings.phaseControl.separatorColor,
                        separatorStyle: settings.phaseControl.separatorStyle,
                        separatorWidth: settings.phaseControl.separatorWidth,
                        showPhaseLabels: settings.phaseControl.showPhaseLabels,
                        phaseLabelFontSize: settings.phaseControl.phaseLabelFontSize,
                        recalculatePerPhase: settings.phaseControl.recalculatePerPhase
                    },
                    selector: null
                });
                break;
            case "oocRules":
                instances.push({
                    objectName,
                    properties: {
                        rule1: settings.oocRules.rule1,
                        rule2: settings.oocRules.rule2,
                        rule3: settings.oocRules.rule3,
                        rule4: settings.oocRules.rule4,
                        rule5: settings.oocRules.rule5,
                        rule6: settings.oocRules.rule6
                    },
                    selector: null
                });
                break;
            case "statistics":
                instances.push({
                    objectName,
                    properties: {
                        show: settings.statistics.show,
                        position: settings.statistics.position,
                        fontSize: settings.statistics.fontSize,
                        showMean: settings.statistics.showMean,
                        showStdDev: settings.statistics.showStdDev,
                        showCp: settings.statistics.showCp,
                        showCpk: settings.statistics.showCpk,
                        showPp: settings.statistics.showPp,
                        showPpk: settings.statistics.showPpk,
                        showOOCCount: settings.statistics.showOOCCount,
                        showOOCPercent: settings.statistics.showOOCPercent,
                        showN: settings.statistics.showN,
                        showRange: settings.statistics.showRange
                    },
                    selector: null
                });
                break;
            case "axis":
                instances.push({
                    objectName,
                    properties: {
                        showXAxis: settings.axis.showXAxis,
                        showYAxis: settings.axis.showYAxis,
                        axisFontSize: settings.axis.axisFontSize,
                        axisColor: settings.axis.axisColor,
                        xLabel: settings.axis.xLabel,
                        yLabel: settings.axis.yLabel
                    },
                    selector: null
                });
                break;
            case "zones":
                instances.push({
                    objectName,
                    properties: {
                        show: settings.zones.show,
                        zoneAColor: settings.zones.zoneAColor,
                        zoneBColor: settings.zones.zoneBColor,
                        zoneCColor: settings.zones.zoneCColor,
                        zoneOpacity: settings.zones.zoneOpacity
                    },
                    selector: null
                });
                break;
            case "lineLabels":
                instances.push({
                    objectName,
                    properties: {
                        show: settings.lineLabels.show,
                        fontSize: settings.lineLabels.fontSize,
                        showValues: settings.lineLabels.showValues
                    },
                    selector: null
                });
                break;
        }

        return instances;
    }
}
