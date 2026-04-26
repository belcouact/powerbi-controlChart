import powerbi from "powerbi-visuals-api";

export interface LineStyle {
    solid: string;
    dashed: string;
    dotted: string;
}

export const LineStyleValues: LineStyle = {
    solid: "solid",
    dashed: "dashed",
    dotted: "dotted"
};

export function getDashArray(style: string): string {
    switch (style) {
        case "dashed": return "8,4";
        case "dotted": return "2,4";
        default: return "none";
    }
}

export interface ControlLinesSettings {
    show: boolean;
    show1Sigma: boolean;
    show2Sigma: boolean;
    show3Sigma: boolean;
    clColor: string;
    clStyle: string;
    clWidth: number;
    uclColor: string;
    uclStyle: string;
    uclWidth: number;
    lclColor: string;
    lclStyle: string;
    lclWidth: number;
    sigma1Color: string;
    sigma2Color: string;
}

export interface SpecLimitsSettings {
    show: boolean;
    uslValue: number;
    uslColor: string;
    uslStyle: string;
    uslWidth: number;
    lslValue: number;
    lslColor: string;
    lslStyle: string;
    lslWidth: number;
}

export interface TargetLineSettings {
    show: boolean;
    targetValue: number;
    color: string;
    style: string;
    width: number;
}

export interface DataPointsSettings {
    show: boolean;
    inControlColor: string;
    outControlColor: string;
    pointSize: number;
    showLine: boolean;
    lineColor: string;
    lineWidth: number;
}

export interface PhaseControlSettings {
    show: boolean;
    separatorColor: string;
    separatorStyle: string;
    separatorWidth: number;
    showPhaseLabels: boolean;
    phaseLabelFontSize: number;
    recalculatePerPhase: boolean;
}

export interface OOCRulesSettings {
    rule1: boolean;
    rule2: boolean;
    rule3: boolean;
    rule4: boolean;
    rule5: boolean;
    rule6: boolean;
}

export interface StatisticsSettings {
    show: boolean;
    position: string;
    fontSize: number;
    showMean: boolean;
    showStdDev: boolean;
    showCp: boolean;
    showCpk: boolean;
    showPp: boolean;
    showPpk: boolean;
    showOOCCount: boolean;
    showOOCPercent: boolean;
    showN: boolean;
    showRange: boolean;
}

export interface AxisSettings {
    showXAxis: boolean;
    showYAxis: boolean;
    axisFontSize: number;
    axisColor: string;
    xLabel: string;
    yLabel: string;
}

export interface ZonesSettings {
    show: boolean;
    zoneAColor: string;
    zoneBColor: string;
    zoneCColor: string;
    zoneOpacity: number;
}

export interface LineLabelsSettings {
    show: boolean;
    fontSize: number;
    showValues: boolean;
}

export interface VisualSettings {
    controlLines: ControlLinesSettings;
    specLimits: SpecLimitsSettings;
    targetLine: TargetLineSettings;
    dataPoints: DataPointsSettings;
    phaseControl: PhaseControlSettings;
    oocRules: OOCRulesSettings;
    statistics: StatisticsSettings;
    axis: AxisSettings;
    zones: ZonesSettings;
    lineLabels: LineLabelsSettings;
}

export const DefaultSettings: VisualSettings = {
    controlLines: {
        show: true,
        show1Sigma: true,
        show2Sigma: true,
        show3Sigma: true,
        clColor: "#1B6EF3",
        clStyle: "solid",
        clWidth: 2,
        uclColor: "#E03E2D",
        uclStyle: "dashed",
        uclWidth: 2,
        lclColor: "#E03E2D",
        lclStyle: "dashed",
        lclWidth: 2,
        sigma1Color: "#A0A0A0",
        sigma2Color: "#C0C0C0"
    },
    specLimits: {
        show: false,
        uslValue: 0,
        uslColor: "#8B0000",
        uslStyle: "dashed",
        uslWidth: 2,
        lslValue: 0,
        lslColor: "#8B0000",
        lslStyle: "dashed",
        lslWidth: 2
    },
    targetLine: {
        show: false,
        targetValue: 0,
        color: "#00B294",
        style: "dashed",
        width: 2
    },
    dataPoints: {
        show: true,
        inControlColor: "#1B6EF3",
        outControlColor: "#E03E2D",
        pointSize: 5,
        showLine: true,
        lineColor: "#666666",
        lineWidth: 1.5
    },
    phaseControl: {
        show: true,
        separatorColor: "#888888",
        separatorStyle: "dashed",
        separatorWidth: 1.5,
        showPhaseLabels: true,
        phaseLabelFontSize: 11,
        recalculatePerPhase: true
    },
    oocRules: {
        rule1: true,
        rule2: true,
        rule3: true,
        rule4: true,
        rule5: true,
        rule6: true
    },
    statistics: {
        show: true,
        position: "topRight",
        fontSize: 11,
        showMean: true,
        showStdDev: true,
        showCp: true,
        showCpk: true,
        showPp: true,
        showPpk: true,
        showOOCCount: true,
        showOOCPercent: true,
        showN: true,
        showRange: true
    },
    axis: {
        showXAxis: true,
        showYAxis: true,
        axisFontSize: 10,
        axisColor: "#333333",
        xLabel: "",
        yLabel: ""
    },
    zones: {
        show: false,
        zoneAColor: "#FDECEA",
        zoneBColor: "#FFF4E5",
        zoneCColor: "#E8F5E9",
        zoneOpacity: 0.4
    },
    lineLabels: {
        show: true,
        fontSize: 10,
        showValues: true
    }
};

export function getSettings(objects: powerbi.DataViewObjects): VisualSettings {
    return {
        controlLines: {
            show: getValue<boolean>(objects, "controlLines", "show", DefaultSettings.controlLines.show),
            show1Sigma: getValue<boolean>(objects, "controlLines", "show1Sigma", DefaultSettings.controlLines.show1Sigma),
            show2Sigma: getValue<boolean>(objects, "controlLines", "show2Sigma", DefaultSettings.controlLines.show2Sigma),
            show3Sigma: getValue<boolean>(objects, "controlLines", "show3Sigma", DefaultSettings.controlLines.show3Sigma),
            clColor: getValue<string>(objects, "controlLines", "clColor", DefaultSettings.controlLines.clColor),
            clStyle: getValue<string>(objects, "controlLines", "clStyle", DefaultSettings.controlLines.clStyle),
            clWidth: getValue<number>(objects, "controlLines", "clWidth", DefaultSettings.controlLines.clWidth),
            uclColor: getValue<string>(objects, "controlLines", "uclColor", DefaultSettings.controlLines.uclColor),
            uclStyle: getValue<string>(objects, "controlLines", "uclStyle", DefaultSettings.controlLines.uclStyle),
            uclWidth: getValue<number>(objects, "controlLines", "uclWidth", DefaultSettings.controlLines.uclWidth),
            lclColor: getValue<string>(objects, "controlLines", "lclColor", DefaultSettings.controlLines.lclColor),
            lclStyle: getValue<string>(objects, "controlLines", "lclStyle", DefaultSettings.controlLines.lclStyle),
            lclWidth: getValue<number>(objects, "controlLines", "lclWidth", DefaultSettings.controlLines.lclWidth),
            sigma1Color: getValue<string>(objects, "controlLines", "sigma1Color", DefaultSettings.controlLines.sigma1Color),
            sigma2Color: getValue<string>(objects, "controlLines", "sigma2Color", DefaultSettings.controlLines.sigma2Color)
        },
        specLimits: {
            show: getValue<boolean>(objects, "specLimits", "show", DefaultSettings.specLimits.show),
            uslValue: getValue<number>(objects, "specLimits", "uslValue", DefaultSettings.specLimits.uslValue),
            uslColor: getValue<string>(objects, "specLimits", "uslColor", DefaultSettings.specLimits.uslColor),
            uslStyle: getValue<string>(objects, "specLimits", "uslStyle", DefaultSettings.specLimits.uslStyle),
            uslWidth: getValue<number>(objects, "specLimits", "uslWidth", DefaultSettings.specLimits.uslWidth),
            lslValue: getValue<number>(objects, "specLimits", "lslValue", DefaultSettings.specLimits.lslValue),
            lslColor: getValue<string>(objects, "specLimits", "lslColor", DefaultSettings.specLimits.lslColor),
            lslStyle: getValue<string>(objects, "specLimits", "lslStyle", DefaultSettings.specLimits.lslStyle),
            lslWidth: getValue<number>(objects, "specLimits", "lslWidth", DefaultSettings.specLimits.lslWidth)
        },
        targetLine: {
            show: getValue<boolean>(objects, "targetLine", "show", DefaultSettings.targetLine.show),
            targetValue: getValue<number>(objects, "targetLine", "targetValue", DefaultSettings.targetLine.targetValue),
            color: getValue<string>(objects, "targetLine", "color", DefaultSettings.targetLine.color),
            style: getValue<string>(objects, "targetLine", "style", DefaultSettings.targetLine.style),
            width: getValue<number>(objects, "targetLine", "width", DefaultSettings.targetLine.width)
        },
        dataPoints: {
            show: getValue<boolean>(objects, "dataPoints", "show", DefaultSettings.dataPoints.show),
            inControlColor: getValue<string>(objects, "dataPoints", "inControlColor", DefaultSettings.dataPoints.inControlColor),
            outControlColor: getValue<string>(objects, "dataPoints", "outControlColor", DefaultSettings.dataPoints.outControlColor),
            pointSize: getValue<number>(objects, "dataPoints", "pointSize", DefaultSettings.dataPoints.pointSize),
            showLine: getValue<boolean>(objects, "dataPoints", "showLine", DefaultSettings.dataPoints.showLine),
            lineColor: getValue<string>(objects, "dataPoints", "lineColor", DefaultSettings.dataPoints.lineColor),
            lineWidth: getValue<number>(objects, "dataPoints", "lineWidth", DefaultSettings.dataPoints.lineWidth)
        },
        phaseControl: {
            show: getValue<boolean>(objects, "phaseControl", "show", DefaultSettings.phaseControl.show),
            separatorColor: getValue<string>(objects, "phaseControl", "separatorColor", DefaultSettings.phaseControl.separatorColor),
            separatorStyle: getValue<string>(objects, "phaseControl", "separatorStyle", DefaultSettings.phaseControl.separatorStyle),
            separatorWidth: getValue<number>(objects, "phaseControl", "separatorWidth", DefaultSettings.phaseControl.separatorWidth),
            showPhaseLabels: getValue<boolean>(objects, "phaseControl", "showPhaseLabels", DefaultSettings.phaseControl.showPhaseLabels),
            phaseLabelFontSize: getValue<number>(objects, "phaseControl", "phaseLabelFontSize", DefaultSettings.phaseControl.phaseLabelFontSize),
            recalculatePerPhase: getValue<boolean>(objects, "phaseControl", "recalculatePerPhase", DefaultSettings.phaseControl.recalculatePerPhase)
        },
        oocRules: {
            rule1: getValue<boolean>(objects, "oocRules", "rule1", DefaultSettings.oocRules.rule1),
            rule2: getValue<boolean>(objects, "oocRules", "rule2", DefaultSettings.oocRules.rule2),
            rule3: getValue<boolean>(objects, "oocRules", "rule3", DefaultSettings.oocRules.rule3),
            rule4: getValue<boolean>(objects, "oocRules", "rule4", DefaultSettings.oocRules.rule4),
            rule5: getValue<boolean>(objects, "oocRules", "rule5", DefaultSettings.oocRules.rule5),
            rule6: getValue<boolean>(objects, "oocRules", "rule6", DefaultSettings.oocRules.rule6)
        },
        statistics: {
            show: getValue<boolean>(objects, "statistics", "show", DefaultSettings.statistics.show),
            position: getValue<string>(objects, "statistics", "position", DefaultSettings.statistics.position),
            fontSize: getValue<number>(objects, "statistics", "fontSize", DefaultSettings.statistics.fontSize),
            showMean: getValue<boolean>(objects, "statistics", "showMean", DefaultSettings.statistics.showMean),
            showStdDev: getValue<boolean>(objects, "statistics", "showStdDev", DefaultSettings.statistics.showStdDev),
            showCp: getValue<boolean>(objects, "statistics", "showCp", DefaultSettings.statistics.showCp),
            showCpk: getValue<boolean>(objects, "statistics", "showCpk", DefaultSettings.statistics.showCpk),
            showPp: getValue<boolean>(objects, "statistics", "showPp", DefaultSettings.statistics.showPp),
            showPpk: getValue<boolean>(objects, "statistics", "showPpk", DefaultSettings.statistics.showPpk),
            showOOCCount: getValue<boolean>(objects, "statistics", "showOOCCount", DefaultSettings.statistics.showOOCCount),
            showOOCPercent: getValue<boolean>(objects, "statistics", "showOOCPercent", DefaultSettings.statistics.showOOCPercent),
            showN: getValue<boolean>(objects, "statistics", "showN", DefaultSettings.statistics.showN),
            showRange: getValue<boolean>(objects, "statistics", "showRange", DefaultSettings.statistics.showRange)
        },
        axis: {
            showXAxis: getValue<boolean>(objects, "axis", "showXAxis", DefaultSettings.axis.showXAxis),
            showYAxis: getValue<boolean>(objects, "axis", "showYAxis", DefaultSettings.axis.showYAxis),
            axisFontSize: getValue<number>(objects, "axis", "axisFontSize", DefaultSettings.axis.axisFontSize),
            axisColor: getValue<string>(objects, "axis", "axisColor", DefaultSettings.axis.axisColor),
            xLabel: getValue<string>(objects, "axis", "xLabel", DefaultSettings.axis.xLabel),
            yLabel: getValue<string>(objects, "axis", "yLabel", DefaultSettings.axis.yLabel)
        },
        zones: {
            show: getValue<boolean>(objects, "zones", "show", DefaultSettings.zones.show),
            zoneAColor: getValue<string>(objects, "zones", "zoneAColor", DefaultSettings.zones.zoneAColor),
            zoneBColor: getValue<string>(objects, "zones", "zoneBColor", DefaultSettings.zones.zoneBColor),
            zoneCColor: getValue<string>(objects, "zones", "zoneCColor", DefaultSettings.zones.zoneCColor),
            zoneOpacity: getValue<number>(objects, "zones", "zoneOpacity", DefaultSettings.zones.zoneOpacity)
        },
        lineLabels: {
            show: getValue<boolean>(objects, "lineLabels", "show", DefaultSettings.lineLabels.show),
            fontSize: getValue<number>(objects, "lineLabels", "fontSize", DefaultSettings.lineLabels.fontSize),
            showValues: getValue<boolean>(objects, "lineLabels", "showValues", DefaultSettings.lineLabels.showValues)
        }
    };
}

function getValue<T>(objects: powerbi.DataViewObjects, objectName: string, propertyName: string, defaultValue: T): T {
    if (objects && objects[objectName]) {
        const object = objects[objectName];
        if (object && object[propertyName] !== undefined && object[propertyName] !== null) {
            return <T>object[propertyName];
        }
    }
    return defaultValue;
}
