import type { Theme } from "@glideapps/glide-data-grid";
import { getDefaultTheme } from "@glideapps/glide-data-grid";

// Excel for Microsoft 365 look (plan 2.5 / 2.6).
export const FONT_FAMILY = 'Aptos, "Segoe UI Variable", "Segoe UI", Calibri, sans-serif';

/** Base cell font size at 100% zoom, in pixels (11 pt). */
export const BASE_FONT_PX = 14.67;
/** Header font size at 100% zoom, in pixels. */
export const HEADER_FONT_PX = 12;
/** Row height at 100% zoom. */
export const ROW_HEIGHT = 20;

export const EXCEL_GREEN = "#107C41";
export const EXCEL_GREEN_HOVER = "#0F703B";
export const FROZEN_LINE = "#B0B0B0";
export const FILTER_BLUE = "#1E5EFF";

const overrides: Partial<Theme> = {
    accentColor: EXCEL_GREEN,
    accentFg: "#FFFFFF",
    accentLight: "rgba(16,124,65,0.12)",
    textDark: "#000000",
    textMedium: "#555555",
    textLight: "#777777",
    textBubble: "#000000",
    bgIconHeader: "#444444",
    fgIconHeader: "#FFFFFF",
    textHeader: "#444444",
    textHeaderSelected: EXCEL_GREEN_HOVER,
    bgCell: "#FFFFFF",
    bgCellMedium: "#FAFAFB",
    bgHeader: "#F3F3F3",
    bgHeaderHasFocus: "#CAEAD8",
    bgHeaderHovered: "#E6E6E6",
    bgBubble: "#EDEDF3",
    bgBubbleSelected: "#FFFFFF",
    bgSearchResult: "#FFF9E3",
    borderColor: "#E1E1E1",
    horizontalBorderColor: "#E1E1E1",
    drilldownBorder: "rgba(0,0,0,0)",
    linkColor: EXCEL_GREEN,
    cellHorizontalPadding: 4,
    cellVerticalPadding: 3,
    headerIconSize: 16,
    headerFontStyle: `600 ${HEADER_FONT_PX}px`,
    baseFontStyle: `${BASE_FONT_PX}px`,
    markerFontStyle: "9px",
    fontFamily: FONT_FAMILY,
    editorFontSize: `${BASE_FONT_PX}px`,
    lineHeight: 1.2,
};

/** The theme passed to DataEditor at 100% zoom. */
export const excelTheme: Theme = { ...getDefaultTheme(), ...overrides };

/** A theme scaled for a zoom percentage (font, header height handled separately). */
export function scaledTheme(zoomPct: number): Theme {
    const z = zoomPct / 100;
    return {
        ...excelTheme,
        baseFontStyle: `${BASE_FONT_PX * z}px`,
        editorFontSize: `${BASE_FONT_PX * z}px`,
        headerFontStyle: `600 ${HEADER_FONT_PX * z}px`,
    };
}
