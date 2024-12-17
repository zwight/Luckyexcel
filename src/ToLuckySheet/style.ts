import { LuckySheetborderInfoCellValueStyle, LuckySheetborderInfoCellForImp, LuckySheetborderInfoCellValue, LuckySheetCelldataBase, LuckySheetCelldataValue, LuckySheetCellFormat, LuckyInlineString } from "./LuckyBase";
import { ReadXml, Element, IStyleCollections, getColor, getlineStringAttr } from "./ReadXml";
import { ST_CellType, indexedColors, OEM_CHARSET, borderTypes, fontFamilys } from "../common/constant"
import { ABCToNumber, isfreezonFuc } from "../common/method";

export enum AbsoluteRefType {
    NONE = 0,
    ROW = 1,
    COLUMN = 2,
    ALL = 3
}
export interface IRange {
    startRow: number;
    startColumn: number;
    endRow: number;
    endColumn: number;
    startAbsoluteRefType: AbsoluteRefType,
    endAbsoluteRefType: AbsoluteRefType,
    rangeType: number
}
export function getBorderInfo(borders: Element[], styles: IStyleCollections): LuckySheetborderInfoCellValueStyle {
    if (borders == null) {
        return null;
    }

    let border = borders[0], attrList = border.attributeList;
    let style: string = attrList.style;
    if (style == null || style == "none") {
        return null;
    }

    let colors = border.getInnerElements("color");
    let colorRet = "#000000";
    if (colors != null) {
        let color = colors[0];
        colorRet = getColor(color, styles, "b");
        if (colorRet == null) {
            colorRet = "#000000";
        }
    }

    let ret = new LuckySheetborderInfoCellValueStyle();
    ret.style = borderTypes[style];
    ret.color = colorRet;

    return ret;
}

export function handleBorder(border: Element, styles: IStyleCollections) {
    const borderCellValue = new LuckySheetborderInfoCellValue();
    let isAdd = false;
    if (!border) {
        return {
            borderCellValue,
            isAdd
        }
    }
    let lefts = border.getInnerElements("left");
    let rights = border.getInnerElements("right");
    let tops = border.getInnerElements("top");
    let bottoms = border.getInnerElements("bottom");
    let diagonals = border.getInnerElements("diagonal");

    let starts = border.getInnerElements("start");
    let ends = border.getInnerElements("end");

    let left = getBorderInfo(lefts, styles);
    let right = getBorderInfo(rights, styles);
    let top = getBorderInfo(tops, styles);
    let bottom = getBorderInfo(bottoms, styles);
    let diagonal = getBorderInfo(diagonals, styles);

    let start = getBorderInfo(starts, styles);
    let end = getBorderInfo(ends, styles);

    if (start != null && start.color != null) {
        borderCellValue.l = start;
        isAdd = true;
    }

    if (end != null && end.color != null) {
        borderCellValue.r = end;
        isAdd = true;
    }

    if (left != null && left.color != null) {
        borderCellValue.l = left;
        isAdd = true;
    }

    if (right != null && right.color != null) {
        borderCellValue.r = right;
        isAdd = true;
    }

    if (top != null && top.color != null) {
        borderCellValue.t = top;
        isAdd = true;
    }

    if (bottom != null && bottom.color != null) {
        borderCellValue.b = bottom;
        isAdd = true;
    }
    return {
        borderCellValue,
        isAdd
    }
}

export function getBackgroundByFill(fill: Element, styles: IStyleCollections): string | null {
    let patternFills = fill.getInnerElements("patternFill");
    if (patternFills != null) {
        let patternFill = patternFills[0];
        let fgColors = patternFill.getInnerElements("fgColor");
        let bgColors = patternFill.getInnerElements("bgColor");
        let fg, bg;
        if (fgColors != null) {
            let fgColor = fgColors[0];
            fg = getColor(fgColor, styles);
        }

        if (bgColors != null) {
            let bgColor = bgColors[0];
            bg = getColor(bgColor, styles);
        }
        // console.log(fgColors,bgColors,clrScheme);
        if (fg != null) {
            return fg;
        }
        else if (bg != null) {
            return bg;
        }
    }
    else {
        let gradientfills = fill.getInnerElements("gradientFill");
        if (gradientfills != null) {
            //graient color fill handler

            return null;
        }
    }
}

export function getFontStyle(font: Element, styles: IStyleCollections) {
    let familyFont = null;
    const cellValue = new LuckySheetCelldataValue();
    let sz = font.getInnerElements("sz");//font size
    let colors = font.getInnerElements("color");//font color
    let family = font.getInnerElements("name");//font family
    let familyOverrides = font.getInnerElements("family");//font family will be overrided by name
    let charset = font.getInnerElements("charset");//font charset
    let bolds = font.getInnerElements("b");//font bold
    let italics = font.getInnerElements("i");//font italic
    let strikes = font.getInnerElements("strike");//font italic
    let underlines = font.getInnerElements("u");//font italic

    if (sz != null && sz.length > 0) {
        let fs = sz[0].attributeList.val;
        if (fs != null) {
            cellValue.fs = parseInt(fs);
        }
    }

    if (colors != null && colors.length > 0) {
        let color = colors[0];
        let fc = getColor(color, styles, "t");
        if (fc != null) {
            cellValue.fc = fc;
        }
    }


    if (familyOverrides != null && familyOverrides.length > 0) {
        let val = familyOverrides[0].attributeList.val;
        if (val != null) {
            familyFont = fontFamilys[val];
        }
    }

    if (family != null && family.length > 0) {
        let val = family[0].attributeList.val;
        if (val != null) {
            cellValue.ff = val;
        }
    }


    if (bolds != null && bolds.length > 0) {
        let bold = bolds[0].attributeList.val;
        if (bold == "0") {
            cellValue.bl = 0;
        }
        else {
            cellValue.bl = 1;
        }
    }

    if (italics != null && italics.length > 0) {
        let italic = italics[0].attributeList.val;
        if (italic == "0") {
            cellValue.it = 0;
        }
        else {
            cellValue.it = 1;
        }
    }

    if (strikes != null && strikes.length > 0) {
        let strike = strikes[0].attributeList.val;
        if (strike == "0") {
            cellValue.cl = 0;
        }
        else {
            cellValue.cl = 1;
        }
    }

    if (underlines != null && underlines.length > 0) {
        let underline = underlines[0].attributeList.val;
        if (underline == "single") {
            cellValue.un = 1;
        }
        else if (underline == "double") {
            cellValue.un = 2;
        }
        else if (underline == "singleAccounting") {
            cellValue.un = 3;
        }
        else if (underline == "doubleAccounting") {
            cellValue.un = 4;
        }
        else {
            cellValue.un = 0;
        }
    }

    return {
        cellValue,
        familyFont
    }
}

export const handleRanges = (sqref: string) => {
    const list = sqref.split(' ');
    return list.map(d => {
        const rangetxt = d.split(':')

        const startRow = parseInt(rangetxt[0].replace(/[^0-9]/g, "")) - 1;
        const startColumn = ABCToNumber(rangetxt[0].replace(/[^A-Za-z]/g, ""));
        const startFreezon = isfreezonFuc(rangetxt[0]);

        const endRow = rangetxt.length === 1 ? startRow : (parseInt(rangetxt[1].replace(/[^0-9]/g, "")) - 1);
        const endColumn = rangetxt.length === 1 ? startColumn : ABCToNumber(rangetxt[1].replace(/[^A-Za-z]/g, ""));
        const endFreezon = rangetxt.length === 1 ? startFreezon : isfreezonFuc(rangetxt[1]);

        const handleType = (freezon: boolean[]) => {
            if (freezon[0] === true && freezon[1] === true) {
                return AbsoluteRefType.ALL;
            }
            if (freezon[0] === true) {
                return AbsoluteRefType.ROW
            }
            if (freezon[1] === true) {
                return AbsoluteRefType.COLUMN
            }
            return AbsoluteRefType.NONE
        }
        return {
            startRow,
            startColumn,
            endRow,
            endColumn,
            startAbsoluteRefType: handleType(startFreezon),
            endAbsoluteRefType: handleType(endFreezon),
            rangeType: 0
        }
    })
}