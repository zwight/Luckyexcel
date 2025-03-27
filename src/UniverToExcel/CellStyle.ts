import { Style } from "@zwight/exceljs";
import { hex2argb, isUndefined } from "./util";

export function cellStyle(style: any = {}, numFmt: any, isCondition = false): Style {
    return {
        numFmt: numFmt,
        font: fontConvert(style),
        alignment: alignmentConvert(style),
        protection: null,
        border: borderConvert(style.bd),
        fill: fillConvert(style.bg?.rgb, isCondition)
    }
}
export function fontConvert(style: any) {
    const univerToExcel: any = {
        underline: {
            10: 'double',
            12: 'single',
        },
        vertAlign: {
            1: undefined,
            2: 'subscript',
            3: 'superscript'
        }
    }
    return {
        name: style.ff,
        size: style.fs,
        family: 1,
        color: isUndefined(style.cl?.rgb, {argb: hex2argb(style.cl?.rgb)}),
        bold: isUndefined(style.bl, style.bl === 1),
        italic: isUndefined(style.it, style.it === 1),
        underline: isUndefined(style.ul?.s, (style.ul?.s === 1 ? (univerToExcel.underline[style.ul.t] || true) : false)),
        vertAlign: isUndefined(style.va, univerToExcel.vertAlign[style.va]),
        strike: isUndefined(style.st?.s, style.st?.s === 1),
        outline: isUndefined(style.ol?.s, style.ol?.s === 1),
        charset: 134
    }
}
function borderConvert(border: any) {
    if (!border) {
		return null
	}
	const borderStyle: any =  {
        0: 'none',
        1: 'thin',
        2: 'hair',
        3: 'dotted',
        4: 'dashDot', // 'Dashed',
        5: 'dashDot',
        6: 'dashDotDot',
        7: 'double',
        8: 'medium',
        9: 'mediumDashed',
        10: 'mediumDashDot',
        11: 'mediumDashDotDot',
        12: 'slantDashDot',
        13: 'thick'
    }
    const template = (bd: any) => {
        if (!bd) return undefined;
        const st: any = {
            style: borderStyle[bd?.s || 1],
            color: {
                argb: hex2argb(bd?.cl?.rgb || '#d9d9d9')
            }
        }
        return st
    }
    const diagonal = template(border.bl_tr || border.tl_br) || {}
    return {
        top: template(border.t),
        right: template(border.r),
        bottom: template(border.b),
        left: template(border.l),
        diagonal: {
            up: border.bl_tr ? true : false,
            down: border.tl_br ? true : false,
            ...diagonal
        }
    }

}
function alignmentConvert(style: any) {
    const univerToExcel: any = {
        horizontal: {
            0: 'left',
            1: 'left',
            2: 'center',
            3: 'right',
        },
        vertical: {
            1: 'top',
            2: 'middle',
            3: 'bottom',
        },
        wrapText: {
            3: true
        }
    }
    return {
        horizontal: isUndefined(style.ht, univerToExcel.horizontal[style.ht]),
        vertical: isUndefined(style.vt, univerToExcel.vertical[style.vt]),
        wrapText: isUndefined(style.tb, style.tb === 3),
        textRotation: isUndefined(style.tr?.a)
    }
}

function fillConvert(bg: string, isCondition = false) {
    if (!bg) return null;
    if (!bg) return {
        type: 'pattern' as const,
        pattern: 'none' as const,
    };
    const fill: any = {
        type: 'pattern' as const,
        pattern: 'solid' as const,
    }
    if (isCondition) {
        fill.bgColor = { argb: hex2argb(bg, 'FF') }
    } else {
        fill.fgColor = { argb: hex2argb(bg, 'FF') }
    }
    return fill
}