import {
    BooleanNumber,
    VerticalAlign,
    Nullable,
    IBorderData,
    IBorderStyleData,
    WrapStrategy,
    TextDecoration,
    ThemeColorType,
    HorizontalAlign,
} from '@univerjs/core';
import { isObject } from '../common/method';
import { IluckySheetborderInfoCellForImp, IluckySheetborderInfoCellValueStyle, IluckySheetCelldata } from '../ToLuckySheet/ILuck';

/**
 * 删除对象中含undefined的值
 * @param object
 * @returns
 */
export function removeEmptyAttr(object: any) {
    for (const key in object) {
        if (Object.prototype.hasOwnProperty.call(object, key)) {
            if (object[key] === undefined) {
                delete object[key]; // 删除值为 undefined 的属性
            } else if (isObject(object[key]) && object[key] !== null) {
                removeEmptyAttr(object[key]); // 对子对象递归
            }
        }
    }
    return object;
}
export const handleStyle = (
    row: IluckySheetCelldata,
    borderInfo?: IluckySheetborderInfoCellForImp,
    domContent: boolean = false
) => {
    const { v } = row;
    if (typeof v === 'string' || v === null || v === undefined) {
        return undefined;
    }
    // 0 middle, 1 up, 2 down
    const VerticalAlignMap: any = {
        0: VerticalAlign.MIDDLE,
        1: VerticalAlign.TOP,
        2: VerticalAlign.BOTTOM,
    };
    let border: Nullable<IBorderData> = undefined;
    if (borderInfo?.value && !domContent) {
        const handleBorder: (
            con: IluckySheetborderInfoCellValueStyle
        ) => IBorderStyleData | null = (con: IluckySheetborderInfoCellValueStyle) => {
            if (!con) return null;
            return {
                s: con.style,
                cl: { rgb: con.color, th: ThemeColorType.DARK1 },
            };
        };
        border = {
            t: handleBorder(borderInfo.value?.t),
            r: handleBorder(borderInfo.value?.r),
            b: handleBorder(borderInfo.value?.b),
            l: handleBorder(borderInfo.value?.l),
        };
    }
    const TextWrap: any = {
        0: WrapStrategy.CLIP,
        1: WrapStrategy.OVERFLOW,
        2: WrapStrategy.WRAP,
    };

    let angle = undefined;
    const vtMap: any = {
        1: 45,
        2: 135,
        3: 255,
        4: 90,
        5: 180,
    };
    if (v.tr) angle = vtMap[v.tr];
    if (v.rt) angle = v.rt;
    const unMap: any = {
        0: TextDecoration.DASH,
        1: TextDecoration.SINGLE,
        2: TextDecoration.DOUBLE,
        3: TextDecoration.SINGLE,
        4: TextDecoration.DOUBLE,
    };
    // 0 center, 1 left, 2 right
    const htMap: any = {
        0: HorizontalAlign.CENTER,
        1: HorizontalAlign.LEFT,
        2: HorizontalAlign.RIGHT,
    };
    return {
        // bbl: , // bottomBorerLine
        bd: border, // border
        bg: v.bg !== undefined ? { rgb: v.bg, th: ThemeColorType.DARK1 } : undefined, // background
        bl: v.bl, // bold 0: false 1: true
        cl: v.fc !== undefined ? { rgb: v.fc, th: ThemeColorType.DARK1 } : undefined, // foreground
        ff: v.ff, // fontFamily
        fs: v.fs, // fontSize
        ht: v.ht !== undefined ? htMap[v.ht] : undefined, // horizontalAlignment
        it: v.it, // italic 0: false 1: true
        n: v.ct?.fa !== undefined ? { pattern: v.ct.fa } : undefined, //Numfmt pattern
        // ol: { s: v.cl === 0 ? BooleanNumber.TRUE : BooleanNumber.FALSE}, // overline
        // pd: , // padding
        st:
            v.cl !== undefined
                ? { s: v.cl === 1 ? BooleanNumber.TRUE : BooleanNumber.FALSE }
                : undefined, // strikethrough
        tb: v.tb !== undefined ? TextWrap[v.tb] : undefined, // wrapStrategy
        // td: , // textDirection
        tr:
            angle !== undefined
                ? {
                      a: angle,
                      v: v.tr || v.rt ? BooleanNumber.TRUE : BooleanNumber.FALSE,
                  }
                : undefined, // textRotation
        ul:
            v.un !== undefined
                ? {
                      s: v.un === undefined ? BooleanNumber.FALSE : BooleanNumber.TRUE,
                      t: v.un ? unMap[v.un] : TextDecoration.DASH,
                  }
                : undefined, // underline
        // va: , // (Subscript 下标 /Superscript上标 Text)
        vt: v.vt !== undefined ? VerticalAlignMap[v.vt] : undefined, // verticalAlignment
    };
};