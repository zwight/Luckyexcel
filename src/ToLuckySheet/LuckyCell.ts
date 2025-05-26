import { IluckyImageDefault, IluckySheetborderInfoCellForImp } from "./ILuck";
import { ReadXml, Element, IStyleCollections, getColor, getlineStringAttr } from "./ReadXml";
import { getXmlAttibute, getcellrange, escapeCharacter, isChinese, isJapanese, isKoera, getPxByEMUs } from "../common/method";
import { ST_CellType, cellImagesRels } from "../common/constant"
import { IattributeList } from "../ICommon";
import { LuckySheetborderInfoCellForImp, LuckySheetCelldataBase, LuckySheetCelldataValue, LuckySheetCellFormat, LuckyInlineString } from "./LuckyBase";
import { getBackgroundByFill, getFontStyle, handleBorder } from './style'
import { ImageList } from "./LuckyImage";
import { replaceSpecialWrap } from "../common/utils";

export class LuckySheetCelldata extends LuckySheetCelldataBase {
    _borderObject: IluckySheetborderInfoCellForImp
    _fomulaRef: string
    _formulaSi: string
    _formulaType: string

    private sheetFile: string
    private readXml: ReadXml
    private cell: Element
    private styles: IStyleCollections
    private sharedStrings: Element[]
    private mergeCells: Element[]
    private cellImages: Element[]
    private imageList:ImageList
    private cellSize: { width: number, height: number }
    constructor(
        cell: Element, 
        cellSize: { width: number, height: number },
        styles: IStyleCollections, 
        sharedStrings: Element[], 
        mergeCells: Element[], 
        sheetFile: string, 
        cellImages: Element[], 
        imageList: ImageList, 
        ReadXml: ReadXml
    ) {
        //Private
        super();
        this.cell = cell;
        this.sheetFile = sheetFile;
        this.styles = styles;
        this.sharedStrings = sharedStrings;
        this.readXml = ReadXml;
        this.mergeCells = mergeCells;
        this.cellImages = cellImages;
        this.imageList = imageList;
        this.cellSize = cellSize;

        let attrList = cell.attributeList;
        let r = attrList.r, s = attrList.s, t = attrList.t;
        let range = getcellrange(r);

        this.r = range.row[0];
        this.c = range.column[0];
        this.v = this.generateValue(s, t);

    }

    /**
    * @param s Style index ,start 1
    * @param t Cell type, Optional value is ST_CellType, it's found at constat.ts
    */
    private generateValue(s: string, t: string) {
        let v = this.cell.getInnerElements("v");
        let f = this.cell.getInnerElements("f");

        if (v == null) {
            v = this.cell.getInnerElements("t");
        }

        let cellXfs = this.styles["cellXfs"] as Element[];
        let cellStyleXfs = this.styles["cellStyleXfs"] as Element[];
        let cellStyles = this.styles["cellStyles"] as Element[];
        let fonts = this.styles["fonts"] as Element[];
        let fills = this.styles["fills"] as Element[];
        let borders = this.styles["borders"] as Element[];
        let numfmts = this.styles["numfmts"] as IattributeList;
        let clrScheme = this.styles["clrScheme"] as Element[];

        let sharedStrings = this.sharedStrings;
        let cellValue = new LuckySheetCelldataValue();

        if (f != null) {
            let formula = f[0], attrList = formula.attributeList;
            let t = attrList.t, ref = attrList.ref, si = attrList.si;
            let formulaValue = f[0].value;
            if (t == "shared") {
                this._fomulaRef = ref;
                this._formulaType = t;
                this._formulaSi = si;
            }
            // console.log(ref, t, si);
            if (ref != null || (formulaValue != null && formulaValue.length > 0)) {
                formulaValue = escapeCharacter(formulaValue);
                cellValue.f = formulaValue[0] === '=' ? formulaValue : "=" + formulaValue;
            }

        }

        let familyFont: string = null;
        let quotePrefix;
        if (s != null) {
            let sNum = parseInt(s);
            let cellXf = cellXfs[sNum];
            let xfId = cellXf.attributeList.xfId;

            let numFmtId, fontId, fillId, borderId;
            let horizontal, vertical, wrapText, textRotation, shrinkToFit, indent, applyProtection;

            if (xfId != null) {
                let cellStyleXf = cellStyleXfs[parseInt(xfId)];
                let attrList = cellStyleXf.attributeList;

                let applyNumberFormat = attrList.applyNumberFormat;
                let applyFont = attrList.applyFont;
                let applyFill = attrList.applyFill;
                let applyBorder = attrList.applyBorder;
                let applyAlignment = attrList.applyAlignment;
                // let applyProtection = attrList.applyProtection;

                applyProtection = attrList.applyProtection;
                quotePrefix = attrList.quotePrefix;

                if (applyNumberFormat != "0" && attrList.numFmtId != null) {
                    // if(attrList.numFmtId!="0"){
                    numFmtId = attrList.numFmtId;
                    // }
                }
                if (applyFont != "0" && attrList.fontId != null) {
                    fontId = attrList.fontId;
                }
                if (applyFill != "0" && attrList.fillId != null) {
                    fillId = attrList.fillId;
                }
                if (applyBorder != "0" && attrList.borderId != null) {
                    borderId = attrList.borderId;
                }
                if (applyAlignment != null && applyAlignment != "0") {
                    let alignment = cellStyleXf.getInnerElements("alignment");
                    if (alignment != null) {
                        let attrList = alignment[0].attributeList;
                        if (attrList.horizontal != null) {
                            horizontal = attrList.horizontal;
                        }
                        if (attrList.vertical != null) {
                            vertical = attrList.vertical;
                        }
                        if (attrList.wrapText != null) {
                            wrapText = attrList.wrapText;
                        }
                        if (attrList.textRotation != null) {
                            textRotation = attrList.textRotation;
                        }
                        if (attrList.shrinkToFit != null) {
                            shrinkToFit = attrList.shrinkToFit;
                        }
                        if (attrList.indent != null) {
                            indent = attrList.indent;
                        }
                    }
                }
            }

            let applyNumberFormat = cellXf.attributeList.applyNumberFormat;
            let applyFont = cellXf.attributeList.applyFont;
            let applyFill = cellXf.attributeList.applyFill;
            let applyBorder = cellXf.attributeList.applyBorder;
            let applyAlignment = cellXf.attributeList.applyAlignment;

            if (cellXf.attributeList.applyProtection != null) {
                applyProtection = cellXf.attributeList.applyProtection;
            }

            if (cellXf.attributeList.quotePrefix != null) {
                quotePrefix = cellXf.attributeList.quotePrefix;
            }

            if (applyNumberFormat != "0" && cellXf.attributeList.numFmtId != null) {
                numFmtId = cellXf.attributeList.numFmtId;
            }
            if (applyFont != "0") {
                fontId = cellXf.attributeList.fontId;
            }
            if (applyFill != "0") {
                fillId = cellXf.attributeList.fillId;
            }
            if (applyBorder != "0") {
                borderId = cellXf.attributeList.borderId;
            }
            if (applyAlignment != "0") {
                let alignment = cellXf.getInnerElements("alignment");
                if (alignment != null && alignment.length > 0) {
                    let attrList = alignment[0].attributeList;
                    if (attrList.horizontal != null) {
                        horizontal = attrList.horizontal;
                    }
                    if (attrList.vertical != null) {
                        vertical = attrList.vertical;
                    }
                    if (attrList.wrapText != null) {
                        wrapText = attrList.wrapText;
                    }
                    if (attrList.textRotation != null) {
                        textRotation = attrList.textRotation;
                    }
                    if (attrList.shrinkToFit != null) {
                        shrinkToFit = attrList.shrinkToFit;
                    }
                    if (attrList.indent != null) {
                        indent = attrList.indent;
                    }
                }
            }



            if (numFmtId != undefined) {
                let numf = numfmts[parseInt(numFmtId)];
                let cellFormat = new LuckySheetCellFormat();
                cellFormat.fa = escapeCharacter(numf);
                // console.log(numf, numFmtId, this.v, cellFormat);
                cellFormat.t = t || 'd';
                cellValue.ct = cellFormat;
            }

            if (fillId != undefined) {
                let fillIdNum = parseInt(fillId);
                let fill = fills[fillIdNum];
                // console.log(cellValue.v);
                let bg = getBackgroundByFill(fill, this.styles);
                if (bg != null) {
                    cellValue.bg = bg;
                }
            }


            if (fontId != undefined) {
                let fontIdNum = parseInt(fontId);
                let font = fonts[fontIdNum];
                if (font != null) {
                    const { cellValue: fontStyle, familyFont: family } = getFontStyle(font, this.styles)
                    cellValue = {
                        ...cellValue,
                        ...fontStyle
                    }
                    familyFont = family;
                }
            }

            // vt: number | undefined//Vertical alignment, 0 middle, 1 up, 2 down, alignment
            // ht: number | undefined//Horizontal alignment,0 center, 1 left, 2 right, alignment
            // tr: number | undefined //Text rotation,0: 0、1: 45 、2: -45、3 Vertical text、4: 90 、5: -90, alignment
            // tb: number | undefined //Text wrap,0 truncation, 1 overflow, 2 word wrap, alignment

            if (horizontal != undefined) {//Horizontal alignment
                if (horizontal == "center") {
                    cellValue.ht = 0;
                }
                else if (horizontal == "centerContinuous") {
                    cellValue.ht = 0;//luckysheet unsupport
                }
                else if (horizontal == "left") {
                    cellValue.ht = 1;
                }
                else if (horizontal == "right") {
                    cellValue.ht = 2;
                }
                else if (horizontal == "distributed") {
                    cellValue.ht = 0;//luckysheet unsupport
                }
                else if (horizontal == "fill") {
                    cellValue.ht = 1;//luckysheet unsupport
                }
                else if (horizontal == "general") {
                    cellValue.ht = 1;//luckysheet unsupport
                }
                else if (horizontal == "justify") {
                    cellValue.ht = 0;//luckysheet unsupport
                }
                else {
                    cellValue.ht = 1;
                }
            }

            if (vertical != undefined) {//Vertical alignment
                if (vertical == "bottom") {
                    cellValue.vt = 2;
                }
                else if (vertical == "center") {
                    cellValue.vt = 0;
                }
                else if (vertical == "distributed") {
                    cellValue.vt = 0;//luckysheet unsupport
                }
                else if (vertical == "justify") {
                    cellValue.vt = 0;//luckysheet unsupport
                }
                else if (vertical == "top") {
                    cellValue.vt = 1;
                }
                else {
                    cellValue.vt = 1;
                }
            }
            else {
                //sometimes bottom style is lost after setting it in excel
                //when vertical is undefined set it to 2.
                cellValue.vt = 2;
            }

            if (wrapText != undefined) {
                if (wrapText == "1") {
                    cellValue.tb = 2;
                }
                else {
                    cellValue.tb = 1;
                }
            }
            else {
                cellValue.tb = 1;
            }

            if (textRotation != undefined) {
                // tr: number | undefined //Text rotation,0: 0、1: 45 、2: -45、3 Vertical text、4: 90 、5: -90, alignment
                if (textRotation == "255") {
                    cellValue.tr = 3;
                }
                // else if(textRotation=="45"){
                //     cellValue.tr = 1;
                // }
                // else if(textRotation=="90"){
                //     cellValue.tr = 4;
                // }
                // else if(textRotation=="135"){
                //     cellValue.tr = 2;
                // }
                // else if(textRotation=="180"){
                //     cellValue.tr = 5;
                // }
                else {
                    cellValue.tr = 0;
                    cellValue.rt = parseInt(textRotation);
                }


            }

            if (shrinkToFit != undefined) {//luckysheet unsupport

            }

            if (indent != undefined) {//luckysheet unsupport
                const result = parseInt(indent)
                if (!isNaN(result)) {
                    cellValue.ti = result
                }
            }

            if (borderId != undefined) {
                let borderIdNum = parseInt(borderId);
                let border = borders[borderIdNum];
                // this._borderId = borderIdNum;

                let borderObject = new LuckySheetborderInfoCellForImp();
                borderObject.rangeType = "cell";
                // borderObject.cells = [];

                const { isAdd, borderCellValue } = handleBorder(border, this.styles);

                borderCellValue.row_index = this.r;
                borderCellValue.col_index = this.c;

                if (isAdd) {
                    borderObject.value = borderCellValue;
                    // this.config._borderInfo[borderId] = borderObject;
                    this._borderObject = borderObject;
                }
            }

        }
        else {
            cellValue.tb = 1;
        }

        if (v != null) {
            let value = v[0].value;

            if (/&#\d+;/.test(value)) {
                value = this.htmlDecode(value);
            }

            const generateInlineString = (sharedSI: Element) =>{
                let rFlag = sharedSI.getInnerElements("r");
                if (rFlag == null) {
                    let tFlag = sharedSI.getInnerElements("t");
                    if (tFlag != null) {
                        let text = "";
                        tFlag.forEach((t) => {
                            text += t.value;
                        });

                        text = escapeCharacter(text);

                        //isContainMultiType(text) &&
                        if (/&#\d+;/.test(text)) {
                            text = this.htmlDecode(text);
                        }
                        if (familyFont == "Roman" && text.length > 0) {
                            let textArray = text.split("");
                            let preWordType: string = null, wordText = "", preWholef: string = null;
                            let wholef = "Times New Roman";
                            if (cellValue.ff != null) {
                                wholef = cellValue.ff;
                            }

                            let cellFormat = cellValue.ct;
                            if (cellFormat == null) {
                                cellFormat = new LuckySheetCellFormat();
                            }

                            if (cellFormat.s == null) {
                                cellFormat.s = [];
                            }

                            for (let i = 0; i < textArray.length; i++) {
                                let w = textArray[i];
                                let type: string = null, ff = wholef;

                                if (isChinese(w)) {
                                    type = "c";
                                    ff = "宋体";
                                }
                                else if (isJapanese(w)) {
                                    type = "j";
                                    ff = "Yu Gothic";
                                }
                                else if (isKoera(w)) {
                                    type = "k";
                                    ff = "Malgun Gothic";
                                }
                                else {
                                    type = "e";
                                }

                                if ((type != preWordType && preWordType != null) || i == textArray.length - 1) {
                                    let InlineString = new LuckyInlineString();

                                    InlineString.ff = preWholef;

                                    if (cellValue.fc != null) {
                                        InlineString.fc = cellValue.fc;
                                    }

                                    if (cellValue.fs != null) {
                                        InlineString.fs = cellValue.fs;
                                    }

                                    if (cellValue.cl != null) {
                                        InlineString.cl = cellValue.cl;
                                    }

                                    if (cellValue.un != null) {
                                        InlineString.un = cellValue.un;
                                    }

                                    if (cellValue.bl != null) {
                                        InlineString.bl = cellValue.bl;
                                    }

                                    if (cellValue.it != null) {
                                        InlineString.it = cellValue.it;
                                    }

                                    if (i == textArray.length - 1) {
                                        if (type == preWordType) {
                                            InlineString.ff = ff;
                                            InlineString.v = wordText + w;
                                        }
                                        else {
                                            InlineString.ff = preWholef;
                                            InlineString.v = wordText;
                                            cellFormat.s.push(InlineString);

                                            let InlineStringLast = new LuckyInlineString();
                                            InlineStringLast.ff = ff;
                                            InlineStringLast.v = w;
                                            if (cellValue.fc != null) {
                                                InlineStringLast.fc = cellValue.fc;
                                            }

                                            if (cellValue.fs != null) {
                                                InlineStringLast.fs = cellValue.fs;
                                            }

                                            if (cellValue.cl != null) {
                                                InlineStringLast.cl = cellValue.cl;
                                            }

                                            if (cellValue.un != null) {
                                                InlineStringLast.un = cellValue.un;
                                            }

                                            if (cellValue.bl != null) {
                                                InlineStringLast.bl = cellValue.bl;
                                            }

                                            if (cellValue.it != null) {
                                                InlineStringLast.it = cellValue.it;
                                            }
                                            cellFormat.s.push(InlineStringLast);

                                            break;
                                        }
                                    }
                                    else {
                                        InlineString.v = wordText;
                                    }


                                    cellFormat.s.push(InlineString);

                                    wordText = w;
                                }
                                else {
                                    wordText += w;
                                }


                                preWordType = type;
                                preWholef = ff;
                            }

                            cellFormat.t = "inlineStr";
                            // cellFormat.s = [InlineString];
                            cellValue.ct = cellFormat;
                            // console.log(cellValue);
                        }
                        else {


                            text = replaceSpecialWrap(text);

                            if (text.indexOf("\r\n") > -1 || text.indexOf("\n") > -1) {
                                let InlineString = new LuckyInlineString();
                                InlineString.v = text;
                                let cellFormat = cellValue.ct;
                                if (cellFormat == null) {
                                    cellFormat = new LuckySheetCellFormat();
                                }

                                if (cellValue.ff != null) {
                                    InlineString.ff = cellValue.ff;
                                }

                                if (cellValue.fc != null) {
                                    InlineString.fc = cellValue.fc;
                                }

                                if (cellValue.fs != null) {
                                    InlineString.fs = cellValue.fs;
                                }

                                if (cellValue.cl != null) {
                                    InlineString.cl = cellValue.cl;
                                }

                                if (cellValue.un != null) {
                                    InlineString.un = cellValue.un;
                                }

                                if (cellValue.bl != null) {
                                    InlineString.bl = cellValue.bl;
                                }

                                if (cellValue.it != null) {
                                    InlineString.it = cellValue.it;
                                }

                                cellFormat.t = "inlineStr";
                                cellFormat.s = [InlineString];
                                cellValue.ct = cellFormat;
                            }
                            else {
                                cellValue.v = text;
                                quotePrefix = "1";
                            }
                        }

                    }
                }
                else {
                    let styles: LuckyInlineString[] = [];
                    rFlag.forEach((r) => {
                        let tFlag = r.getInnerElements("t");
                        let rPr = r.getInnerElements("rPr");

                        let InlineString = new LuckyInlineString();

                        if (tFlag != null && tFlag.length > 0) {
                            let text = tFlag[0].value;
                            text = replaceSpecialWrap(text);
                            text = escapeCharacter(text);
                            InlineString.v = text;
                        }

                        if (rPr != null && rPr.length > 0) {
                            let frpr = rPr[0];
                            let sz = getlineStringAttr(frpr, "sz"), rFont = getlineStringAttr(frpr, "rFont"), family = getlineStringAttr(frpr, "family"), charset = getlineStringAttr(frpr, "charset"), scheme = getlineStringAttr(frpr, "scheme"), b = getlineStringAttr(frpr, "b"), i = getlineStringAttr(frpr, "i"), u = getlineStringAttr(frpr, "u"), strike = getlineStringAttr(frpr, "strike"), vertAlign = getlineStringAttr(frpr, "vertAlign"), color;


                            let cEle = frpr.getInnerElements("color");
                            if (cEle != null && cEle.length > 0) {
                                color = getColor(cEle[0], this.styles, "t");
                            }


                            let ff;
                            // if(family!=null){
                            //     ff = fontFamilys[family];
                            // }
                            if (rFont != null) {
                                ff = rFont;
                            }
                            if (ff != null) {
                                InlineString.ff = ff;
                            }
                            else if (cellValue.ff != null) {
                                InlineString.ff = cellValue.ff;
                            }

                            if (color != null) {
                                InlineString.fc = color;
                            }
                            // else if(cellValue.fc!=null){
                            //     InlineString.fc = cellValue.fc;
                            // }

                            if (sz != null) {
                                InlineString.fs = parseInt(sz);
                            }
                            else if (cellValue.fs != null) {
                                InlineString.fs = cellValue.fs;
                            }

                            if (strike != null) {
                                InlineString.cl = parseInt(strike);
                            }
                            else if (cellValue.cl != null) {
                                InlineString.cl = cellValue.cl;
                            }

                            if (u != null) {
                                InlineString.un = parseInt(u);
                            }
                            else if (cellValue.un != null) {
                                InlineString.un = cellValue.un;
                            }

                            if (b != null) {
                                InlineString.bl = parseInt(b);
                            }
                            else if (cellValue.bl != null) {
                                InlineString.bl = cellValue.bl;
                            }

                            if (i != null) {
                                InlineString.it = parseInt(i);
                            }
                            else if (cellValue.it != null) {
                                InlineString.it = cellValue.it;
                            }

                            if (vertAlign != null) {
                                InlineString.va = parseInt(vertAlign);
                            }


                            // ff:string | undefined //font family
                            // fc:string | undefined//font color
                            // fs:number | undefined//font size
                            // cl:number | undefined//strike
                            // un:number | undefined//underline
                            // bl:number | undefined//blod
                            // it:number | undefined//italic
                            // v:string | undefined
                        }
                        else {
                            if (InlineString.ff == null && cellValue.ff != null) {
                                InlineString.ff = cellValue.ff;
                            }

                            if (InlineString.fc == null && cellValue.fc != null) {
                                InlineString.fc = cellValue.fc;
                            }

                            if (InlineString.fs == null && cellValue.fs != null) {
                                InlineString.fs = cellValue.fs;
                            }

                            if (InlineString.cl == null && cellValue.cl != null) {
                                InlineString.cl = cellValue.cl;
                            }

                            if (InlineString.un == null && cellValue.un != null) {
                                InlineString.un = cellValue.un;
                            }

                            if (InlineString.bl == null && cellValue.bl != null) {
                                InlineString.bl = cellValue.bl;
                            }

                            if (InlineString.it == null && cellValue.it != null) {
                                InlineString.it = cellValue.it;
                            }
                        }


                        styles.push(InlineString);
                    });

                    let cellFormat = cellValue.ct;
                    if (cellFormat == null) {
                        cellFormat = new LuckySheetCellFormat();
                    }
                    cellFormat.t = "inlineStr";
                    cellFormat.s = styles;
                    cellValue.ct = cellFormat;
                }
            }
            if (t == ST_CellType["SharedString"]) {
                let siIndex = parseInt(v[0].value);
                let sharedSI = sharedStrings[siIndex];
                generateInlineString(sharedSI);
                
            }
            else if (t == ST_CellType["InlineString"] && v != null) {
                cellValue.v = this.htmlDecode(value);

                let cellFormat = cellValue.ct;
                if (cellFormat == null) {
                    cellFormat = new LuckySheetCellFormat();
                }
                cellFormat.t = "s";
                cellValue.ct = cellFormat;
                
                let is = this.cell.getInnerElements("is");

                if (is.length) {
                    generateInlineString(is[0]);
                }
            }
            else if (t == ST_CellType["String"] && value.includes('=DISPIMG')) {
                let cellFormat = cellValue.ct;
                if (cellFormat == null) {
                    cellFormat = new LuckySheetCellFormat();
                }
                cellFormat.t = "str";
                cellFormat.ci = this.getCellImage(cellValue, value);
                cellValue.ct = cellFormat;
            }
            else {
                value = escapeCharacter(value);
                cellValue.v = value;
            }
        }

        if (quotePrefix != null) {
            cellValue.qp = parseInt(quotePrefix);
        }

        if (t !== null && !cellValue.ct?.t) {
            let cellFormat = new LuckySheetCellFormat();
            cellFormat.t = t || 'd';
            cellValue.ct = cellFormat;
        }

        return cellValue;

    }

    private htmlDecode(str: string): string {
        return str.replace(/&#(x)?([^&]{1,5});/g, function ($, $1, $2) {
            return String.fromCharCode(parseInt($2, $1 ? 16 : 10));
        });
    };

    private getCellImage(cellValue: LuckySheetCelldataValue, value: string): any {
        const id = this.extractImageId(value);
        let ci: any = {};
        this.cellImages.forEach((element) => {
            const pic = element.getInnerElements('xdr:pic')[0];

            const picpr = pic.getInnerElements('xdr:nvPicPr')[0]
            const nvpr = picpr.getInnerElements('xdr:cNvPr')[0];

            const blipfill = pic.getInnerElements('xdr:blipFill')[0]
            const blip = blipfill.getInnerElements('a:blip')[0];

            const picId = nvpr.get('name');
            const picRid = blip.get('r:embed') as string;

            if (id == picId) {
                const obj = this.getBase64ByRid(picRid, cellImagesRels)
                ci = obj;

                let x_n =0,y_n = 0;
                let cx_n = 0, cy_n = 0;
                const sp = pic.getInnerElements('xdr:spPr')[0];
                const xfrm = sp.getInnerElements('a:xfrm')[0];
                const off = xfrm.getInnerElements('a:off')[0];
                const ext = xfrm.getInnerElements('a:ext')[0];
                cx_n = getPxByEMUs(parseInt(ext.get('cx') as string)),cy_n = getPxByEMUs(parseInt(ext.get('cy') as string));
                x_n = getPxByEMUs(parseInt(off.get('x') as string)),y_n = getPxByEMUs(parseInt(off.get('y') as string));

                const rateX = cx_n / this.cellSize.width;
                const rateY = cy_n / this.cellSize.height;

                if(rateX > 1 && rateX > rateY) {
                    cy_n = cy_n / rateX;
                    cx_n = this.cellSize.width;
                }

                if (rateY > 1 && rateY > rateX) {
                    cx_n = cx_n / rateY;
                    cy_n = this.cellSize.height;
                }
                let imageDefault:IluckyImageDefault = {
                    height: cy_n,
                    left: x_n,
                    top: y_n,
                    width: cx_n
                }
                ci.default = imageDefault;
                ci.descr = nvpr.get('descr');
            }
        })
        return ci
    }
    private extractImageId(formula: string) {
        const regex = /ID_[A-Za-z0-9]{32}/;
        const match = formula.match(regex);
        return match ? match[0] : null;
    }
    private getBase64ByRid(rid:string, drawingRelsFile:string){
        let Relationships = this.readXml.getElementsByTagName("Relationships/Relationship", drawingRelsFile);

        if(Relationships!=null && Relationships.length>0){
            for(let i=0;i<Relationships.length;i++){
                let Relationship = Relationships[i];
                let attrList = Relationship.attributeList;
                let Id = getXmlAttibute(attrList, "Id", null);
                let src = getXmlAttibute(attrList, "Target", null);
                if(Id == rid){
                    src = src.replace(/\.\.\//g, "");
                    src = "xl/" + src;
                    let imgage = this.imageList.getImageByName(src);
                    return imgage;
                }
            }
        }

        return {};
    }
}

