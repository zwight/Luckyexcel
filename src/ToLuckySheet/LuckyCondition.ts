/**:
* @description: 返回的condition格式为univer使用的格式，但是规则配置中的style是luckyexcel的格式，
* 后面考虑新增配置用来输出univer所需的数据结构
* @author: Created by zwight on 2024-09-20 16:13:34
*/
import { escapeCharacter, generateRandomIndex, getXmlAttibute, str2num } from "../common/method";
import { IattributeList } from "../ICommon";
import { LuckySheetCellFormat } from "./LuckyBase";
import { Element, getColor, IStyleCollections, ReadXml } from "./ReadXml";
import { getBackgroundByFill, getFontStyle, handleBorder, handleRanges, IRange } from "./style";


export interface ConditionFormationRule{
    type: string,
    subType?: string,
    style?: {},
    value?: string,
    isShowValue?: boolean
}
export interface LuckyConditionFormat<C = any> {
    ranges: IRange[],
    cfId: string,
    stopIfTrue: boolean,
    order: number,
    rule: C
}

export enum CFTextOperator {
    beginsWith = 'beginsWith',
    endsWith = 'endsWith',
    containsText = 'containsText',
    notContainsText = 'notContainsText',
    equal = 'equal',
    notEqual = 'notEqual',
    containsBlanks = 'containsBlanks',
    notContainsBlanks = 'notContainsBlanks',
    containsErrors = 'containsErrors',
    notContainsErrors = 'notContainsErrors',
}
export enum CFTimePeriodOperator {
    today = 'today',
    yesterday = 'yesterday',
    tomorrow = 'tomorrow',
    last7Days = 'last7Days',
    thisMonth = 'thisMonth',
    lastMonth = 'lastMonth',
    nextMonth = 'nextMonth',
    thisWeek = 'thisWeek',
    lastWeek = 'lastWeek',
    nextWeek = 'nextWeek',
}
export enum CFNumberOperator {
    greaterThan = 'greaterThan',
    greaterThanOrEqual = 'greaterThanOrEqual',
    lessThan = 'lessThan',
    lessThanOrEqual = 'lessThanOrEqual',
    notBetween = 'notBetween',
    between = 'between',
    equal = 'equal',
    notEqual = 'notEqual',
}

export enum CFRuleType {
    highlightCell = 'highlightCell',
    dataBar = 'dataBar',
    colorScale = 'colorScale',
    iconSet = 'iconSet',
}
export enum CFSubRuleType {
    uniqueValues = 'uniqueValues',
    duplicateValues = 'duplicateValues',
    rank = 'rank',
    text = 'text',
    timePeriod = 'timePeriod',
    number = 'number',
    average = 'average',
    formula = 'formula',
}

interface ConditionElement extends Element { 
    parentAttribute: IattributeList
    extLst: Element,
    isExtLst: boolean
}
export class LuckyCondition implements LuckyConditionFormat {
    ranges: IRange[];
    cfId: string;
    stopIfTrue: boolean = false;
    order: number;
    rule: any;
    constructor(ele: ConditionElement, readXml: ReadXml, styles: IStyleCollections) {
        const { attributeList, parentAttribute } = ele;
        if(parentAttribute?.sqref) this.ranges = handleRanges(parentAttribute.sqref);
        this.order = Number(getXmlAttibute(attributeList, 'priority', '1'));
        this.cfId = generateRandomIndex('condition');
        this.stopIfTrue = getXmlAttibute(attributeList, 'stopIfTrue', '0') === '1' ? true : false;
        this.handleRules(ele, readXml, styles)
    }

    private handleRules = (ele: ConditionElement, readXml: ReadXml, styles: IStyleCollections) => {
        const { attributeList, value, extLst, isExtLst } = ele
        const type = getXmlAttibute(attributeList, 'type', 'expression')
        const rule: any = {
            type: CFRuleType.highlightCell
        }
        const operator = getXmlAttibute(attributeList, 'operator', '')
        const rank = getXmlAttibute(attributeList, 'rank', '');
        
        const formula = readXml.getElementsByTagName("formula", value, false);
        const aboveAverage = getXmlAttibute(attributeList, 'aboveAverage', '');

        if (operator) rule.operator = operator;
        if (formula[0]?.value || formula[0]?.value == '0') rule.value = str2num(formula[0]?.value);
        switch (type) {
            case 'expression':
                rule.subType = CFSubRuleType.formula;
                break;
            case 'cellIs':
                rule.subType = CFSubRuleType.number;
                break;
            case 'top10':
                rule.subType = CFSubRuleType.rank;
                const percent = getXmlAttibute(attributeList, 'percent', '0');
                const bottom = getXmlAttibute(attributeList, 'bottom', '0');
                if (rank) rule.value = str2num(rank);
                rule.isBottom = bottom === '1' ? true : false;
                rule.isPercent = percent === '1' ? true : false;
                break;
            case 'aboveAverage':
                rule.subType = CFSubRuleType.average;
                rule.operator = rule.operator || CFNumberOperator.lessThan
                break;
            case 'timePeriod':
                rule.subType = CFSubRuleType.timePeriod;
                rule.operator = getXmlAttibute(attributeList, 'timePeriod', undefined);
                break;
            case 'duplicateValues':
                rule.subType = CFSubRuleType.duplicateValues;
                break;
            case 'containsText': 
                rule.subType = CFSubRuleType.text;
                rule.operator = 'containsText'
                rule.value = getXmlAttibute(attributeList, 'text', '');
                break;
            case 'colorScale':
                const cfvo = readXml.getElementsByTagName("colorScale/cfvo", value, false)
                const color = readXml.getElementsByTagName("colorScale/color", value, false)

                rule.type = CFRuleType.colorScale;
                rule.config = cfvo.map((d, index) => {
                    const type = getXmlAttibute(d.attributeList, 'type', '');
                    const value = getXmlAttibute(d.attributeList, 'val', undefined);
                    const colorVal = color[index] ? getColor(color[index], styles) : undefined
                    return {
                        index: 0,
                        color: colorVal,
                        value: {
                            type,
                            value: str2num(value)
                        }
                    }
                })
                break;
            case 'dataBar':
                rule.type = CFRuleType.dataBar;

                const dataBar = readXml.getElementsByTagName("dataBar", value, false)?.[0];
                const barCfvo = readXml.getElementsByTagName("cfvo", dataBar.value, false);
                const barColor = readXml.getElementsByTagName("color", dataBar.value, false);
                const isShowValue = getXmlAttibute(dataBar.attributeList, 'showValue', '1');
                rule.isShowValue = isShowValue === '1';
                
                let positiveColor = barColor[0] ? getColor(barColor[0], styles) : undefined;
                let isGradient = true;
                let nativeColor = '';
                if (extLst) {
                    const x14DataBar = readXml.getElementsByTagName("x14:dataBar", extLst.value, false)[0];
                    const gradient = getXmlAttibute(x14DataBar.attributeList, 'gradient', null);
                    const negativeFillColor = readXml.getElementsByTagName("x14:dataBar/x14:negativeFillColor", extLst.value, false)?.[0];
                    nativeColor = negativeFillColor ? getColor(negativeFillColor, styles) : undefined;
                    isGradient = gradient !== '0';
                }
                rule.config = {
                    min: {
                        type: getXmlAttibute(barCfvo[0]?.attributeList, 'type', 'min'),
                        value: str2num(getXmlAttibute(barCfvo[0]?.attributeList, 'val', undefined))
                    },
                    max: {
                        type: getXmlAttibute(barCfvo[1]?.attributeList, 'type', 'max'),
                        value: str2num(getXmlAttibute(barCfvo[1]?.attributeList, 'val', undefined))
                    },
                    isGradient,
                    positiveColor,
                    nativeColor,
                }

                break;
            case 'iconSet':
                rule.type = CFRuleType.iconSet;
                if (!isExtLst) {
                    const iconSet = readXml.getElementsByTagName("iconSet", value, false);
                    const iconCfvo = readXml.getElementsByTagName("iconSet/cfvo", value, false);

                    rule.isShowValue = iconSet[0]?.attributeList?.showValue === '0' ? false : true;
                    rule.config = iconCfvo.map((d, index) => {
                        return {
                            operator: rule.operator || CFNumberOperator.greaterThanOrEqual,
                            value: {
                                type: d.attributeList.type,
                                value: str2num(d.attributeList.val)
                            },
                            iconType: iconSet[0].attributeList.iconSet,
                            iconId: index
                        }
                    })
                } else {
                    const iconSet = readXml.getElementsByTagName("x14:iconSet", value, false)[0];
                    const iconCfvo = readXml.getElementsByTagName("x14:iconSet/x14:cfvo", value, false);
                    const cfIcon = readXml.getElementsByTagName("x14:iconSet/x14:cfIcon", value, false);
                    const isCustom = getXmlAttibute(iconSet?.attributeList, 'custom', '0') === '1'

                    rule.isShowValue = iconSet?.attributeList?.showValue === '0' ? false : true;
                    rule.config = iconCfvo.map((d, index) => {
                        const value = readXml.getElementsByTagName("xm:f", d.value, false)[0];

                        let iconInfo = cfIcon[index]?.attributeList
                        const iconType = isCustom ? iconInfo?.iconSet : iconSet?.attributeList.iconSet;
                        const iconId = (str2num(iconType.charAt(0)) as number) - (str2num(iconInfo?.iconId) as number) -1;
                        return {
                            operator: CFNumberOperator.greaterThanOrEqual,
                            value: {
                                type: d.attributeList.type,
                                value: str2num(value?.value)
                            },
                            iconType,
                            iconId,
                            // sort: str2num(iconInfo?.iconId)
                        }
                    }).reverse();
                }
                
                break;
        }
        const dxfId = getXmlAttibute(attributeList, 'dxfId', null);
        if (dxfId) {
            let dxfs = styles["dxfs"] as Element[];
            const dxf = dxfId !== null ? dxfs[Number(dxfId)] : undefined;
            // let numFmt,font,fill,border;

            const font = dxf.getInnerElements('font')?.[0];
            const numFmt = dxf.getInnerElements('numFmt')?.[0];
            const fill = dxf.getInnerElements('fill')?.[0];
            const border = dxf.getInnerElements('border')?.[0];

            let numfmts = styles["numfmts"] as IattributeList;

            let style: any = {}
            if (border) {
                const { borderCellValue } = handleBorder(border, styles);
                style.border = borderCellValue
            }
            if (fill) {
                const bg = getBackgroundByFill(fill, styles);
                style.bg = bg;
            }

            if (numFmt) {
                let numf = numfmts[parseInt(numFmt?.attributeList?.numFmtId)];
                let cellFormat = new LuckySheetCellFormat();
                cellFormat.fa = numf ? escapeCharacter(numf) : numFmt.attributeList.formatCode;
                style.ct = cellFormat;
            }

            if (font) {
                const { cellValue } = getFontStyle(font, style)
                style = { ...style, ...cellValue };
            }

            rule.style = style;
        }
        this.rule = rule;
    }

}