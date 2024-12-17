import { Anchor, Workbook, Worksheet } from "@zwight/exceljs";
import { escapeCharacter, getEmusByPx, getRangetxt, isEmpty, jsonParse, removeEmptyAttr, str2num } from "../common/method";
import { CFRuleType, CFSubRuleType } from "../ToLuckySheet/LuckyCondition";
import { hex2argb } from "./util";
import { cellStyle } from "./CellStyle";

export class Resource {
    sheetId: string;
    workbook: Workbook;
    worksheet: Worksheet;
    resources: any;
    constructor(
        sheetId: string,
        workbook: Workbook,
        worksheet: Worksheet,
        resources: any
    ) {
        this.sheetId = sheetId;
        this.workbook = workbook;
        this.worksheet = worksheet;
        this.resources = resources;

        this.setImages();
        this.setConditional();
        this.setDataValidation();
        this.setFilter();
    }

    private handleRang(range: any) {
        const { startRow, startColumn, endRow, endColumn } = range;
        return getRangetxt({
            row: [startRow, endRow],
            column: [startColumn, endColumn],
            sheetIndex: 0
        }, '')
    }
    // private setRangeProtection() {
    //     const rangeProtection = this.getSheetResource('SHEET_RANGE_PROTECTION_PLUGIN');
    // }
    private setFilter() {
        const filters = this.getSheetResource('SHEET_FILTER_PLUGIN');
        if (!filters) return;
        this.worksheet.autoFilter = this.handleRang(filters.ref);
    }
    private setConditional() {
        const conditionals = this.getSheetResource('SHEET_CONDITIONAL_FORMATTING_PLUGIN');
        const ruleList: any[] = [];
        if (!conditionals) return;
        conditionals.forEach((conditional: any) => {
            const { ranges, rule, stopIfTrue } = conditional;
            ranges.forEach((range: any) => {
                const ref = this.handleRang(range);
                const index = ruleList.findIndex(d => d.ref === ref);
                const ruleValue = this.handleRule(conditional);
                if (index > -1) {
                    ruleList[index].rules.push(ruleValue)
                    return;
                }
                ruleList.push({
                    ref,
                    rules: [ruleValue]
                })
            });
        });
        // console.log(this.worksheet.name, ruleList)
        ruleList.forEach(d => {
            this.worksheet.addConditionalFormatting(d);
        });
    }

    private setDataValidation() {
        const datavalidations = this.getSheetResource('SHEET_DATA_VALIDATION_PLUGIN');
        const excelToDate = (v: number, date1904 = false) => {
            // eslint-disable-next-line no-mixed-operators
            const millisecondSinceEpoch = Math.round((v - 25569 + (date1904 ? 1462 : 0)) * 24 * 3600 * 1000);
            return new Date(millisecondSinceEpoch);
        }
        datavalidations?.forEach((validate: any) => {
            const {
                ranges = [],
                type,
                allowBlank,
                operator,
                formula1,
                formula2,
                showErrorMessage,
                showInputMessage,
                prompt,
                promptTitle,
                error,
                errorTitle,
                errorStyle
            } = validate || {};


            const styleMap = ['information', 'stop', 'warning']
            
            let formulae = [formula1];
            if (!isEmpty(formula2)) formulae.push(formula2);

            if (type === 'list') {
                formulae = [`"${formula1}"`];
                if (!isEmpty(formula2)) formulae.push(`"${formula2}"`);
            }
            if (type === 'date') {
                formulae = [excelToDate(formula1)];
                if (!isEmpty(formula2)) formulae.push(excelToDate(formula2));
            }

            const list = ranges.map((range: any) => this.handleRang(range));

            
            const valid = {
                type,
                allowBlank,
                operator,
                formulae,
                showErrorMessage,
                showInputMessage,
                prompt,
                promptTitle,
                error,
                errorTitle,
                errorStyle: styleMap[errorStyle]
            }
            // this.worksheet.dataValidations.add(list.join(' '), removeEmptyAttr(valid))
            // console.log(this.sheetId, this.worksheet.name, list, valid);
            list.forEach((address: string) => {
                this.worksheet.dataValidations.add(address, removeEmptyAttr(valid))
            });
        });
    }

    private setImages() {
        const images = this.getSheetResource('SHEET_DRAWING_PLUGIN');
        const sheetIamges = images?.data;
        if (!sheetIamges) return;
        for (const key in sheetIamges) {
            const element = sheetIamges[key];

            const images = this.workbook.getImages();
            let imageId = images.findIndex(d => d.extension === 'png' && d.base64 === element.source);
            if (imageId === -1) {
                imageId = this.workbook.addImage({
                    base64: element.source,
                    extension: 'png'
                })
            }
            
            const handlePosition = (position: any) => {
                return {
                    nativeCol: position.column,
                    nativeColOff: getEmusByPx(position.columnOffset),
                    nativeRow: position.row,
                    nativeRowOff: getEmusByPx(position.rowOffset),
                }
            }
            // console.log(handlePosition(element.sheetTransform.from), handlePosition(element.sheetTransform.to))

            this.worksheet.addImage(imageId, {
                tl: handlePosition(element.sheetTransform.from) as Anchor,
                br: handlePosition(element.sheetTransform.to) as Anchor
            })
        }
    }

    private getSheetResource(name: string) {
        const resources = jsonParse(this.resources.find((d: any) => d.name === name)?.data);
        const sheetResources = resources[this.sheetId];
        return sheetResources;
    }

    private handleRule(conditional: any) {
        const { rule, stopIfTrue, order } = conditional;
        const ruleValue: any = {};
        if (stopIfTrue) ruleValue.stopIfTrue = 1;
        ruleValue.priority = order;
        if (rule.style) {
            // const formatCode = rule.style.n?.pattern ? { formatCode: rule.style.n?.pattern } : undefined;
            ruleValue.style = cellStyle(rule.style, rule.style.n?.pattern, true);
        }
        if (rule.operator) ruleValue.operator = rule.operator;
        switch (rule.type) {
            case CFRuleType.colorScale:
                ruleValue.type = CFRuleType.colorScale;
                ruleValue.cfvo = rule.config?.map((d: any) => {
                    return {
                        type: d.value.type,
                        value: d.value.value
                    }
                });
                ruleValue.color = rule.config?.map((d: any) => {
                    return {
                        argb: hex2argb(d.color)
                    }
                });
                break;
            case CFRuleType.dataBar:
                ruleValue.type = CFRuleType.dataBar;
                ruleValue.showValue = rule.isShowValue;
                ruleValue.gradient = rule.config.isGradient;
                ruleValue.cfvo = [{
                    type: rule.config.min.type,
                    value: rule.config.min.value
                }, {
                    type: rule.config.max.type,
                    value: rule.config.max.value
                }];
                ruleValue.negativeFillColor = { argb: hex2argb(rule.config.nativeColor) };
                ruleValue.color = { argb: hex2argb(rule.config.positiveColor) }

                // ruleValue.border = false;
                ruleValue.axisPosition = 'auto';
                ruleValue.direction = 'leftToRight';
                ruleValue.minLength = 0;
                ruleValue.maxLength = 100;
                ruleValue.negativeBarColorSameAsPositive = true;
                ruleValue.negativeBarBorderColorSameAsPositive = true;

                break;
            case CFRuleType.iconSet:
                ruleValue.type = CFRuleType.iconSet;
                ruleValue.reverse = false;
                ruleValue.showValue = rule.isShowValue;
                ruleValue.icons = rule.config?.map((d: any) => {
                    const iconId = (str2num(d.iconType.charAt(0)) as number) - (str2num(d?.iconId) as number) -1
                    return {
                        iconId,
                        iconSet: d.iconType
                    }
                }).reverse();
                ruleValue.custom = true;
                // ruleValue.iconSet = rule.config[0]?.iconType;
                ruleValue.cfvo = rule.config?.map((d: any) => {
                    return {
                        type: d.value.type,
                        value: d.value.value
                    }
                }).reverse();
                break;
            case CFRuleType.highlightCell:
                switch (rule.subType) {
                    case CFSubRuleType.average:
                        ruleValue.type = 'aboveAverage';
                        ruleValue.aboveAverage = false;
                        break;
                    case CFSubRuleType.duplicateValues:
                        ruleValue.type = 'duplicateValues';
                        break;
                    case CFSubRuleType.formula:
                        ruleValue.type = 'expression';
                        ruleValue.formulae = [escapeCharacter(rule.value)];
                        break;
                    case CFSubRuleType.number:
                        ruleValue.type = 'cellIs';
                        ruleValue.formulae = [rule.value];
                        break;
                    case CFSubRuleType.rank:
                        ruleValue.type = 'top10';
                        ruleValue.rank = rule.value;
                        ruleValue.percent = rule.isPercent;
                        ruleValue.bottom = rule.isBottom;
                        break;
                    case CFSubRuleType.text:
                        ruleValue.type = 'containsText';
                        ruleValue.text = rule.value
                        break;
                    case CFSubRuleType.timePeriod:
                        ruleValue.type = 'timePeriod';
                        ruleValue.timePeriod = rule.operator
                        break;
                    // case CFSubRuleType.uniqueValues:
                    //     ruleValue.type = 'uniqueValues';
                    //     break;
                }
                break;
        }
        return ruleValue;
    }
}