/**:
* @description: 这里处理的dataverification格式是univer用的格式
* @author: Created by zwight on 2024-09-23 17:50:46
*/
import { generateRandomIndex, getMultiFormulaValue, getPeelOffX14, getXmlAttibute } from "../common/method";
import { Element } from "./ReadXml";
import { handleRanges, IRange } from "./style";
export declare enum DataValidationType {
    /**
     * custom formula
     */
    CUSTOM = "custom",
    LIST = "list",
    LIST_MULTIPLE = "listMultiple",
    NONE = "none",
    TEXT_LENGTH = "textLength",
    DATE = "date",
    TIME = "time",
    /**
     * integer number
     */
    WHOLE = "whole",
    /**
     * decimal number
     */
    DECIMAL = "decimal",
    CHECKBOX = "checkbox"
}

export declare enum DataValidationOperator {
    BETWEEN = "between",
    EQUAL = "equal",
    GREATER_THAN = "greaterThan",
    GREATER_THAN_OR_EQUAL = "greaterThanOrEqual",
    LESS_THAN = "lessThan",
    LESS_THAN_OR_EQUAL = "lessThanOrEqual",
    NOT_BETWEEN = "notBetween",
    NOT_EQUAL = "notEqual"
}
export declare enum DataValidationErrorStyle {
    INFO = 0,
    STOP = 1,
    WARNING = 2
}
export declare enum DataValidationImeMode {
    DISABLED = 0,
    FULL_ALPHA = 1,
    FULL_HANGUL = 2,
    FULL_KATAKANA = 3,
    HALF_ALPHA = 4,
    HALF_HANGUL = 5,
    HALF_KATAKANA = 6,
    HIRAGANA = 7,
    NO_CONTROL = 8,
    OFF = 9,
    ON = 10
}
export declare enum DataValidationRenderMode {
    TEXT = 0,
    ARROW = 1,
    CUSTOM = 2
}
export interface LuckyVerificationFormat {
    type: DataValidationType;
    allowBlank?: boolean;
    formula1?: string;
    formula2?: string;
    operator?: DataValidationOperator;
    imeMode?: DataValidationImeMode;
    /**
     * for dropdown
     */
    showDropDown?: boolean;
    /**
     * validator error
     */
    showErrorMessage?: boolean;
    error?: string;
    errorStyle?: DataValidationErrorStyle;
    errorTitle?: string;
    /**
     * input prompt, not for using now.
     */
    showInputMessage?: boolean;
    prompt?: string;
    promptTitle?: string;
    /**
     * cell render mode of data validation
     */
    renderMode?: DataValidationRenderMode;
    /**
     * custom biz info
     */
    bizInfo?: {
        showTime?: boolean;
    };
    uid: string;
    ranges: IRange[];
}
export class LuckyVerification implements LuckyVerificationFormat {
    type: DataValidationType;
    allowBlank?: boolean;
    formula1?: string;
    formula2?: string;
    operator?: DataValidationOperator;
    imeMode?: DataValidationImeMode;
    showDropDown?: boolean;
    showErrorMessage?: boolean;
    error?: string;
    errorStyle?: DataValidationErrorStyle;
    errorTitle?: string;
    showInputMessage?: boolean;
    prompt?: string;
    promptTitle?: string;
    renderMode?: DataValidationRenderMode;
    bizInfo?: { showTime?: boolean; };
    uid: string;
    ranges: IRange[];
    constructor(ele: Element, extLst: Element[]) {
        let attrList = ele.attributeList;
        let formulaValue = ele.value;

        let type = getXmlAttibute(attrList, "type", undefined);
        if (!type) {
            return;
        }

        this.uid = generateRandomIndex('verification')

        let valueArr: string[] = [], sqref = '';
        const operator = getXmlAttibute(attrList, "operator", undefined);
        const allowBlank = getXmlAttibute(attrList, "allowBlank", undefined) !== "1" ? false : true;
        const showInputMessage = getXmlAttibute(attrList, "showInputMessage", undefined) !== "1" ? false : true;
        const showErrorMessage = getXmlAttibute(attrList, "showErrorMessage", undefined) !== "1" ? false : true;
        const prompt = getXmlAttibute(attrList, "prompt", undefined);
        const promptTitle = getXmlAttibute(attrList, "promptTitle", undefined);
        const error = getXmlAttibute(attrList, "error", undefined);
        const errorTitle = getXmlAttibute(attrList, "errorTitle", undefined);
        const errorStyle = getXmlAttibute(attrList, "errorStyle", 'stop');


        const formulaReg = new RegExp(/<x14:formula1>|<xm:sqref>/g)
        if (formulaReg.test(formulaValue) && extLst?.length >= 0) {
            const peelOffData = getPeelOffX14(formulaValue);
            sqref = peelOffData?.sqref;
            valueArr = getMultiFormulaValue(peelOffData?.formula);
        } else {
            valueArr = getMultiFormulaValue(formulaValue);
            sqref = getXmlAttibute(attrList, "sqref", null);
        }

        let _value1: string | number = valueArr?.length >= 1 ? valueArr[0] : undefined;
        let _value2: string | number = valueArr?.length === 2 ? valueArr[1] : undefined;

        if (sqref) this.ranges = handleRanges(sqref);
        this.type = type as DataValidationType;
        this.allowBlank = allowBlank;
        this.operator = operator as DataValidationOperator;
        this.formula1 = _value1;
        this.formula2 = _value2;
        this.showErrorMessage = showErrorMessage;
        this.showInputMessage = showInputMessage;
        this.prompt = prompt;
        this.promptTitle = promptTitle;
        this.error = error;
        this.errorTitle = errorTitle;
        switch (errorStyle) {
            case 'information':
                this.errorStyle = 0
                break;
            case 'warning':
                this.errorStyle = 2
                break;
            case 'stop':
                this.errorStyle = 1
                break;
        }
        // imeMode?: DataValidationImeMode;
        // renderMode?: DataValidationRenderMode;
        // bizInfo?: { showTime?: boolean; };

        this.ranges = handleRanges(sqref);
    }
}