import { workBookFile } from "../common/constant";
import { escapeCharacter, generateRandomId } from "../common/method";
import { Element, ReadXml } from "./ReadXml";

export interface IDefinedNameParam {
    id: string;
    name: string;
    formulaOrRefString: string;
    comment?: string;
    localSheetId?: string;
    hidden?: boolean;
}
export interface IDefinedNames {
    [key: string]: LuckyDefineName
}
export class LuckyDefineNames {
    defineNames: IDefinedNames;
    constructor(readXml: ReadXml) {
        let definedNames = readXml.getElementsByTagName("definedNames/definedName", workBookFile);
        const obj: IDefinedNames = {};
        const validNames = (name: string) => {
            try {
                return typeof name === 'string' && /^[a-zA-Z_][\w.]*$/.test(name);
            } catch {
                return false;
            }
        };
        definedNames.forEach(d => {
            const definedName = new LuckyDefineName(d)
            if (validNames(definedName.name)) {
                obj[definedName.id] = definedName
            }
        })
        this.defineNames = obj;
    }
}
export class LuckyDefineName implements IDefinedNameParam {
    id: string;
    name: string;
    formulaOrRefString: string;
    comment?: string;
    localSheetId?: string;
    hidden?: boolean;
    constructor(ele: Element) {
        this.id = generateRandomId(6);
        this.name = ele.get('name') as string;
        this.formulaOrRefString = escapeCharacter(ele.value);
        this.comment = ele.get('comment') as string;
        this.localSheetId = ele.get('localSheetId') as string;
        this.hidden = ele.get('hidden') === '1';
    }

}