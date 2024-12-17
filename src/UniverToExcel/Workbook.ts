import exceljs from "@zwight/exceljs";
import { jsonParse } from "../common/method";
import { ExcelWorkSheet } from "./WorkSheet";
// SHEET_HYPER_LINK_PLUGIN
// SHEET_DRAWING_PLUGIN
// SHEET_DEFINED_NAME_PLUGIN
// SHEET_CONDITIONAL_FORMATTING_PLUGIN
// SHEET_DATA_VALIDATION_PLUGIN
// SHEET_FILTER_PLUGIN
const Workbook = exceljs.Workbook;
export class WorkBook extends Workbook {
    constructor(snapshot: any) {
        super()
        this.init(snapshot);
    }

    private init(snapshot: any) {
        // this.properties.date1904 = true;
        this.calcProperties.fullCalcOnLoad = true;

        this.setDefineNames(snapshot.resources);
        ExcelWorkSheet(this, snapshot)
    }

    private setDefineNames(resources: any[]) {
        const definedNames = jsonParse(resources.find(d => d.name === 'SHEET_DEFINED_NAME_PLUGIN')?.data);
        for (const key in definedNames) {
            const element = definedNames[key]
            this.definedNames.add(element.formulaOrRefString, element.name);
        }
    }
    // private setSheetProtection(snapshot: any) {
    //     const { resources, id } = snapshot
    //     const sheetProtections = jsonParse(resources.find((d: any) => d.name === 'SHEET_WORKSHEET_PROTECTION_PLUGIN')?.data);
    //     const protection = sheetProtections?.[id];
    //     if (!protection) return;

    // }
}