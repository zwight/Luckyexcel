import {
    ICellData,
    IObjectMatrixPrimitiveType,
    IResources,
    IStyleData,
    IWorkbookData,
    IWorksheetData,
    LocaleType,
    Nullable,
} from '@univerjs/core';
import { UniverSheetBase } from './UniverSheetBase';
import { generateRandomId } from '../common/method';

export class UniverCsvWorkBook implements IWorkbookData {
    id: string;
    rev?: number | undefined;
    name: string;
    appVersion!: string;
    locale!: LocaleType;
    styles!: Record<string, Nullable<IStyleData>>;
    sheetOrder!: string[];
    sheets!: { [sheetId: string]: Partial<IWorksheetData> };
    resources?: IResources | undefined;
    constructor(data: string[][]) {
        console.log(data);
        const cellData: IObjectMatrixPrimitiveType<ICellData> = {};

        let rowCount = 0,
            colCount = 0;

        data.forEach((row, rowIndex) => {
            if (rowIndex + 1 > rowCount) rowCount = rowIndex + 1;
            row.forEach((col, colIndex) => {
                if (colIndex + 1 > colCount) colCount = colIndex + 1;

                if (!cellData[rowIndex]) cellData[rowIndex] = {};
                cellData[rowIndex][colIndex] = { v: col || '' };
            });
        });
        const sheetId = `sheet1`;
        const sheet = new UniverSheetBase({ id: sheetId, name: sheetId, cellData, rowCount, colCount });
        this.sheets = { [sheetId]: sheet };
        this.sheetOrder = [sheetId];
        this.id = generateRandomId(6);
        this.name = this.id;
    }
    get mode(): IWorkbookData {
        return {
            id: this.id,
            rev: this.rev,
            name: this.name,
            appVersion: this.appVersion,
            locale: this.locale,
            styles: this.styles,
            sheetOrder: this.sheetOrder,
            sheets: this.sheets,
            resources: this.resources,
        };
    }
}
