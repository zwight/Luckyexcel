import { IStyleData, IWorkbookData, IWorksheetData, LocaleType, Nullable } from '@univerjs/core';
import { IResources } from '@univerjs/core/lib/types/services/resource-manager/type';
import { UniverSheetBase } from './UniverSheetBase';

export class UniverCsvWorkBook implements IWorkbookData {
    id!: string;
    rev?: number | undefined;
    name!: string;
    appVersion!: string;
    locale!: LocaleType;
    styles!: Record<string, Nullable<IStyleData>>;
    sheetOrder!: string[];
    sheets!: { [sheetId: string]: Partial<IWorksheetData> };
    resources?: IResources | undefined;
    constructor(data: string[][]) {
        const cellData = data.map((row) => {
            return row.map((cell) => {
                return {
                    v: cell || '',
                };
            });
        });
        const id = `sheet1`;
        const sheet = new UniverSheetBase(id, id, cellData);
        this.sheets = { [id]: sheet };
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
