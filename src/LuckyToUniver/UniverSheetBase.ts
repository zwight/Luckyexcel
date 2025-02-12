import {
    BooleanNumber,
    ICellData,
    IColumnData,
    IFreeze,
    IObjectArrayPrimitiveType,
    IObjectMatrixPrimitiveType,
    IRange,
    IRowData,
    IWorksheetData,
    SheetTypes,
} from '@univerjs/core';
export interface UniverSheetBaseParams {
    id?: string;
    name?: string;
    cellData?: IObjectMatrixPrimitiveType<ICellData>;
    rowCount?: number;
    colCount?: number;
}
export class UniverSheetBase implements IWorksheetData {
    id: string;
    name!: string;
    type: SheetTypes = SheetTypes.GRID;
    tabColor: string = '';
    hidden: BooleanNumber = 0;
    freeze: IFreeze = {
        xSplit: 0,
        ySplit: 0,
        startRow: -1,
        startColumn: -1,
    };
    rowCount: number = 100;
    columnCount: number = 20;
    zoomRatio: number = 1;
    scrollTop: number = 0;
    scrollLeft: number = 0;
    defaultColumnWidth: number = 93;
    defaultRowHeight: number = 27;
    mergeData: IRange[] = [];
    cellData: IObjectMatrixPrimitiveType<ICellData> = {};
    rowData: IObjectArrayPrimitiveType<Partial<IRowData>> = {};
    columnData: IObjectArrayPrimitiveType<Partial<IColumnData>> = {};
    rowHeader: { width: number; hidden?: BooleanNumber | undefined } = {
        width: 46,
        hidden: 0,
    };
    columnHeader: { height: number; hidden?: BooleanNumber | undefined } = {
        height: 20,
        hidden: 0,
    };
    showGridlines: BooleanNumber = 1;
    rightToLeft: BooleanNumber = 0;
    selections: string[] = [];

    constructor(params?: UniverSheetBaseParams) {
        const { id, name, cellData, rowCount = 0, colCount = 0 } = params || {};
        this.id = id || '';
        this.name = name || '';
        this.cellData = cellData || {};

        this.rowCount = Math.max(this.rowCount, rowCount);
        this.columnCount = Math.max(this.columnCount, colCount);
    }
}
