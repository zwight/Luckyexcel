import {
    BooleanNumber,
    ICellData,
    IRange,
    IObjectMatrixPrimitiveType,
    IObjectArrayPrimitiveType,
    IRowData,
    IColumnData,
    CellValueType,
    Nullable,
    IDocumentData,
} from '@univerjs/core';
import { UniverSheetBase } from './UniverSheetBase';
import { handleStyle, removeEmptyAttr } from './utils';
import { str2num, generateRandomId } from '../common/method';
import { IluckySheet, IluckySheetConfig, IluckySheetCelldata, IluckysheetHyperlink, IluckysheetFrozen, IluckySheetCelldataValue } from '../ToLuckySheet/ILuck';

export interface HyperLink {
    id: string;
    payload: string | { gid: string; range: string };
    row: number;
    column: number;
}
export interface UniverSheetMode extends UniverSheetBase {
    hyperLink: HyperLink[];
    mode: UniverSheetMode;
}
export class UniverSheet extends UniverSheetBase {
    hyperLink: HyperLink[] = [];
    constructor(sheetData: IluckySheet) {
        super();
        const {
            color,
            zoomRatio,
            celldata,
            config = {} as IluckySheetConfig,
            showGridLines,
            defaultColWidth,
            defaultRowHeight,
            hide,
        } = sheetData || {};
        this.name = sheetData.name;
        this.id = `sheet-${sheetData.index}`;
        if (sheetData) {
            this.tabColor = color;
            this.zoomRatio = zoomRatio;
            this.showGridlines = Number(showGridLines);
            this.defaultColumnWidth = defaultColWidth;
            this.defaultRowHeight = defaultRowHeight;
            this.hidden = hide;
            this.handleSheetLink(sheetData.hyperlink);

            if (config.merge) this.mergeData = this.handleMerge(config);

            if (celldata?.length) {
                const { cellData, rowCount, colCount } = this.handleCellData(celldata, config);
                this.cellData = cellData;
                this.rowCount = this.rowCount > rowCount ? this.rowCount : rowCount + 10;
                this.columnCount = this.columnCount > colCount ? this.columnCount : colCount + 5;
            }
            console.log(this.rowCount, this.columnCount)
            this.handleRowAndColumnData(config);
            if (sheetData.freezen) this.handleFreeze(sheetData.freezen);
        }
    }
    get mode(): Omit<UniverSheetMode, 'mode'> {
        return {
            id: this.id,
            name: this.name,
            type: this.type,
            tabColor: this.tabColor,
            hidden: this.hidden,
            freeze: this.freeze,
            rowCount: this.rowCount,
            columnCount: this.columnCount,
            zoomRatio: this.zoomRatio,
            scrollTop: this.scrollTop,
            scrollLeft: this.scrollLeft,
            defaultColumnWidth: this.defaultColumnWidth,
            defaultRowHeight: this.defaultRowHeight,
            mergeData: this.mergeData,
            cellData: this.cellData,
            rowData: this.rowData,
            columnData: this.columnData,
            rowHeader: this.rowHeader,
            columnHeader: this.columnHeader,
            showGridlines: this.showGridlines,
            rightToLeft: this.rightToLeft,
            selections: this.selections,
            hyperLink: this.hyperLink,
        };
    }
    private handleMerge = (config: IluckySheetConfig): IRange[] => {
        const merges = config.merge;
        if (!merges) return [];
        return Object.values(merges).map((merge) => {
            return {
                startRow: merge.r,
                endRow: merge.r + merge.rs - 1,
                startColumn: merge.c,
                endColumn: merge.c + merge.cs - 1,
            };
        });
    };
    private handleCellData = (celldata: IluckySheetCelldata[], config: IluckySheetConfig) => {
        const handleCell = (row: IluckySheetCelldata): ICellData => {
            const { v } = row;
            if (typeof v === 'string' || v === null || v === undefined) {
                return { v: v as string };
            }
            const tMap: any = {
                s: CellValueType.STRING,
                n: CellValueType.NUMBER,
                b: CellValueType.BOOLEAN,
                str: CellValueType.STRING,
            };
            const borderConf = config.borderInfo?.find(
                (d) => d.value.col_index === row.c && d.value.row_index === row.r
            );

            let cellType = v.ct?.t && tMap[v.ct?.t] ? tMap[v.ct?.t] : CellValueType.NUMBER;

            let val = cellType === CellValueType.NUMBER ? str2num(v.v) : v.v;
            if (cellType === CellValueType.BOOLEAN) val = v.v ? 1 : 0;

            if (Number.isNaN(Number(val)) && cellType === CellValueType.NUMBER)
                cellType = CellValueType.STRING;
            if (this.hyperLink.findIndex((d) => d.column === row.c && d.row === row.r) > -1)
                cellType = CellValueType.STRING;

            const cell: ICellData = {
                // custom: v., // User stored custom fields
                f: v.f,
                // p: , // The unique key, a random string, is used for the plug-in to associate the cell. When the cell information changes, the plug-in does not need to change the data, reducing the pressure on the back-end interface id?: string.
                s: handleStyle(row, borderConf),
                si: v.f, // Id of the formula.
                t: cellType,
                v: val,
            };
            const pVal = this.handleDocument(row, config);
            if (pVal) cell.p = pVal;
            return removeEmptyAttr(cell);
        };
        let row: number | undefined = undefined;
        let colCount = 0;
        const rowData = celldata.reduce((pre: any, cur) => {
            if (row === cur.r) {
                pre[cur.r].push(cur);
            } else {
                row = cur.r;
                pre[row] = [cur];
            }
            if (cur.c > colCount) colCount = cur.c;
            return pre;
        }, []);
        const cell: IObjectMatrixPrimitiveType<ICellData> = {};
        // console.log(rowData, celldata, colCount)
        rowData.forEach((row: IluckySheetCelldata[], rowIndex: number) => {
            for (let index = 0; index < colCount + 1; index++) {
                const element = row.find((d) => d.c === index) || {
                    r: rowIndex,
                    c: index,
                    v: null,
                };
                if (!cell[element.r]) cell[element.r] = {};
                cell[element.r][element.c] = handleCell(element);
            }
        });
        return {
            cellData: cell,
            rowCount: rowData.length,
            colCount,
        };
    };
    private handleDocument = (row: IluckySheetCelldata, config: IluckySheetConfig) => {
        const matchArray = (str: string, charToFind: string) => {
            const regex = new RegExp(charToFind, 'g');
            let match;
            const indices = [];

            while ((match = regex.exec(str))) {
                indices.push(match.index);
            }

            return indices;
        };
        const removeLastChar = (str: string, charToRemove: string) => {
            const regex = new RegExp(`${charToRemove}$`);
            return str.replace(regex, '');
        };
        let pVlaue: Nullable<IDocumentData> = null;
        const { v } = row;
        if (typeof v === 'string' || v === null || v === undefined) {
            return undefined;
        }
        if (v.ct && v.ct.t === 'inlineStr') {
            let dataStream = v.ct.s.reduce((prev, cur) => {
                return prev + removeLastChar(cur.v || '', '\r\n');
            }, '');
            dataStream = dataStream?.replace(/\n/g, '\r') + '\r\n';
            const matchChart = {
                r: '\r', // PARAGRAPH
                n: '\n', // SECTION_BREAK
                v: '\v', // COLUMN_BREAK
                f: '\f', // PAGE_BREAK
                '0': '\0', // DOCS_END
                t: '\t', // TAB
                b: '\b', // customBlock
                x1A: '\x1A', // table start
                x1B: '\x1B', // table row start
                x1C: '\x1C', // table cell start
                x1D: '\x1D', // table cell end
                x1E: '\x1E', // table row end || customRange end
                x1F: '\x1F', // table end || customRange start
            };
            const paragraphs = matchArray(dataStream, matchChart.r).map((d) => {
                return {
                    startIndex: d,
                };
            });
            const sectionBreaks = matchArray(dataStream, matchChart.n).map((d) => {
                return {
                    startIndex: d,
                };
            });
            const textRuns = v.ct.s?.map((d, index) => {
                const start = v.ct!.s?.reduce((prev, cur, curi) => {
                    if (curi < index) return prev + (cur.v?.length || 0);
                    return prev;
                }, 0);
                const end = start + (v.ct!.s?.[index]?.v?.length || 0);
                const borderConf = config.borderInfo?.find(
                    (d) => d.value.col_index === row.c && d.value.row_index === row.r
                );
                return {
                    st: start,
                    ed: end,
                    ts: handleStyle(
                        {
                            v: (v.ct!.s[index] || v.ct!.s[0]) as IluckySheetCelldataValue,
                            r: row.r,
                            c: row.c,
                        },
                        borderConf,
                        true
                    ),
                };
            });
            pVlaue = {
                id: generateRandomId(6),
                documentStyle: {
                    documentFlavor: 0,
                    pageSize: { width: 0, height: 0 },
                    renderConfig: {},
                    textStyle: {},
                },
                body: {
                    dataStream,
                    paragraphs,
                    sectionBreaks,
                    textRuns,
                },
                drawings: {},
            };
        }
        return pVlaue;
    };

    private handleRowAndColumnData = (config: IluckySheetConfig) => {
        const columnData: IObjectArrayPrimitiveType<Partial<IColumnData>> = {};
        const rowData: IObjectArrayPrimitiveType<Partial<IRowData>> = {};
        for (let index = 0; index < this.rowCount; index++) {
            rowData[index] = {
                h: config.rowlen?.[index] || this.defaultRowHeight,
                ia: !config.rowlen?.[index] ? BooleanNumber.TRUE : BooleanNumber.FALSE,
                ah: this.defaultRowHeight,
                hd: config.rowhidden?.[index] === 0 ? BooleanNumber.TRUE : BooleanNumber.FALSE,
            };
        }

        for (let index = 0; index < this.columnCount; index++) {
            columnData[index] = {
                w: config.columnlen?.[index] || this.defaultColumnWidth,
                hd: config.colhidden?.[index] === 0 ? BooleanNumber.TRUE : BooleanNumber.FALSE,
            };
        }
        this.rowData = rowData;
        this.columnData = columnData;
    };

    /**
     * 处理链接
     * @param sheetName IluckysheetHyperlink
     */
    private handleSheetLink = (hyperlinks: IluckysheetHyperlink) => {
        if (!hyperlinks) return;
        const links = Object.keys(hyperlinks).map((d) => {
            const row = d.split('_')[0],
                column = d.split('_')[1];

            const item = hyperlinks[d];
            let payload: any = item.linkAddress;
            if (item.linkType === 'internal') {
                const locationList = item.linkAddress.split('!');
                payload = {};
                if (locationList[0]) payload['gid'] = locationList[0];
                if (locationList[1]) payload['range'] = locationList[1];
            }
            return {
                id: generateRandomId(6),
                row: Number(row),
                column: Number(column),
                payload,
            };
        });
        this.hyperLink = links;
    };
    
    private handleFreeze = (freeze: IluckysheetFrozen) => {
        this.freeze = {
            xSplit: freeze.vertical,
            ySplit: freeze.horizen,
            startColumn: freeze.vertical,
            startRow: freeze.horizen,
        };
    };
}
