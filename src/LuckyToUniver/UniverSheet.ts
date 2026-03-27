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
    PositionedObjectLayoutType,
    DrawingTypeEnum,
} from '@univerjs/core';
import { UniverSheetBase } from './UniverSheetBase';
import { handleStyle, removeEmptyAttr } from './utils';
import { str2num, generateRandomId } from '../common/method';
import { IluckySheet, IluckySheetConfig, IluckySheetCelldata, IluckysheetHyperlink, IluckysheetFrozen, IluckySheetCelldataValue } from '../ToLuckySheet/ILuck';
import { ImageSourceType } from './ILuckInterface';

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
                this.rowCount = this.rowCount > rowCount ? this.rowCount : rowCount + 1;
                this.columnCount = this.columnCount > colCount ? this.columnCount : colCount + 1;
            }
            const maxConfiguredRow = Math.max(
                0,
                ...Object.keys(config.rowlen || {}).map(Number),
                ...Object.keys(config.rowhidden || {}).map(Number)
            );
            const maxConfiguredColumn = Math.max(
                0,
                ...Object.keys(config.columnlen || {}).map(Number),
                ...Object.keys(config.colhidden || {}).map(Number)
            );
            this.rowCount = Math.max(this.rowCount, maxConfiguredRow + 1);
            this.columnCount = Math.max(this.columnCount, maxConfiguredColumn + 1);
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
        const borderInfoMap = new Map(
            (config.borderInfo || []).map((item) => [
                `${item.value.row_index}_${item.value.col_index}`,
                item,
            ])
        );
        const hyperLinkCells = new Set(
            this.hyperLink.map((item) => `${item.row}_${item.column}`)
        );
        const getBorderConf = (row: number, col: number) =>
            borderInfoMap.get(`${row}_${col}`);
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
            const borderConf = getBorderConf(row.r, row.c);

            let cellType = v.ct?.t && tMap[v.ct?.t] ? tMap[v.ct?.t] : CellValueType.NUMBER;

            let val = cellType === CellValueType.NUMBER ? str2num(v.v) : v.v;
            if (cellType === CellValueType.BOOLEAN) val = v.v == '1' ? 1 : 0;

            if (Number.isNaN(Number(val)) && cellType === CellValueType.NUMBER)
                cellType = CellValueType.STRING;
            if (hyperLinkCells.has(`${row.r}_${row.c}`))
                cellType = CellValueType.STRING;

            const f = v.f?.replace(/=_xlfn./g, '=');
            const cell: ICellData = {
                // custom: v., // User stored custom fields
                f,
                // p: , // The unique key, a random string, is used for the plug-in to associate the cell. When the cell information changes, the plug-in does not need to change the data, reducing the pressure on the back-end interface id?: string.
                s: handleStyle(row, borderConf),
                // si: f, // Id of the formula.
                t: cellType,
                v: val,
            };
            const pVal = this.handleDocument(row, borderInfoMap);
            if (pVal) cell.p = pVal;

            const pValImg = this.handleCellImage(row, borderInfoMap);
            if (pValImg) {
                cell.p = pValImg;
                cell.f = undefined;
                cell.v = undefined;
            }
            return removeEmptyAttr(cell);
        };
        const cell: IObjectMatrixPrimitiveType<ICellData> = {};
        let rowCount = 0;
        let colCount = 0;
        celldata.forEach((item) => {
            if (!cell[item.r]) cell[item.r] = {};
            cell[item.r][item.c] = handleCell(item);
            if (item.r > rowCount) rowCount = item.r;
            if (item.c > colCount) colCount = item.c;
        });
        return {
            cellData: cell,
            rowCount,
            colCount,
        };
    };
    private handleDocument = (
        row: IluckySheetCelldata,
        borderInfoMap: Map<string, any>
    ) => {
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
            const regex = new RegExp(`${charToRemove}`, 'g');
            return str.replace(regex, '\r');
        };
        let pVlaue: Nullable<IDocumentData> = null;
        const { v } = row;
        if (typeof v === 'string' || v === null || v === undefined) {
            return undefined;
        }
        if (v.ct && v.ct.t === 'inlineStr') {

            v.ct.s = v.ct.s?.map(d => {
                d.v = removeLastChar(d.v || '', '\r\n');
                return d
            }) || []

            let dataStream = v.ct.s.reduce((prev, cur) => {
                return prev + cur.v;
            }, '');
            dataStream = dataStream ? dataStream.replace(/\n/g, '\r') + '\r\n' : '';
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
            const textRuns = v.ct.s.map((d, index) => {
                const start = v.ct.s.reduce((prev, cur, curi) => {
                    if (curi < index) return prev + (cur.v?.length || 0);
                    return prev;
                }, 0);
                const end = start + (v.ct!.s?.[index]?.v?.length || 0);
                const borderConf = borderInfoMap.get(`${row.r}_${row.c}`);
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

    private handleCellImage = (
        row: IluckySheetCelldata,
        borderInfoMap: Map<string, any>
    ) => {
        let pVlaue: Nullable<any> = null;
        const { v } = row;
        if (typeof v === 'string' || v === null || v === undefined) {
            return undefined;
        }
        if (v.ct && v.ct.t === 'str' && v.ct.ci) {
            const blockId = generateRandomId(6);
            const valueId = generateRandomId(6);
            const { default: defaultData,  src, descr } = v.ct.ci || {};
            const borderConf = borderInfoMap.get(`${row.r}_${row.c}`);
            pVlaue = {
                id: valueId,
                documentStyle: {
                    documentFlavor: 0,
                    pageSize: { width: 0, height: 0 },
                    renderConfig: {},
                    textStyle: {},
                },
                body: {
                    dataStream: '\b\r\n',
                    paragraphs: [{
                        startIndex: 1,
                        paragraphStyle: { horizontalAlign: v.ht }
                    }],
                    sectionBreaks: [{ startIndex: 2 }],
                    textRuns: [{
                        ed: 1,
                        st: 0,
                        ts: handleStyle(
                            {
                                v: v,
                                r: row.r,
                                c: row.c,
                            },
                            borderConf,
                            true
                        ),
                    }],
                    customBlocks: [{ startIndex: 0, blockId }]
                },
                drawings: {
                    [blockId]: {
                        unitId: valueId,
                        subUnitId: valueId,
                        drawingId: blockId,
                        layoutType: PositionedObjectLayoutType.INLINE,
                        title: '',
                        description: descr,
                        docTransform: {
                            size: {
                                width: defaultData.width,
                                height: defaultData.height
                            },
                            positionH: {
                                relativeFrom: 0,
                                posOffset: 0
                            },
                            positionV: {
                                relativeFrom: 1,
                                posOffset: 0
                            },
                            angle: 0
                        },
                        drawingType: DrawingTypeEnum.DRAWING_IMAGE,
                        imageSourceType: ImageSourceType.BASE64,
                        source: src,
                        transform: defaultData
                    }
                },
                drawingsOrder: [blockId]
            };
        }
        return pVlaue;
    }

    private handleRowAndColumnData = (config: IluckySheetConfig) => {
        const columnData: IObjectArrayPrimitiveType<Partial<IColumnData>> = {};
        const rowData: IObjectArrayPrimitiveType<Partial<IRowData>> = {};
        const rowIndexes = new Set([
            ...Object.keys(config.rowlen || {}),
            ...Object.keys(config.rowhidden || {}),
        ]);
        rowIndexes.forEach((indexTxt) => {
            const index = Number(indexTxt);
            rowData[index] = {
                h: config.rowlen?.[index] || this.defaultRowHeight,
                ia: !config.rowlen?.[index] ? BooleanNumber.TRUE : BooleanNumber.FALSE,
                ah: this.defaultRowHeight,
                hd: config.rowhidden?.[index] === 0 ? BooleanNumber.TRUE : BooleanNumber.FALSE,
            };
        });

        const columnIndexes = new Set([
            ...Object.keys(config.columnlen || {}),
            ...Object.keys(config.colhidden || {}),
        ]);
        columnIndexes.forEach((indexTxt) => {
            const index = Number(indexTxt);
            columnData[index] = {
                w: config.columnlen?.[index] || this.defaultColumnWidth,
                hd: config.colhidden?.[index] === 0 ? BooleanNumber.TRUE : BooleanNumber.FALSE,
            };
        });
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
