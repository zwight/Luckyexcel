import { Workbook, Worksheet, WorksheetViewCommon, WorksheetViewFrozen } from "@zwight/exceljs";
import { convertSheetIdToName, heightConvert, hex2argb, wdithConvert } from "./util";
import { cellStyle, fontConvert } from "./CellStyle";
import { jsonParse, removeEmptyAttr } from "../common/method";
import { Resource } from "./Resource";
export class ViewCommon implements WorksheetViewCommon{
    rightToLeft: boolean;
    activeCell: string;
    showRuler: boolean;
    showRowColHeaders: boolean;
    showGridLines: boolean;
    zoomScale: number;
    zoomScaleNormal: number;
}
export class FrozenView implements WorksheetViewFrozen{
    state: "frozen";
    style?: "pageBreakPreview";
    xSplit?: number;
    ySplit?: number;
    topLeftCell?: string;
}
export function ExcelWorkSheet(workbook: Workbook, snapshot: any) {
    const { sheetOrder, sheets, styles, resources } = snapshot;
    sheetOrder.forEach((sheetId: string) => {
        const sheet = sheets[sheetId];
        const { 
            id,
            name, 
            tabColor, 
            defaultRowHeight, 
            defaultColumnWidth, 
            hidden,
            rightToLeft,
            showGridlines,
            freeze,
            mergeData
        } = sheet;
        const commonView = new ViewCommon();
        commonView.rightToLeft = rightToLeft === 1;
        commonView.showGridLines = showGridlines === 1;
        const frozenView = new FrozenView();
        if (freeze.xSplit > 0 || freeze.ySplit > 0){
            frozenView.state = 'frozen';
            frozenView.xSplit = freeze.xSplit;
            frozenView.ySplit = freeze.ySplit;
        }
        const views = Object.assign(commonView, frozenView)

        const defaultColWidth = wdithConvert(defaultColumnWidth);
        const defaultRowHeightR = heightConvert(defaultRowHeight);
        const worksheet = workbook.addWorksheet(name, {
            views: [views],
            state: hidden === 1 ? 'hidden' : 'visible',
            properties: {
                tabColor: tabColor ? { argb: hex2argb(tabColor) } : undefined,
                defaultColWidth: defaultColWidth,
                defaultRowHeight: defaultRowHeightR,
                dyDescent: 0
            }
        })
        setColumns(worksheet, sheet.columnData, defaultColWidth);
        setRows(worksheet, sheet.rowData, defaultRowHeightR)
        setCell(worksheet, sheet, styles, snapshot);
        setMerges(worksheet, mergeData);
        
        new Resource(id, workbook, worksheet, resources);
    });
}


function setMerges(worksheet: Worksheet, mergeData: any[]) {
    mergeData.forEach(d => {
        worksheet.mergeCells(d.startRow + 1, d.startColumn + 1, d.endRow + 1, d.endColumn + 1)
    })
}

function setCell(worksheet: Worksheet, sheet: any, styles: any, snapshot: any) {
    const { resources, sheets } = snapshot;
    const { cellData, id } = sheet;
    for (const rowid in cellData) {
        const row = cellData[rowid];
        for (const columnid in row) {
            const cell = row[columnid];
            if (!cell) continue;
            // console.log(rowid + 1, columnid + 1)
            const target = worksheet.getCell(Number(rowid) + 1, Number(columnid) + 1)

            target.value = handleValue(cell, {
                resources,
                sheetId: id,
                rowId: rowid,
                columnId: columnid,
                sheets
            } );
            
            let originStyle = cell.s;
            if (typeof cell.s === 'string') {
                originStyle = styles[cell.s]
            }
            const style = removeEmptyAttr(cellStyle(originStyle, originStyle?.n?.pattern || cell.f))
            Object.assign(target, style)
            // console.log(target)
        }
    }
}
function getHyperLink(cellSource: any) {
    const { resources, sheetId, rowId, columnId } = cellSource;
    const hyperlinks = jsonParse(resources.find((d: any) => d.name === 'SHEET_HYPER_LINK_PLUGIN')?.data);
    const list = hyperlinks?.[sheetId] || [];
    const hyperlink = list.find((d: any) => d.row === Number(rowId) && d.column === Number(columnId));
    return hyperlink
}

function handleValue(cell: any, cellSource: any) {
    const { sheets } = cellSource
    const hyperlink = getHyperLink(cellSource)
    let value;
    if (cell.p) {
        const body = cell.p?.body;
        value = {
            richText: body?.textRuns.map((d: any) => {
                return {
                    text: body.dataStream.substring(d.st, d.ed),
                    font: fontConvert(d.ts)
                }
            })
        }
    } else if (cell.si) {
        value = { formula: cell.si, result: cell.v }
    } else {
        value = cell.v
    }
    if (hyperlink) {
        const { payload } = hyperlink;
        let link = '';
        let model = '';
        if (payload.includes('#gid=') || payload.includes('range=')) {
            const str = payload.replace('#', '');
            const arr = str.split('&');
            link += '';
            if (arr.length === 1 && arr[0].includes('range=')) {
                link += arr[0].replace('range=')
            }
            if (arr.length === 2) {
                link += `\'${convertSheetIdToName(sheets, arr[0].replace('gid=', ''))}\'`;
                link += `!${arr[1].replace('range=', '')}`
            }
        } else {
            link = payload
            model = 'External'
        }
        const text = value?.richText?.map?.((d: any) => d.text)?.join('') || value?.result || value;
        if (link) value = {
            text: text,
            hyperlink: link,
            hyperlinkModel: model
        }
    }
    return value
}

function setColumns(worksheet: Worksheet, columnData: any = {}, defaultColumnWidth: number) {
    for (const key in columnData) {
        if (Object.prototype.hasOwnProperty.call(columnData, key)) {
            const element = columnData[key];
            const column = worksheet.getColumn(Number(key) + 1)
            column.width = element.w ? wdithConvert(element.w) : defaultColumnWidth;
            column.hidden = element.hd === 1;
        }
    }
}
function setRows(worksheet: Worksheet, rowData: any = {}, defaultRowHeight: number) {
    for (const key in rowData) {
        if (Object.prototype.hasOwnProperty.call(rowData, key)) {
            const element = rowData[key];
            const row = worksheet.getRow(Number(key) + 1)
            row.height = element.h ? heightConvert(element.h) : defaultRowHeight;
            row.hidden = element.hd === 1;
        }
    }
}