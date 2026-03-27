import {
    DrawingTypeEnum,
    IStyleData,
    IWorkbookData,
    IWorksheetData,
    LocaleType,
    Nullable,
    PresetGeometryType,
} from '@univerjs/core';
import { IResources } from '@univerjs/core/lib/types/services/resource-manager/type';
import { HyperLink, UniverSheet } from './UniverSheet';
// import { ISheetDrawing, SheetDrawingAnchorType } from '@univerjs/sheets-drawing';
import { handleStyle } from './utils';
import { generateRandomId } from '../common/method';
import { IluckySheet, ILuckyFile, IWorkBookInfo } from '../ToLuckySheet/ILuck';
import { ImageSourceType } from './ILuckInterface';
import { handleRanges } from '../ToLuckySheet/style';
import type { TransformExcelToUniverOptions } from '../main';
import { createProfileLogger } from '../common/profile';

interface Sheets {
    [sheetId: string]: Partial<IWorksheetData & { hyperLink: HyperLink[] }>;
}
interface LuckySheetObj {
    [sheetId: string]: Partial<IluckySheet>;
}
export class UniverWorkBook implements IWorkbookData {
    id!: string;
    rev?: number | undefined;
    name!: string;
    appVersion!: string;
    locale!: LocaleType;
    styles!: Record<string, Nullable<IStyleData>>;
    sheetOrder!: string[];
    sheets!: Sheets;
    resources?: IResources | undefined = [];
    constructor(file: ILuckyFile, options: TransformExcelToUniverOptions = {}) {
        const { info, sheets, workbook } = file;
        const profiler = createProfileLogger(options.profile, 'UniverWorkBook');
        this.id = generateRandomId(6);
        this.name = info.name;
        this.appVersion = info.appversion;
        this.locale = LocaleType.ZH_CN;
        const sourceSheetsByName = sheets.reduce((map, sheet) => {
            map[sheet.name] = sheet;
            return map;
        }, {} as LuckySheetObj);

        const workSheets: Sheets = {},
            order: string[] = [],
            sheetsObj: LuckySheetObj = {};
        sheets
            .sort((a, b) => Number(a.order) - Number(b.order))
            .forEach((d) => {
                const sheetProfiler = createProfileLogger(options.profile, `UniverWorkBook.sheet:${d.name}`);
                const sheet = new UniverSheet(d);
                workSheets[sheet.id] = sheet.mode;
                sheetsObj[sheet.id] = d;
                order.push(sheet.id);
                sheetProfiler.end({
                    rowCount: sheet.rowCount,
                    columnCount: sheet.columnCount,
                    cellRows: Object.keys(sheet.cellData || {}).length,
                });
            });

        // console.log(workSheets,sheets)
        this.handleHyperLinks(workSheets);
        profiler.mark('hyperlinks mapped');
        this.handleImage(workSheets, sourceSheetsByName);
        profiler.mark('images mapped');
        if (options.includeCharts !== false) {
            this.handleChart(workSheets, sourceSheetsByName);
            profiler.mark('charts mapped');
        } else {
            profiler.mark('charts skipped');
        }
        this.handleNames(workbook);
        profiler.mark('defined names mapped');
        this.handleCondition(sheetsObj);
        profiler.mark('conditional formatting mapped');
        this.handleVerification(sheetsObj);
        profiler.mark('data validation mapped');
        this.handleFilter(sheetsObj);
        profiler.mark('filters mapped');
        this.sheetOrder = order;

        this.sheets = workSheets;
        profiler.end({
            sheetCount: order.length,
            resourceCount: this.resources?.length || 0,
        });
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
    private handleHyperLinks = (workSheets: Sheets) => {
        const hyperLinks: { [key: string]: HyperLink[] } = {};
        const sheetIdByName = Object.values(workSheets).reduce((map, sheet) => {
            map[sheet.name || ''] = sheet.id || '';
            return map;
        }, {} as Record<string, string>);
        for (const key in workSheets) {
            const link = workSheets[key].hyperLink;
            if (!link?.length) continue;
            hyperLinks[key] = link.map((d: any) => {
                let payload = d.payload;
                if (typeof d.payload !== 'string') {
                    payload = '#';
                    const gid = d.payload.gid.replace(/'|"/g, '');
                    const sheetId = sheetIdByName[gid];

                    if (gid && sheetId) {
                        payload += `gid=${sheetId}`;
                    }
                    if (gid && sheetId && d.payload.range) payload += '&';
                    if (d.payload.range) payload += `range=${d.payload.range}`;
                }
                return {
                    ...d,
                    payload,
                };
            });
        }
        // console.log(workSheets, hyperLinks)
        this.resources?.push({
            name: 'SHEET_HYPER_LINK_PLUGIN',
            data: JSON.stringify(hyperLinks),
        });
    };
    private handleImage = (workSheets: Sheets, sheetsByName: LuckySheetObj) => {
        const drawerList: {
            [key: string]: { order: string[]; data: { [key: string]: any } };
        } = {};

        Object.values(workSheets).forEach((sheet) => {
            const images = sheetsByName[sheet.name || '']?.images;
            if (!images) return;
            const order = Object.keys(images);
            const data: { [key: string]: any } = {};
            order.forEach((key) => {
                const image = images[key];
                if (sheet.columnCount < image.toCol) {
                    sheet.columnCount = image.toCol + 1;
                }
                if (sheet.rowCount < image.toRow) {
                    sheet.rowCount = image.toRow + 1;
                }
                let imageObj: any = {
                    unitId: this.id,
                    subUnitId: sheet.id || '',
                    drawingId: key,
                    transform: {
                        width: 0,
                        height: 0,
                        scaleX: 0,
                        scaleY: 0,
                        left: 0,
                        top: 0,
                        angle: 0,
                        skewX: 0,
                        skewY: 0,
                        flipX: false,
                        flipY: false,
                        ...(image.transform || {}),
                    },
                    sheetTransform: {
                        angle: 0,
                        skewX: 0,
                        skewY: 0,
                        flipX: false,
                        flipY: false,
                        from: {
                            column: image.fromCol,
                            columnOffset: image.fromColOff,
                            row: image.fromRow,
                            rowOffset: image.fromRowOff,
                        },
                        to: {
                            column: image.toCol,
                            columnOffset: image.toColOff,
                            row: image.toRow,
                            rowOffset: image.toRowOff,
                        },
                    },
                }
                if (image.type === 'chart') {
                    imageObj = {
                        ...imageObj,
                        drawingType: DrawingTypeEnum.DRAWING_CHART,
                        componentKey: 'Chart',
                        data: {
                            ...(image.data || {}),
                            range: `${sheet.name}!${image.data.range}`
                        },
                        allowTransform: true
                    }
                } else {
                    imageObj = {
                        ...imageObj,
                        drawingType: DrawingTypeEnum.DRAWING_IMAGE,
                        imageSourceType: ImageSourceType.BASE64,
                        source: image.src,
                        prstGeom: 'rect' as Nullable<PresetGeometryType>,
                        anchorType: '1',
                    }
                }

                data[key] = imageObj;
            });
            drawerList[sheet.id!] = {
                data,
                order,
            };
        });
        this.resources?.push({
            name: 'SHEET_DRAWING_PLUGIN',
            data: JSON.stringify(drawerList),
        });
    };

    private handleChart = (workSheets: Sheets, sheetsByName: LuckySheetObj) => {
        const chartList: {
            [key: string]: any
        } = {};
        Object.values(workSheets).forEach((sheet) => {
            const charts = sheetsByName[sheet.name || '']?.charts;
            if (!charts) return;
            charts.forEach((chart) => {
                if (!chartList[sheet.id!]) {
                    chartList[sheet.id!] = []
                }
                chartList[sheet.id!].push({
                    rangeInfo: {
                        isRowDirection: chart.isRowDirection,
                        rangeInfo: {
                            unitId: this.id,
                            subUnitId: sheet.id || '',
                            range: handleRanges(chart.range)[0]
                        }
                    },
                    id: chart.id,
                    chartType: chart.chartType,
                    context: chart.context,
                    style: chart.style,
                    dataAggregation: {}
                })
            });
        })
        // console.log('chartList', chartList)
        this.resources?.push({
            name: 'SHEET_CHART_PLUGIN',
            data: JSON.stringify(chartList),
        });
    }
    private handleNames = (workbook: IWorkBookInfo) => {
        this.resources?.push({
            name: 'SHEET_DEFINED_NAME_PLUGIN',
            data: JSON.stringify(workbook.defineNames),
        });
    };
    private handleCondition = (sheets: LuckySheetObj) => {
        const obj: any = {};
        Object.keys(sheets).forEach((d) => {
            const condition = sheets[d].conditionalFormatting?.map((d: any) => {
                if (d.rule?.style) {
                    d.rule.style = handleStyle(
                        { v: d.rule.style, r: 0, c: 0 },
                        { value: d.rule?.style?.border, rangeType: '' }
                    );
                }
                return d;
            });
            obj[d] = condition;
        });
        this.resources?.push({
            name: 'SHEET_CONDITIONAL_FORMATTING_PLUGIN',
            data: JSON.stringify(obj),
        });
    };

    private handleVerification = (sheets: LuckySheetObj) => {
        const obj: any = {};
        Object.keys(sheets).forEach((d) => {
            obj[d] = sheets[d].dataVerificationList;
        });
        this.resources?.push({
            name: 'SHEET_DATA_VALIDATION_PLUGIN',
            data: JSON.stringify(obj),
        });
    };
    private handleFilter = (sheets: LuckySheetObj) => {
        const obj: any = {};
        Object.keys(sheets).forEach((d) => {
            obj[d] = sheets[d].filter;
        });
        this.resources?.push({
            name: 'SHEET_FILTER_PLUGIN',
            data: JSON.stringify(obj),
        });
    };
}
