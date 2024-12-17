import { LuckyFile } from "./ToLuckySheet/LuckyFile";
// import {SecurityDoor,Car} from './content';

import { HandleZip } from './HandleZip';

import { IuploadfileList } from "./ICommon";

import { WorkBook } from "./UniverToExcel/Workbook";
import exceljs from "@zwight/exceljs";
import { CSV } from "./UniverToCsv/CSV";
import { isObject } from "./common/method";
import { UniverWorkBook } from "./LuckyToUniver/UniverWorkBook";
import { IWorkbookData } from "@univerjs/core";
export class LuckyExcel {
    constructor() { }
    static transformExcelToLucky(excelFile: File,
        callback?: (files: IuploadfileList, fs?: string) => void,
        errorHandler?: (err: Error) => void) {
        let handleZip: HandleZip = new HandleZip(excelFile);

        // const fileReader = new FileReader();
        // fileReader.onload = async (e) => {
        //     const { result } = e.target as any;
        //     const workbook = new exceljs.Workbook();
        //     const data = await workbook.xlsx.load(result);
        //     // console.log('exceljs', data)
        // }
        // fileReader.readAsArrayBuffer(excelFile)

        handleZip.unzipFile(function (files: IuploadfileList) {
            let luckyFile = new LuckyFile(files, excelFile.name);
            let luckysheetfile = luckyFile.Parse();
            let exportJson = JSON.parse(luckysheetfile);
            // console.log('output---->', exportJson)
            if (callback != undefined) {
                callback(exportJson, luckysheetfile);
            }
        },
            function (err: Error) {
                if (errorHandler) {
                    errorHandler(err);
                } else {
                    console.error(err);
                }
            });
    }

    static transformExcelToLuckyByUrl(
        url: string,
        name: string,
        callBack?: (files: IuploadfileList, fs?: string) => void,
        errorHandler?: (err: Error) => void) {
        let handleZip: HandleZip = new HandleZip();
        handleZip.unzipFileByUrl(url, function (files: IuploadfileList) {
            let luckyFile = new LuckyFile(files, name);
            let luckysheetfile = luckyFile.Parse();
            let exportJson = JSON.parse(luckysheetfile);
            if (callBack != undefined) {
                callBack(exportJson, luckysheetfile);
            }
        },
            function (err: Error) {
                if (errorHandler) {
                    errorHandler(err);
                } else {
                    console.error(err);
                }
            });
    }


    static transformExcelToUniver(excelFile: File,
        callback?: (files: IWorkbookData, fs?: string) => void,
        errorHandler?: (err: Error) => void) {
        let handleZip: HandleZip = new HandleZip(excelFile);

        handleZip.unzipFile(function (files: IuploadfileList) {
            let luckyFile = new LuckyFile(files, excelFile.name);
            let luckysheetfile = luckyFile.Parse();
            let exportJson = JSON.parse(luckysheetfile);
            console.log('output---->', exportJson, files)
            if (callback != undefined) {
                const univerData = new UniverWorkBook(exportJson)
                callback(univerData.mode, luckysheetfile);
            }
        },
            function (err: Error) {
                if (errorHandler) {
                    errorHandler(err);
                } else {
                    console.error(err);
                }
            });
    }

    static async transformUniverToExcel(params: {
        snapshot: any,
        fileName?: string,
        getBuffer?: boolean,
        success?: (buffer?: exceljs.Buffer) => void,
        error?: (err: Error) => void
    }) {
        const { snapshot, fileName = `excel_${(new Date).getTime()}.xlsx`, getBuffer = false, success, error } = params;
        try {
            // console.log(1, new Date())
            const workbook = new WorkBook(snapshot);
            // console.log(snapshot, workbook)
            // console.log(2, new Date())
            const buffer = await workbook.xlsx.writeBuffer();
            // console.log(3, new Date())
            if (getBuffer) {
                success?.(buffer);
            } else {
                this.downloadFile(fileName, buffer);
                success?.()
            }

        } catch (err) {
            error?.(err)
        }
    }

    static async transformUniverToCsv(params: {
        snapshot: any,
        fileName?: string,
        getBuffer?: boolean,
        sheetName?: string,
        success?: (csvContent?: string | { [key: string]: string }) => void,
        error?: (err: Error) => void
    }) {
        const { snapshot, fileName = `csv_${(new Date).getTime()}.csv`, getBuffer = false, success, error, sheetName } = params;
        try {
            const csv = new CSV(snapshot);
            console.log(csv);

            let contents: string | { [key: string]: string };
            if (sheetName) {
                contents = csv.csvContent[sheetName];
            } else {
                contents = csv.csvContent;
            }
            if (getBuffer) {
                success?.(contents);
            } else {
                if (isObject(contents)) {
                    for (const key in contents) {
                        if (Object.prototype.hasOwnProperty.call(contents, key)) {
                            const element = contents[key];
                            this.downloadFile(`${fileName}_${key}`, element);
                        }
                    }
                } else {
                    this.downloadFile(fileName, contents);
                }
                success?.()
            }
        } catch (err) {
            error(err)
        }
    }
    
    private static downloadFile(fileName: string, buffer: exceljs.Buffer | string) {
        const link = document.createElement('a');
        let blob: Blob;
        if (typeof buffer === 'string') {
            blob = new Blob([buffer], { type: "text/csv;charset=utf-8;" });
        } else {
            blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8' });
        }

        const url = URL.createObjectURL(blob);
        link.href = url;
        link.download = fileName;
        document.body.appendChild(link);
        link.click();

        link.addEventListener('click', () => {
            link.remove();
            setTimeout(() => {
                URL.revokeObjectURL(url)
            }, 200);
        })
    }
}