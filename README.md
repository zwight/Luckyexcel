English| [简体中文](./README-zh.md)

>Warning:
this project was forked from [Luckyexcel](https://github.com/dream-num/Luckyexcel), base on the last commit on 2022-06-10 [5a0be428b9fead1479a0890d19fb26ae0a291a1c](https://github.com/dream-num/Luckyexcel/commit/5a0be428b9fead1479a0890d19fb26ae0a291a1c)

## Introduction
This project is based on the import of [Luckyexcel](https://github.com/dream-num/Luckyexcel), and adds the conversion of [Luckysheet](https://github.com/mengshukeji/Luckysheet) data structure into [Univer](https://github.com/dream-num/univer) data structure. It can directly import and return the data structure required by Univer. In addition, this project also implements the export function based on Univer, supporting the export of .xlsx and .csv format files

## Features
Support Univer import excel and export excel/csv adapter list

- Cell style
- Cell border
- Cell format, such as number format, date, percentage, etc.
- Formula
- Conditional Formatting
- Sort
- Filter

### Plan
The goal is to support all features supported by Univer

- Pivot table
- Chart
- Annotation

## Usage

### CDN
```html
<script src="https://cdn.jsdelivr.net/npm/@zwight/luckyexcel/dist/luckyexcel.umd.min.js"></script>
<script>
    // Univer import Excel file
    LuckyExcel.transformExcelToUniver(
        file,
        async (exportJson: any) => {
            // After obtaining the converted table data, use univer to initialize, or update the existing univer workbook
            // Note: Univer needs to introduce dependent packages and initialize the table container before it can be used
            univer.createUnit<IWorkbookData, Workbook>(
                UniverInstanceType.UNIVER_SHEET,
                exportJson || {}
            );
        },
        (error: any) => {
            console.log(error);
        }
    );
    // Export Univer to CSV file
    // snapshot is the Univer snapshot data, getBuffer: true will not download the file, only return the csv content, false will download directly
    // sheetName: Because Univer may have multiple sheets, csv does not have sheets, if sheetName has a value, only the data of the specified sheet name will be downloaded. If it is not passed, all sheets will be downloaded. The file name is ${fileName}_${sheet.name}
    LuckyExcel.transformUniverToCsv({
        snapshot,
        fileName,
        getBuffer: true,
        sheetName: snapshot.sheetOrder[0],
        success: (buffer: string | { [key: string]: string }) => {
            console.log('success');
        },
        error: (error: Error) => {
            console.log('error', error);
        },
    });
    // Export Univer to XLSX file
    // getBuffer: true will not download the file, only return the file's buffer data, false will download directly
    LuckyExcel.transformUniverToExcel({
        snapshot,
        fileName,
        getBuffer: true,
        success: (buffer: Buffer) => {
            console.log('success');
        },
        error: (error: Error) => {
            console.log('error', error);
        },
    });
</script>
```
> Case [Demo index.html](./src/index.html) shows the detailed usage

### ES and Node.js

#### Installation
```shell
npm install @zwight/luckyexcel
```

#### ES import
```js
import LuckyExcel from '@zwight/luckyexcel'

// After getting the xlsx file
LuckyExcel.transformExcelToUniver(
    file,
    async (exportJson: any) => {
        // Get the worksheet data after conversion
    },
    (error: any) => {
        // handle error if any thrown
    }
);
```

## Development

### Requirements
[Node.js](https://nodejs.org/en/) Version >= 6 

### Installation
```
npm install -g gulp-cli
npm install
```
### Development
```
npm run dev
```
### Package
```
npm run build
```

A third-party plug-in is used in the project: [JSZip](https://github.com/Stuk/jszip), thanks!

## Authors and acknowledgment
- [@dream-num](https://github.com/dream-num)

## License
[MIT](http://opensource.org/licenses/MIT)

Copyright (c) 2020-present, zwight
