简体中文 | [English](./README.md)

>注意:
本项目是fork [Luckyexcel](https://github.com/dream-num/Luckyexcel), 基于2022年6月10日的最后一次提交[5a0be428b9fead1479a0890d19fb26ae0a291a1c](https://github.com/dream-num/Luckyexcel/commit/5a0be428b9fead1479a0890d19fb26ae0a291a1c)

## 介绍

本项目基于[Luckyexcel](https://github.com/dream-num/Luckyexcel)的导入，添加了[Luckysheet](https://github.com/mengshukeji/Luckysheet)数据结构转换为[Univer](https://github.com/dream-num/univer)数据结构，可以直接导入后返回Univer所需的数据结构，并且本项目还实现了基于Univer的导出功能，支持导出.xlsx和.csv格式文件

## 特性
支持Univer导入excel和导出excel/csv适配列表

- 单元格样式
- 单元格边框
- 单元格格式，如数字格式、日期、百分比等
- 公式
- 条件格式
- 排序
- 筛选

### 计划

目标是支持所有Univer支持的特性

- 数据透视表
- 图表
- 批注

## 用法

### CDN
```html
<script src="https://cdn.jsdelivr.net/npm/@zwight/luckyexcel/dist/luckyexcel.umd.min.js"></script>
<script>
    // Univer导入Excel文件
    LuckyExcel.transformExcelToUniver(
        file,
        async (exportJson: any) => {
            // 获得转化后的表格数据后，使用univer初始化，或者更新已有的univer工作簿
            // 注：univer需要引入依赖包和初始化表格容器才可以使用
            univer.createUnit<IWorkbookData, Workbook>(
                UniverInstanceType.UNIVER_SHEET,
                exportJson || {}
            );
        },
        (error: any) => {
            console.log(error);
        }
    );
    // 导出Univer到CSV文件
    // snapshot为Univer快照数据，getBuffer：true不会下载文件，只会返回csv内容，false会直接下载
    // sheetName：因为Univer可能存在多个sheet，csv没有sheet，sheetName有值的情况下只会下载指定sheet名称的数据，不传会下载所有sheet，文件名为${fileName}_${sheet.name}
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
    // 导出Univer到XLSX文件
    // getBuffer：true不会下载文件，只会返回文件的buffer数据，false会直接下载
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
> 案例 [Demo index.html](./src/index.html)展示了详细的用法

### ES 和 Node.js

#### 安装
```shell
npm install @zwight/luckyexcel
```

#### ES导入
```js
import LuckyExcel from '@zwight/luckyexcel'

// 得到xlsx文件后
LuckyExcel.transformExcelToUniver(
    file,
    async (exportJson: any) => {
        // 转换后获取工作表数据
    },
    (error: any) => {
        //如果抛出任何错误，则处理错误
    }
);
```

## 开发

### 环境
[Node.js](https://nodejs.org/en/) Version >= 6 

### 安装
```
npm install -g gulp-cli
npm install
```
### 开发
```
npm run dev
```
### 打包
```
npm run build
```

项目中使用了第三方插件：[JSZip](https://github.com/Stuk/jszip)，感谢！

## 贡献者和感谢
- [@dream-num](https://github.com/dream-num)

## 版权信息
[MIT](http://opensource.org/licenses/MIT)

Copyright (c) 2020-present, zwight
