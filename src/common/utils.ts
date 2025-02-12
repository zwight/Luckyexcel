import Papa from 'papaparse';

export enum CHARSET_TYPE {
    UTF8 = 'UTF-8',
    GBK = 'GBK',
    CP936 = 'CP936',
    ISO8859 = 'ISO-8859',
}
export const getDataByFile = ({ file, charset = CHARSET_TYPE.UTF8 }: any): Promise<string> => {
    return new Promise((resolve, reject) => {
        if (!(file instanceof File)) resolve('');
        const fileReader = new FileReader();
        fileReader.onload = (e) => {
            try {
                const { result } = e.target as any;
                // 老的逻辑，.txt、.log、.csv 直接用 readAsText 读取后返回即可
                // 新的 xlsx、xls 用 readAsBinaryString 读取处理数据
                resolve(result);
            } catch (err) {
                reject(err);
            }
        };
        // 二进制文件通过 readAsBinaryString 读取
        fileReader.readAsText(file, charset);
    });
};

export const formatSheetData = (sheetData: string, file: File) => {
    const splitData = sheetData.replace(/\r\n/g, '\n').replace(/\r/g, '\n').split('\n');
    let arr = [];
    if (file.name.endsWith('.csv') || file.type === 'text/csv') {
        const csvData: Papa.ParseResult<string[]> = Papa.parse(sheetData, {
            delimiter: ',',
            skipEmptyLines: true,
        });
        arr = csvData.data;
    } else {
        for (let i = 0; i < splitData.length; i++) {
            const str = splitData[i].replace(/\r/, ''); // 清除无用\r字符
            if (str) {
                arr.push(str.split(','));
            }
        }
    }

    return arr;
};