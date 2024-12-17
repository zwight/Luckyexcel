export class CSV {
    csvList: { [key: string]: Array<Array<any>> } = {};
    csvContent: { [key: string]: string } = {};
    constructor(snapshot: any) {
        this.init(snapshot);
    }

    init(snapshot: any) {
        if (!snapshot) {
            return;
        }
        const { sheetOrder, sheets } = snapshot;
        const data: any = {}
        sheetOrder.forEach((d: string) => {
            const sheet = sheets[d];
            if (!sheet) return;
            const { cellData, name } = sheet;
            const list: Array<Array<any>> = []
            for (const key in cellData) {
                const rows = cellData[key];
                for (const key in rows) {
                    const row = Number(key);
                    const col = Number(key);
                    if (!list[row]) list[row] = [];
                    list[row][col] = rows[key]?.v
                }
            }
            data[name] = list;
        })
        this.csvList = data;
        this.handleCsvContent();
    }

    handleCsvContent() {
        const data: any = {}
        for (const key in this.csvList) {
            const csv = this.csvList[key];
            let csvContent = "data:text/csv;charset=utf-8,";
            
            // 拼接csv数据
            csv.forEach(row => {
                csvContent += row.join(",") + "\r\n";
            });
            data[key] = csvContent
        }
        this.csvContent = data;
    }
}