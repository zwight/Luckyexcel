import { generateRandomId, getPxByEMUs, getRelationShip, getcellrange, hexToRgbArray, numberToABC } from "../common/method"
import { LuckyChart, LuckyChartImageBase } from "./LuckyBase"
import { Element, IStyleCollections, ReadXml } from "./ReadXml"
import { ChartTypeBits, LabelContentType } from "../common/constant"
import { ILuckyChartContext, ILuckyChartStyle, ILuckyChartStyleAxis, ILuckyChartStyleBase, ILuckyChartStyleLegend } from "./ILuck"

interface ChartGroup {
    chartEle: Element;
    chartType: ChartTypeBits;
}
const getChartType = (readXml: ReadXml, chartFile: string) => {
    const barChart = readXml.getElementsByTagName('c:barChart', chartFile)[0];
    const lineChart = readXml.getElementsByTagName('c:lineChart', chartFile)[0];
    const pieChart = readXml.getElementsByTagName('c:pieChart', chartFile)[0];
    const doughnutChart = readXml.getElementsByTagName('c:doughnutChart', chartFile)[0];
    const areaChart = readXml.getElementsByTagName('c:areaChart', chartFile)[0];
    const radarChart = readXml.getElementsByTagName('c:radarChart', chartFile)[0];
    const scatterChart = readXml.getElementsByTagName('c:scatterChart', chartFile)[0];

    let chartGroup: ChartGroup[] = []
    let chartEle: Element | ChartGroup[] = barChart;
    let chartType: ChartTypeBits = ChartTypeBits.Column;

    if (barChart) {
        const barDirVal = barChart.getInnerElements('c:barDir')?.[0]?.get('val');
        const groupingVal = barChart.getInnerElements('c:grouping')?.[0]?.get('val');
        if (barDirVal === 'col') {
            chartType = ChartTypeBits.Column;
            if (groupingVal === 'stacked') {
                chartType = ChartTypeBits.ColumnStacked;
            } else if (groupingVal === 'percentStacked') {
                chartType = ChartTypeBits.ColumnPercentStacked;
            }
        } else if (barDirVal === 'bar') {
            chartType = ChartTypeBits.Bar;
            if (groupingVal === 'stacked') {
                chartType = ChartTypeBits.BarStacked;
            } else if (groupingVal === 'percentStacked') {
                chartType = ChartTypeBits.BarPercentStacked;
            }
        }
        chartEle = barChart;

        chartGroup.push({ chartEle, chartType })
    }
    if (lineChart) {
        chartEle = lineChart;
        chartType = ChartTypeBits.Line;
        chartGroup.push({ chartEle, chartType })
    }
    if (pieChart) {
        chartEle = pieChart;
        chartType = ChartTypeBits.Pie;
        chartGroup.push({ chartEle, chartType })
    } 
    if (doughnutChart) {
        chartEle = doughnutChart;
        chartType = ChartTypeBits.Doughnut;
        chartGroup.push({ chartEle, chartType })
    } 
    if (areaChart) {
        chartEle = areaChart;
        chartType = ChartTypeBits.Area;
        const groupingVal = areaChart.getInnerElements('c:grouping')?.[0]?.get('val');
        if (groupingVal === 'stacked') {
            chartType = ChartTypeBits.AreaStacked;
        } else if (groupingVal === 'percentStacked') {
            chartType = ChartTypeBits.AreaPercentStacked;
        }
        chartGroup.push({ chartEle, chartType })
    } 
    if (radarChart) {
        chartEle = radarChart;
        chartType = ChartTypeBits.Radar;
        chartGroup.push({ chartEle, chartType })
    } 
    if (scatterChart) {
        chartEle = scatterChart;
        chartType = ChartTypeBits.Scatter;
        chartGroup.push({ chartEle, chartType })
    }

    if (chartGroup.length > 1) {
        chartType = ChartTypeBits.Combination;
        chartEle = chartGroup
    }

    return {
        chartEle,
        chartType
    }
}
export class LuckyChartImage extends LuckyChartImageBase {
    id: string
    constructor(id: string, xdr_xfrm: Element, range: string, chartType: ChartTypeBits) {
        super()
        this.id = id;
        this.type = 'chart'
        this.data = {
            chartType,
            range,
            border: '#979DAC',
            background: 'rgba(0,0,0,0)',
            isRowDirection: true,
        }
        let x_n = 0, y_n = 0;
        let cx_n = 0, cy_n = 0;
        const off = xdr_xfrm.getInnerElements('a:off')[0];
        const ext = xdr_xfrm.getInnerElements('a:ext')[0];
        cx_n = getPxByEMUs(parseInt(ext.get('cx') as string)), cy_n = getPxByEMUs(parseInt(ext.get('cy') as string));
        x_n = getPxByEMUs(parseInt(off.get('x') as string)), y_n = getPxByEMUs(parseInt(off.get('y') as string));
        this.transform = {
            width: cx_n,
            height: cy_n,
            top: y_n,
            left: x_n,
        }

    }
}
export class ChartImageGroup {
    chart: LuckyChart;
    image: LuckyChartImage;

    readXml: ReadXml;
    constructor({
        graphicFrame,
        readXml,
        drawingRelsFile,
        styles
    }: {
        graphicFrame: Element,
        readXml: ReadXml,
        drawingRelsFile: string,
        styles: IStyleCollections
    }) {
        this.readXml = readXml;

        const xdr_xfrm = graphicFrame.getInnerElements('xdr:xfrm')[0];

        const cChart = readXml.getElementsByTagName('a:graphic/a:graphicData/c:chart', graphicFrame.value, false)[0];
        const chartRid = cChart.get('r:id') as string;
        const chartFile = getRelationShip({
            rid: chartRid,
            fileName: drawingRelsFile,
            readXml
        });

        const { chartEle, chartType } = getChartType(readXml, chartFile);
        
        const range = this.getChartRange(chartEle);
        
        const id = generateRandomId()

        this.image = new LuckyChartImage(id, xdr_xfrm, range, chartType);

        const chart = new Chart({
            id,
            range,
            chartType,
            chartFile,
            readXml,
            image: this.image,
            styles,
        });
        this.chart = chart.model
    }

    private getChartRange = (chartEl: Element | ChartGroup[]) => {
        let maxColumn = 0, maxRow = 0, minColumn = 0, minRow = 0;
        if (Array.isArray(chartEl)) {
            const rangeNumArray = chartEl.map(d => {
                return this.getChartRef(d.chartEle.value);
            })
            rangeNumArray.forEach((d, index) => {
                if (index === 0) {
                    maxColumn = d.maxColumn;
                    maxRow = d.maxRow;
                    minColumn = d.minColumn;
                    minRow = d.minRow;
                } else {
                    maxColumn = Math.max(maxColumn, d.maxColumn);
                    maxRow = Math.max(maxRow, d.maxRow);
                    minColumn = Math.min(minColumn, d.minColumn);
                    minRow = Math.min(minRow, d.minRow);
                }
            })
        } else {
            const rangeNum = this.getChartRef(chartEl.value);
            maxColumn = rangeNum.maxColumn;
            maxRow = rangeNum.maxRow;
            minColumn = rangeNum.minColumn;
            minRow = rangeNum.minRow;
        }
        
        const maxRef = numberToABC(maxColumn) + (maxRow + 1);
        const minRef = numberToABC(minColumn) + (minRow + 1);
        return minRef + ':' + maxRef;
    }

    private getChartRef = (chart: string) => {
        const catNumRef = this.readXml.getElementsByTagName('c:ser/c:cat/c:numRef/c:f', chart, false)?.[0];
        const catStrRef = this.readXml.getElementsByTagName('c:ser/c:cat/c:strRef/c:f', chart, false)?.[0];
        const cXvalStrRef = this.readXml.getElementsByTagName('c:ser/c:xVal/c:strRef/c:f', chart, false)?.[0];
        const cXvalNumRef = this.readXml.getElementsByTagName('c:ser/c:xVal/c:numRef/c:f', chart, false)?.[0];

        const catRef = catNumRef || catStrRef || cXvalStrRef || cXvalNumRef;

        const strRef = this.readXml.getElementsByTagName('c:ser/c:tx/c:strRef/c:f', chart, false);
        const x = catRef, y = strRef[strRef.length - 1]

        const xRange = getcellrange(x.value), yRange = getcellrange(y.value);

        const column = [...xRange.column, ...yRange.column];
        const row = [...xRange.row, ...yRange.row];
        const maxColumn = Math.max(...column), maxRow = Math.max(...row);
        const minColumn = Math.min(...column), minRow = Math.min(...row);
        return {
            maxColumn,
            maxRow,
            minColumn,
            minRow
        }
    }
}
class Chart extends LuckyChart {
    chartFile: string;
    readXml: ReadXml;
    styles: IStyleCollections;
    constructor(args: {
        id: string,
        range: string;
        chartType: ChartTypeBits,
        chartFile: string,
        readXml: ReadXml,
        image: LuckyChartImage,
        styles: IStyleCollections
    }) {
        super();
        const { id, range, chartType, chartFile, readXml, image, styles } = args;

        this.styles = styles;
        this.range = range;
        this.chartType = chartType;
        this.readXml = readXml;
        this.chartFile = chartFile;
        this.isRowDirection = true;
        this.id = id;

        this.context = this.getContext();
        this.style = this.getStyle(image);
    }

    private getStyle = (image: LuckyChartImage): ILuckyChartStyle => {
        const bodySpPr = this.readXml.getElementsByTagNameLink('c:spPr', this.chartFile)[0];
        
        const fill = bodySpPr.getInnerElements('a:solidFill')[0];
        const ln = bodySpPr.getInnerElements('a:ln')[0];
        const backgroundColor = fill ? this.getColor(fill.getInnerElements('a:schemeClr')[0]) : undefined;
        const borderColor = ln ? this.getColor(ln.getInnerElements('a:schemeClr')[0]) : undefined;

        const allTitle = this.getAllTitle();
        const autoTitleDeleted = this.readXml.getElementsByTagNameLink('c:chart/c:autoTitleDeleted', this.chartFile)[0]

        const plotArea = this.readXml.getElementsByTagNameLink('c:chart/c:plotArea', this.chartFile)[0];
        const xAxis = this.getAxis(plotArea?.getInnerElements('c:catAx')?.[0]);
        const yAxis = this.getAxis(plotArea?.getInnerElements('c:valAx')?.[0]);

        const cLegend = this.readXml.getElementsByTagNameLink('c:chart/c:legend', this.chartFile)[0]
        const legend = this.getLegend(cLegend);
        return {
            titles: {
                ...allTitle,
                titlePosition: autoTitleDeleted?.get('val') === '1' ? 'hide' : 'top',
            },
            runtime: {},
            width: image.transform.width,
            height: image.transform.height,
            backgroundColor,
            borderColor,
            xAxis,
            yAxis,
            legend,
            ...this.getChartSeries()
        }
    }

    private getChartSeries = () => {
        const { chartEle, chartType } = getChartType(this.readXml, this.chartFile)
        if (!chartEle) return {};
        if (Array.isArray(chartEle)) {
            const seriesStyleArray = chartEle.map(d => {
                return this.getChartSeriesBase(d.chartEle, d.chartType, true);
            })
            const seriesStyleMap = seriesStyleArray.reduce((pre, cur) => {
                return Object.assign(pre, cur)
            }, {})
            return {
                seriesStyleMap
            }
        }
        const allDlbls = chartEle.getInnerElementsTagLink('c:dLbls')[0];
        const show = allDlbls.getInnerElements('c:showVal')[0];
        const allSeriesStyle = {
            label: {
                visible: show ? show.get('val') === '1' : false
            }
        }
        const seriesStyleMap = this.getChartSeriesBase(chartEle, chartType);

        return {
            allSeriesStyle,
            seriesStyleMap
        }
    }

    private getChartSeriesBase = (chartEle: Element, chartType: ChartTypeBits, isGroup?: boolean) => {
        const cSer = chartEle.getInnerElements('c:ser');
        const seriesStyleMap: any = {};
        cSer.forEach((element) => {
            const idx = element.getInnerElements('c:idx')[0];
            const idxVal = parseInt(idx.get('val') as string) + 1;

            const spPr = element.getInnerElements('c:spPr')[0];

            const fill = spPr?.getInnerElementsTagLink('a:solidFill')?.[0];
            const ln = spPr?.getInnerElements('a:ln')?.[0];
            const barColor = this.getColor(fill?.getInnerElements('a:schemeClr')?.[0]);
            const fillOpacity = fill?.getInnerElements('a:alpha')?.[0]?.get('val') as string;

            const border = ln ? this.getLine(ln) : undefined

            const dLbls = element.getInnerElements('c:dLbls')[0];
            const showValue = dLbls.getInnerElements('c:showVal')?.[0]?.get('val') as string;
            const showCatName = dLbls.getInnerElements('c:showCatName')?.[0]?.get('val') as string;
            const showSerName = dLbls.getInnerElements('c:showSerName')?.[0]?.get('val') as string;
            const showPercent = dLbls.getInnerElements('c:showPercent')?.[0]?.get('val') as string;
            const labelStyle = this.getBaseStyle(dLbls.getInnerElements('a:defRPr')?.[0]);

            const contentType = 0 | 
            (showValue === '1' ? LabelContentType.Value : 0) | 
            (showCatName === '1' ? LabelContentType.CategoryName : 0) | 
            (showSerName === '1' ? LabelContentType.SeriesName : 0) | 
            (showPercent === '1' ? LabelContentType.Percentage : 0);

            const base = {
                border,
                label: {
                    visible: !!contentType,
                    contentType,
                    ...labelStyle,
                },
                color: barColor,
                fillOpacity: fillOpacity ? parseInt(fillOpacity) / 100000 : 1,
                chartType: isGroup ? chartType : undefined,
            }
            let serConf = this.getExtraSerise(base, element, chartType)

            seriesStyleMap[idxVal] = serConf
        });

        return seriesStyleMap;
    }

    private getExtraSerise = (baseObj: any, series: Element, chartType: ChartTypeBits) => {
        if (chartType === ChartTypeBits.Line) {
            if (baseObj.border?.color) {
                baseObj.color = baseObj.border.color
            }
            const marker = series.getInnerElements('c:marker')[0];
            const symbol = marker.getInnerElements('c:symbol')?.[0]?.get('val') as string;
            const size = marker.getInnerElements('c:size')?.[0]?.get('val') as string;
            const schemeClr = marker.getInnerElementsTagLink('c:spPr/a:solidFill/a:schemeClr')?.[0];
            const color = this.getColor(schemeClr);
            baseObj['point'] = {
                color,
                size: size ? parseInt(size) : undefined,
                shape: symbol === 'none' && baseObj.label.visible ? 'circle' : symbol
            }
        }
        return baseObj
    }

    private getLegend = (legend: Element): ILuckyChartStyleLegend => {
        if (!legend) return {
            position: 'hide'
        };
        const positionMap: { [key: string]: 'top' | 'bottom' | 'left' | 'right' | 'hide' } = {
            t: 'top',
            b: 'bottom',
            l: 'left',
            r: 'right',
        }
        const position = legend.getInnerElements('c:legendPos')[0]?.get('val') as string;
        const txPr = legend.getInnerElements('c:txPr')[0];
        const pPr = txPr.getInnerElements('a:pPr')[0];
        const labelStyle = this.getBaseStyle(pPr.getInnerElements('a:defRPr')[0]);

        return {
            position: position? positionMap[position] : 'bottom',
            label: labelStyle
        }
    }

    private getAxis = (axis: Element): ILuckyChartStyleAxis => {
        if (!axis) return undefined;
        const visible = !axis.getInnerElementsTagLink('c:spPr/a:ln/a:noFill')?.length
        // const visible = axis.getInnerElements('c:delete')[0]?.get('val') === '0';
        const scaling = axis.getInnerElements('c:scaling')[0];
        const reverse = scaling?.getInnerElements('c:orientation')[0]?.get('val') === 'maxMin';
        const max = scaling?.getInnerElements('c:max')?.[0]?.get('val') as string;
        const min = scaling?.getInnerElements('c:min')?.[0]?.get('val') as string;
        const titleAlignMap: { [key: string]: 'start' | 'center' | 'end' } = {
            l: 'start',
            r: 'end',
            ctr: 'center'
        }
        const axisTitleAlign = axis.getInnerElements('c:lblAlgn')?.[0]?.get('val') as string;
        const txPr = axis.getInnerElements('c:txPr')[0];
        const pPr = txPr.getInnerElements('a:pPr')[0];
        const rotate = txPr?.getInnerElements('a:bodyPr')?.[0]?.get('rot') as string;
        const labelStyle = this.getBaseStyle(pPr.getInnerElements('a:defRPr')[0]);

        const majorGridlines = axis.getInnerElements('c:majorGridlines')?.[0];
        const ln = majorGridlines?.getInnerElements('a:ln')[0];
        const gridLineWidth = ln?.get('w') ? getPxByEMUs(parseInt(ln.get('w') as string)) : undefined

        const majorTickMark = axis.getInnerElements('c:majorTickMark')?.[0];
        const gridLineColor = this.getColor(ln?.getInnerElements('a:schemeClr')?.[0]);

        return {
            lineVisible: visible,
            reverse,
            max: max ? parseInt(max) : undefined,
            min: min ? parseInt(min) : undefined,
            label: {
                axisTitleAlign: axisTitleAlign ? titleAlignMap[axisTitleAlign] : 'center',
                rotate: rotate && parseInt(rotate) > 0 ? parseInt(rotate) / 60000 : 0,
                ...labelStyle
            },
            gridLine: {
                visible: !!majorGridlines?.value,
                width: gridLineWidth,
                color: gridLineColor,
            },
            tick: {
                visible: majorTickMark?.get('val') !== 'none',
                position: majorTickMark?.get('val') === 'out' ? 'outside' : 'inside',
                lineWidth: gridLineWidth
            }
        }
    }

    private getLine = (ln: Element) => {
        const borderWidth = ln?.get('w') ? getPxByEMUs(parseInt(ln.get('w') as string)) : 0
        const borderColor = this.getColor(ln?.getInnerElements('a:schemeClr')?.[0]);
        const borderType = ln.getInnerElements('a:prstDash')?.[0]?.get('val') as string;
        const borderMap: any = {
            solid: 'solid',
            dot: 'dotted',
            dash: 'dashed',
            sysDot: 'dotted',
            sysDash: 'dashed',
            lgDash: 'dashed',
        }
        
        return {
            dashType: borderType && borderMap[borderType] ? borderMap[borderType] : 'solid',
            color: borderColor,
            width: borderWidth
        }
    }
    private getAllTitle = () => {
        const mainTitle = this.readXml.getElementsByTagNameLink('c:chart/c:title', this.chartFile)[0];
        const catTitle = this.readXml.getElementsByTagNameLink('c:chart/c:plotArea/c:catAx/c:title', this.chartFile)[0];
        const valTitle = this.readXml.getElementsByTagNameLink('c:chart/c:plotArea/c:valAx/c:title', this.chartFile)[0];
        const title = this.getTitle(mainTitle);
        const xTitle = this.getTitle(catTitle);
        const yTitle = this.getTitle(valTitle);

        return {
            title,
            xAxisTitle: xTitle,
            yAxisTitle: yTitle,
        }
    }

    /**
     * 标题
     * @returns 
     */
    private getTitle = (title: Element) => {
        if (!title) return;
        const ar = title.getInnerElements('a:r');
        const pPr = title.getInnerElements('a:pPr')[0];
        const titleBase = this.getBaseStyle(pPr.getInnerElements('a:defRPr')[0]);
        const titleStyle = ar?.map(d => {
            return this.getBaseStyle(d.getInnerElements('a:rPr')?.[0]);
        }) || []

        return {
            content: ar?.map(d => d.getInnerElements('a:t')[0].value).join(''),
            ...titleBase,
            ...(titleStyle.length ? titleStyle[0] : {}),
        }
    }

    private getBaseStyle = (style: Element) => {
        if (!style) return {}
        const schemaClr = style.getInnerElements('a:schemeClr')?.[0];
        const solidFill = style.getInnerElements('a:solidFill')?.[0];

        const color = (!schemaClr && solidFill) ? this.getThemColor(solidFill) : this.getColor(schemaClr)
        const obj: ILuckyChartStyleBase = {};

        if (style.get('sz')) {
            obj.fontSize = parseInt(style.get('sz') as string) / 100;
        }
        if (style.get('b')) {
            obj.bold = style.get('b') === '1';
        }
        if (style.get('i')) {
            obj.italic = style.get('i') === '1';
        }
        if (color) {
            obj.color = color;
        }
        return obj
    }

    private getColor = (ele: Element) => {
        if (!ele) return undefined;
        const val = ele.get('val') as string;
        const clrScheme = this.styles['clrScheme'] as Element[];
        const schema = clrScheme.find((item: Element) => {
            const clrs = item.getInnerElements("a:sysClr|a:srgbClr")[0]
            if (item.container.includes(val)) return true;
            if (val === 'tx1') return clrs.get('val') === 'windowText';
            if (val === 'bg1') return clrs.get('val') === 'window';
            return false;
        });
        let color = '#000000';
        if (schema) {
            color = this.getThemColor(schema)
        }
        const mod = ele.getInnerElements('a:lumMod')?.[0]?.get('val') as string;
        const lum = ele.getInnerElements('a:lumOff')?.[0]?.get('val') as string;

        const rgbArray = hexToRgbArray(color).map(d => {
            const sumMod = mod ? d * parseInt(mod) / 100000 : d;
            const sumLum = lum ? Math.round(255 * (parseInt(lum) / 100000)) : 0;
            return Math.round(sumMod + sumLum)
        });

        return `rgb(${rgbArray.join(',')})`
    }

    private getThemColor = (schema: Element) => {
        let color = '#000000'
        const clrs = schema.getInnerElements("a:sysClr|a:srgbClr");
        if(clrs!=null){
            const clr = clrs[0];
            const clrAttrList = clr.attributeList;
            if(clr.container.indexOf("sysClr")>-1){
                if(clrAttrList.lastClr!=null){
                    color = "#" + clrAttrList.lastClr;
                } else if(clrAttrList.val!=null){
                    color = "#" + clrAttrList.val;
                }
            } else if(clr.container.indexOf("srgbClr")>-1){
                color = "#" + clrAttrList.val;
            }
        }
        return color
    }

    private getContext = (): ILuckyChartContext => {
        const cSer = this.readXml.getElementsByTagName('c:ser', this.chartFile);
        const indexs = cSer.map((d) => {
            const idx = d.getInnerElements('c:idx')[0];
            const idxVal = parseInt(idx.get('val') as string) + 1;
            return idxVal
        })
        // console.log(cSer, indexs)
        return {
            categoryIndex: 0,
            seriesIndexes: indexs
        }
    }

    get model(): LuckyChart {
        return {
            id: this.id,
            range: this.range,
            chartType: this.chartType,
            context: this.context,
            style: this.style,
            isRowDirection: this.isRowDirection,
        }
    }
}