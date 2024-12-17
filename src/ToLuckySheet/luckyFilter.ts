/**:
* @description: 返回的filter属性是univer格式
* @author: Created by zwight on 2024-09-23 17:50:01
*/
import { str2num } from "../common/method";
import { ReadXml } from "./ReadXml";
import { handleRanges, IRange } from "./style";
export interface IFilters {
    blank?: boolean;
    filters?: Array<string>;
}
export interface ICustomFilter {
    val: string | number;
    operator?: CustomFilterOperator;
}
export declare enum CustomFilterOperator {
    EQUAL = "equal",
    GREATER_THAN = "greaterThan",
    GREATER_THAN_OR_EQUAL = "greaterThanOrEqual",
    LESS_THAN = "lessThan",
    LESS_THAN_OR_EQUAL = "lessThanOrEqual",
    NOT_EQUALS = "notEqual"
}
export interface ICustomFilters {
    and?: 1;
    customFilters: [ICustomFilter] | [ICustomFilter, ICustomFilter];
}
export interface IFilterColumn {
    colId: number;
    filters?: IFilters;
    customFilters?: ICustomFilters;
}
export interface LuckyFilterFormat {
    ref: IRange,
    filterColumns: IFilterColumn[],
    cachedFilteredOut: number[],
}
export class LuckFilter implements LuckyFilterFormat {
    ref: IRange;
    filterColumns: IFilterColumn[];
    cachedFilteredOut: number[];
    constructor(readXml: ReadXml, sheetFile: string) {
        const autoFilter = readXml.getElementsByTagName('autoFilter', sheetFile)[0];
        if (!autoFilter) return;

        this.ref = handleRanges(autoFilter.attributeList.ref)?.[0];

        const filterColumn = autoFilter.getInnerElements('filterColumn');
        this.filterColumns = filterColumn.map(d => {
            const filters = d.getInnerElements('filters')?.[0];
            const filter = filters.getInnerElements('filter')
            return {
                colId: str2num(d.attributeList.colId) as number,
                filters: {
                    blank: filters.attributeList?.blank === '1',
                    filters: filter.map(d => d.attributeList.val)
                },
            }
        })
    }
}