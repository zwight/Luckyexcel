/**:
* @description: 返回的filter属性是univer格式
* @author: Created by zwight on 2024-09-23 17:50:01
*/
import { BooleanNumber } from "@univerjs/core";
import { str2num } from "../common/method";
import { ReadXml } from "./ReadXml";
import { handleRanges, IRange } from "./style";
import { replaceSpecialWrap } from "../common/utils";
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

        this.filterColumns = filterColumn?.map(d => {
            const customFilters = d.getInnerElements('customFilters')?.[0];
            const filters = d.getInnerElements('filters')?.[0];
            let colId = str2num(d.attributeList.colId) as number;

            colId = this.ref ? (colId + this.ref.startRow) : colId;

            let customFiltersVal = undefined
            if (customFilters) {
                const customFilter = customFilters.getInnerElements('customFilter');
                customFiltersVal = {
                    and: (customFilters.attributeList?.and === '1' ? 1 : undefined) as BooleanNumber.TRUE | undefined,
                    customFilters: customFilter.map(f => ({
                        val: replaceSpecialWrap(f.attributeList.val),
                        operator: f.attributeList?.operator as CustomFilterOperator,
                    })) as [ICustomFilter] | [ICustomFilter, ICustomFilter],
                }
                
            }
            let filtersVal = undefined;
            if (filters) {
                const filter = filters.getInnerElements('filter');
                filtersVal = {
                    blank: filters.attributeList?.blank === '1',
                    filters: filter.map(d => replaceSpecialWrap(d.attributeList.val))
                }
            }
            return {
                colId,
                filters: filtersVal,
                customFilters: customFiltersVal,
            }
        })
    }
}