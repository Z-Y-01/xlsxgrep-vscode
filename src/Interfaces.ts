export interface ISearchData {
    targetVal: string;
    isWholeMatch: boolean;
    isCaseMatch: boolean;
    bOnlyActiveFiles: boolean;
    filePattern?: string; // 搜索的文件名关键字 (部分匹配)
    TargetRegexExp?: RegExp;
}

export interface ISearchResult {
    path: string;
    fileName: string;
    sheet: string;
    row: number;
    col: string;
    cellContent: string;
    rowContent: string;
}
