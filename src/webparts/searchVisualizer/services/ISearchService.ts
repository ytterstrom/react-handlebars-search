export interface ISearchResults {
    PrimaryQueryResult: IPrimaryQueryResult;
}

export interface IPrimaryQueryResult {
    RelevantResults: IRelevantResults;
}

export interface IRelevantResults {
    Table: ITable;
    TotalRows: number;
    TotalRowsIncludingDuplicates: number;
}
export interface IRefinementResult {
    FilterName: string;
    Values: IRefinementValue[];
}

export interface IRefinementValue {
    RefinementCount: number;
    RefinementName: string;
    RefinementToken: string;
    RefinementValue: string;
}

export interface IRefinementFilter {
    FilterName: string;
    Value: IRefinementValue;
}
export interface ITable {
    Rows: Array<ICells>;
}

export interface ICells {
    Cells: Array<ICellValue>;
}

export interface ICellValue {
    Key: string;
    Value: string;
    ValueType: string;
}

export interface ISearchResponse {
    results: any[];
    totalResults: number;
    totalResultsIncludingDuplicates: number;
    searchUrl: string;
}
