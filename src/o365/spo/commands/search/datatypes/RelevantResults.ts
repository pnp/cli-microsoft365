import { SearchResultProperty } from "./SearchResultProperty";
import { ResultTable } from "./ResultTable";

export interface RelevantResults {
  RelevantResults:string;
  ItemTemplateId:string;
  Properties:SearchResultProperty[];
  ResultTitle:string;
  ResultTitleUrl:string;
  RowCount:number;
  Table:ResultTable;
  TotalRows:number;
  TotalRowsIncludingDuplicates:number;
}