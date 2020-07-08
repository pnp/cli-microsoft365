import { SearchResultProperty } from "./SearchResultProperty";
import { ResultTable } from "./ResultTable";

export interface RelevantResults {
  GroupTemplateId: string | null;
  ItemTemplateId: string | null;
  Properties: SearchResultProperty[];
  ResultTitle: string | null;
  ResultTitleUrl: string | null;
  RowCount: number;
  Table: ResultTable;
  TotalRows: number;
  TotalRowsIncludingDuplicates: number;
}