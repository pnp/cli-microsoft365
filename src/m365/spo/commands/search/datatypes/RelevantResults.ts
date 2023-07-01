import { ResultTable } from "./ResultTable.js";
import { SearchResultProperty } from "./SearchResultProperty.js";

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