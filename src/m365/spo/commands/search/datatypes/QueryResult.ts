import { RefinementResult } from "./RefinementResult.js";
import { RelevantResults } from "./RelevantResults.js";
import { ResultTable } from "./ResultTable.js";
import { SpecialTermResult } from "./SpecialTermResult.js";

export interface QueryResult {
  CustomResults: ResultTable[];
  QueryId: string;
  QueryRuleId: string;
  RefinementResults: RefinementResult[] | null;
  RelevantResults: RelevantResults;
  SpecialTermResults: SpecialTermResult[] | null;
}