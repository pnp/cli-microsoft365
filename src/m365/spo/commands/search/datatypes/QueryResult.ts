import { RefinementResult } from "./RefinementResult";
import { RelevantResults } from "./RelevantResults";
import { ResultTable } from "./ResultTable";
import { SpecialTermResult } from "./SpecialTermResult";

export interface QueryResult {
  CustomResults: ResultTable[];
  QueryId: string;
  QueryRuleId: string;
  RefinementResults: RefinementResult[] | null;
  RelevantResults: RelevantResults;
  SpecialTermResults: SpecialTermResult[] | null;
}