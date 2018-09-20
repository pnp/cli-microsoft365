import { ResultTable } from "./ResultTable";
import { RefinementResult } from "./RefinementResult";
import { SpecialTermResult } from "./SpecialTermResult";
import { RelevantResults } from "./RelevantResults";

export interface QueryResult {
  CustomResults:ResultTable[];
  QueryId:string;
  QueryRuleId:string;
  RefinementResults:RefinementResult[];
  RelevantResults:RelevantResults;
  SpecialTermResults:SpecialTermResult[];
}