import { SearchResultProperty } from "./SearchResultProperty";
import { QueryResult } from "./QueryResult";

export interface SearchResult {
  ElapsedTime: number;
  PrimaryQueryResult: QueryResult;
  Properties: SearchResultProperty[];
  SecondaryQueryResults: QueryResult[];
  SpellingSuggestion: string;
  TriggeredRules: string[];
}