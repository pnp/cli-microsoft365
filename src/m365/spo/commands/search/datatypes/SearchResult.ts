import { QueryResult } from "./QueryResult";
import { SearchResultProperty } from "./SearchResultProperty";

export interface SearchResult {
  ElapsedTime: number;
  PrimaryQueryResult: QueryResult;
  Properties: SearchResultProperty[];
  SecondaryQueryResults: QueryResult[];
  SpellingSuggestion: string;
  TriggeredRules: string[];
}