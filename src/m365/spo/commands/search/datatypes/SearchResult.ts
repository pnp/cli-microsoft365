import { QueryResult } from "./QueryResult.js";
import { SearchResultProperty } from "./SearchResultProperty.js";

export interface SearchResult {
  ElapsedTime: number;
  PrimaryQueryResult: QueryResult;
  Properties: SearchResultProperty[];
  SecondaryQueryResults: QueryResult[];
  SpellingSuggestion: string;
  TriggeredRules: string[];
}