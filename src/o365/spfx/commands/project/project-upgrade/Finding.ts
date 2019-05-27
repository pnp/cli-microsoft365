import { Occurrence } from "./";

export interface Finding {
  description: string;
  id: string;
  occurrences: Occurrence[];
  resolutionType: string;
  severity: string;
  supersedes: string[];
  title: string;
}