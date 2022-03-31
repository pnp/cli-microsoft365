export interface OccurrencePosition {
  character: number;
  line: number;
}

export interface Occurrence {
  file: string;
  position?: OccurrencePosition;
  resolution: string;
}