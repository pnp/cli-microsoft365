export interface Occurrence {
  file: string;
  position?: {
    character: number;
    line: number;
  };
  resolution: string;
}