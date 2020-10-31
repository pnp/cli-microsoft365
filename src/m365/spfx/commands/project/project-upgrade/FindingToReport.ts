export interface FindingToReport {
  description: string;
  id: string;
  file: string;
  position?: {
    character: number;
    line: number;
  };
  resolution: string;
  resolutionType: string;
  severity: string;
  title: string;
}