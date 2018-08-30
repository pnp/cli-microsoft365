export interface Finding {
  description: string;
  file: string;
  id: string;
  position?: {
    character: number;
    line: number;
  };
  resolution: string;
  resolutionType: string;
  severity: string;
  title: string;
}