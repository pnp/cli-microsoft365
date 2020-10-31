export interface CommandOptionInfo {
  autocomplete: string[] | undefined;
  long?: string;
  name: string;
  required: boolean;
  short?: string;
}