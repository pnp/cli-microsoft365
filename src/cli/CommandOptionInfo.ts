export interface CommandOptionInfo {
  autocomplete?: string[];
  long?: string;
  name: string;
  required: boolean;
  short?: string;
  // undefined as a temporary solution to avoid breaking changes
  // eventually it should be required
  type?: 'string' | 'boolean' | 'number';
}