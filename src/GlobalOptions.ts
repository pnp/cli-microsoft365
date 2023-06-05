export enum Output {
  text,
  json,
  csv,
  md
}

export default interface GlobalOptions {
  query?: string;
  output?: string;
  debug?: boolean;
  verbose?: boolean;
  interactive?: boolean;
  // allow command-specific options. Required for tests to avoid casting to 'any'
  [arg: string]: any;
}