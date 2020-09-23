export enum Output {
  text,
  json
}

export default interface GlobalOptions {
  query?: string;
  output?: string;
  debug?: boolean;
  verbose?: boolean;
  // allow command-specific options. Required for tests to avoid casting to 'any'
  [arg: string]: any;
}