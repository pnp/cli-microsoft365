export enum Output {
  text,
  json
}

export default interface GlobalOptions {
  query?: string;
  output?: string;
  debug?: boolean;
  verbose?: boolean;
}