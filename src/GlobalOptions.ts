export enum Output {
  text,
  json
}

export default interface GlobalOptions {
  output?: string;
  debug?: boolean;
  verbose?: boolean;
}