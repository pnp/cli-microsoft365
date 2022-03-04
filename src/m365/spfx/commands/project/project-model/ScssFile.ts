import * as fs from 'fs';

export class ScssFile {
  private _source: string | undefined;
  public get source(): string | undefined {
    if (!this._source) {
      try {
        this._source = fs.readFileSync(this.path, 'utf-8');
      }
      catch { }
    }

    return this._source;
  }

  constructor(public path: string) {
  }
}