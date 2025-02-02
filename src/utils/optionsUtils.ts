import { CommandOption } from "../Command";

export const optionsUtils = {
  getUnknownOptions(options: any, defaultOptions: any): any {
    const unknownOptions: any = JSON.parse(JSON.stringify(options));
    // remove minimist catch-all option
    delete unknownOptions._;

    const knownOptions: CommandOption[] = defaultOptions;
    const longOptionRegex: RegExp = /--([^\s]+)/;
    const shortOptionRegex: RegExp = /-([a-z])\b/;
    knownOptions.forEach(o => {
      const longOptionName: string = (longOptionRegex.exec(o.option) as RegExpExecArray)[1];
      delete unknownOptions[longOptionName];

      // short names are optional so we need to check if the current command has
      // one before continuing
      const shortOptionMatch: RegExpExecArray | null = shortOptionRegex.exec(o.option);
      if (shortOptionMatch) {
        const shortOptionName: string = shortOptionMatch[1];
        delete unknownOptions[shortOptionName];
      }
    });

    return unknownOptions;
  },

  addUnknownOptionsToPayload(payload: any, options: any, defaultOptions: any): void {
    const unknownOptions: any = this.getUnknownOptions(options, defaultOptions);
    const unknownOptionsNames: string[] = Object.getOwnPropertyNames(unknownOptions);
    unknownOptionsNames.forEach(o => {
      payload[o] = unknownOptions[o];
    });
  }
};