import auth, { AuthType } from '../../Auth';
import { Logger } from '../../cli';
import Command, { CommandArgs, CommandError } from '../../Command';
const csomDefs = require('../../../csom.json');

export default abstract class SpoCommand extends Command {
  /**
   * Defines list of options that contain URLs in spo commands. CLI will use
   * this list to expand server-relative URLs specified in these options to
   * absolute.
   * If a command requires one of these options to contain a server-relative
   * URL, it should override this method and remove the necessary property from
   * the array before returning it.
   */
  protected getNamesOfOptionsWithUrls(): string[] {
    const namesOfOptionsWithUrls: string[] = [
      'appCatalogUrl',
      'siteUrl',
      'webUrl',
      'origin',
      'url',
      'imageUrl',
      'actionUrl',
      'logoUrl',
      'libraryUrl',
      'thumbnailUrl',
      'targetUrl',
      'newSiteUrl',
      'previewImageUrl',
      'NoAccessRedirectUrl',
      'StartASiteFormUrl',
      'OrgNewsSiteUrl',
      'parentWebUrl',
      'siteLogoUrl'
    ];
    const excludedOptionsWithUrls: string[] | undefined = this.getExcludedOptionsWithUrls();
    if (!excludedOptionsWithUrls) {
      return namesOfOptionsWithUrls;
    }
    else {
      return namesOfOptionsWithUrls.filter(o => excludedOptionsWithUrls.indexOf(o) < 0);
    }
  }

  /**
   * Array of names of options with URLs that should be excluded
   * from processing. To be overriden in commands that require
   * specific options to be a server-relative URL
   */
  protected getExcludedOptionsWithUrls(): string[] | undefined {
    return undefined;
  }

  public async processOptions(options: any): Promise<void> {
    const namesOfOptionsWithUrls: string[] = this.getNamesOfOptionsWithUrls();
    const optionNames = Object.getOwnPropertyNames(options);
    for (const optionName of optionNames) {
      if (namesOfOptionsWithUrls.indexOf(optionName) < 0) {
        continue;
      }

      const optionValue: any = options[optionName];
      if (typeof optionValue !== 'string' ||
        !optionValue.startsWith('/')) {
        continue;
      }

      await auth.restoreAuth();

      if (!auth.service.spoUrl) {
        throw new Error(`SharePoint URL is not available. Set SharePoint URL using the 'm365 spo set' command or use absolute URLs`);
      }

      options[optionName] = auth.service.spoUrl + optionValue;
    }
  }

  protected validateUnknownCsomOptions(options: any, csomObject: string, csomPropertyType: 'get' | 'set'): string | boolean {
    const unknownOptions: any = this.getUnknownOptions(options);
    const optionNames: string[] = Object.getOwnPropertyNames(unknownOptions);
    if (optionNames.length === 0) {
      return true;
    }

    for (let i: number = 0; i < optionNames.length; i++) {
      const optionName: string = optionNames[i];
      const csomOptionType: string = csomDefs[csomObject][csomPropertyType][optionName];

      if (!csomOptionType) {
        return `${optionName} is not a valid ${csomObject} property`;
      }

      if (['Boolean', 'String', 'Int32'].indexOf(csomOptionType) < 0) {
        return `Unknown properties of type ${csomOptionType} are not yet supported`;
      }
    }

    return true;
  }

  public action(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    auth
      .restoreAuth()
      .then((): void => {
        if (auth.service.connected && AuthType[auth.service.authType] === AuthType[AuthType.Secret]) {
          cb(new CommandError(`SharePoint does not support authentication using client ID and secret. Please use a different login type to use SharePoint commands.`));
          return;
        }

        super.action(logger, args, cb);
      }, (error: any): void => {
        cb(new CommandError(error));
      });
  }
}
