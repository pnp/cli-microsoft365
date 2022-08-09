import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { validation } from '../../../../utils';
import commands from '../../commands';
import { SpoAppBaseCommand } from './SpoAppBaseCommand';


interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  siteUrl: string;
}

class SpoAppInStanceListCommand extends SpoAppBaseCommand {
  public get name(): string {
    return commands.APP_INSTANCE_LIST;
  }

  public get description(): string {
    return 'Retrieve apps installed in a site';
  }

  public defaultProperties(): string[] | undefined {
    return [`Title`, `AppId`];
  }

  constructor() {
    super();
  
    this.#initOptions();
    this.#initValidators();
  }
  
  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --siteUrl <siteUrl>'
      }
    );
  }
  
  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        const isValidSharePointUrl: boolean | string = validation.isValidSharePointUrl(args.options.siteUrl);
        if (isValidSharePointUrl !== true) {
          return isValidSharePointUrl;
        }
    
        return true;
      }
    );
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    if (this.verbose) {
      logger.logToStderr(`Retrieving installed apps in site at ${args.options.siteUrl}...`);
    }

    const requestOptions: any = {
      url: `${args.options.siteUrl}/_api/web/AppTiles?$filter=AppType eq 3`,
      method: 'GET',
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    request.get(requestOptions)
      .then((apps: any): void => {
        if (apps.value && apps.value.length > 0) {
          logger.log(apps.value);
        }
        else {
          if (this.verbose) {
            logger.logToStderr('No apps found');
          }
        }
        cb();
      }, (rawRes: any): void => this.handleRejectedODataJsonPromise(rawRes, logger, cb));

  }
}

module.exports = new SpoAppInStanceListCommand();