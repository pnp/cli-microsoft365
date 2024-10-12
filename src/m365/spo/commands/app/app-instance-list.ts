import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import { odata } from '../../../../utils/odata.js';
import { validation } from '../../../../utils/validation.js';
import commands from '../../commands.js';
import { SpoAppBaseCommand } from './SpoAppBaseCommand.js';


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

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving installed apps in site at ${args.options.siteUrl}...`);
    }

    try {
      const apps = await odata.getAllItems<any>(`${args.options.siteUrl}/_api/web/AppTiles?$filter=AppType eq 3`);
      await logger.log(apps);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoAppInStanceListCommand();