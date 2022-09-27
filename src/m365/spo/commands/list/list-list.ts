import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { validation } from '../../../../utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { ListInstanceCollection } from "./ListInstanceCollection";

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  executeWithLimitedPermissions?: boolean;
}

class SpoListListCommand extends SpoCommand {
  public get name(): string {
    return commands.LIST_LIST;
  }

  public get description(): string {
    return 'Lists all available list in the specified site';
  }

  public defaultProperties(): string[] | undefined {
    return ['Title', 'Url', 'Id'];
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        executeWithLimitedPermissions: !!args.options.executeWithLimitedPermissions
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '--executeWithLimitedPermissions'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => validation.isValidSharePointUrl(args.options.webUrl)
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      logger.logToStderr(`Retrieving all lists in site at ${args.options.webUrl}...`);
    }

    const suffix = args.options.executeWithLimitedPermissions ? '&$select=RootFolder/ServerRelativeUrl,*' : '';

    const requestOptions: any = {
      url: `${args.options.webUrl}/_api/web/lists?$expand=RootFolder${suffix}`,
      method: 'GET',
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    try {
      const listInstances = await request.get<ListInstanceCollection>(requestOptions);
      listInstances.value.forEach(l => {
        l.Url = l.RootFolder.ServerRelativeUrl;
      });

      logger.log(listInstances.value);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new SpoListListCommand();