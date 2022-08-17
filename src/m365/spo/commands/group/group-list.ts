import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { validation } from '../../../../utils';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';
import { GroupPropertiesCollection } from "./GroupPropertiesCollection";

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
}

class SpoGroupListCommand extends SpoCommand {
  public get name(): string {
    return commands.GROUP_LIST;
  }

  public get description(): string {
    return 'Lists all the groups within specific web';
  }

  public defaultProperties(): string[] | undefined {
    return ['Id', 'Title', 'LoginName', 'IsHiddenInUI', 'PrincipalType'];
  }

  constructor() {
    super();

    this.#initOptions();
    this.#initValidators();
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --webUrl <webUrl>'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => validation.isValidSharePointUrl(args.options.webUrl)
    );
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    if (this.verbose) {
      logger.logToStderr(`Retrieving list of groups for specified web at ${args.options.webUrl}...`);
    }

    const requestUrl = `${args.options.webUrl}/_api/web/sitegroups`;

    const requestOptions: any = {
      url: requestUrl,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    request
      .get<GroupPropertiesCollection>(requestOptions)
      .then((groupProperties: GroupPropertiesCollection): void => {
        logger.log(groupProperties.value);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }
}

module.exports = new SpoGroupListCommand();