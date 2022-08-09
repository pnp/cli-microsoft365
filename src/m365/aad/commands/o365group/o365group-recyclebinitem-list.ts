import { DirectoryObject } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import { odata } from '../../../../utils';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  displayName?: string;
  mailNickname?: string;
}

class AadO365GroupRecycleBinItemListCommand extends GraphCommand {
  public get name(): string {
    return commands.O365GROUP_RECYCLEBINITEM_LIST;
  }

  public get description(): string {
    return 'Lists Microsoft 365 Groups deleted in the current tenant';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        displayName: typeof args.options.displayName !== 'undefined',
        mailNickname: typeof args.options.mailNickname !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-d, --displayName [displayName]'
      },
      {
        option: '-m, --mailNickname [mailNickname]'
      }
    );
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'displayName', 'mailNickname'];
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const filter: string = `?$filter=groupTypes/any(c:c+eq+'Unified')`;
    const displayNameFilter: string = args.options.displayName ? ` and startswith(DisplayName,'${encodeURIComponent(args.options.displayName).replace(/'/g, `''`)}')` : '';
    const mailNicknameFilter: string = args.options.mailNickname ? ` and startswith(MailNickname,'${encodeURIComponent(args.options.mailNickname).replace(/'/g, `''`)}')` : '';
    const topCount: string = '&$top=100';

    const endpoint: string = `${this.resource}/v1.0/directory/deletedItems/Microsoft.Graph.Group${filter}${displayNameFilter}${mailNicknameFilter}${topCount}`;

    odata
      .getAllItems<DirectoryObject>(endpoint)
      .then((recycleBinItems): void => {
        logger.log(recycleBinItems);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }
}

module.exports = new AadO365GroupRecycleBinItemListCommand();