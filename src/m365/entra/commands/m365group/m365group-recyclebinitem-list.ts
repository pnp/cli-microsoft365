import { DirectoryObject } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import { formatting } from '../../../../utils/formatting.js';
import { odata } from '../../../../utils/odata.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  groupName?: string;
  groupMailNickname?: string;
}

class EntraM365GroupRecycleBinItemListCommand extends GraphCommand {
  public get name(): string {
    return commands.M365GROUP_RECYCLEBINITEM_LIST;
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
        groupName: typeof args.options.groupName !== 'undefined',
        groupMailNickname: typeof args.options.groupMailNickname !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-d, --groupName [groupName]'
      },
      {
        option: '-m, --groupMailNickname [groupMailNickname]'
      }
    );
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'displayName', 'mailNickname'];
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const filter: string = `?$filter=groupTypes/any(c:c+eq+'Unified')`;
      const displayNameFilter: string = args.options.groupName ? ` and startswith(DisplayName,'${formatting.encodeQueryParameter(args.options.groupName).replace(/'/g, `''`)}')` : '';
      const mailNicknameFilter: string = args.options.groupMailNickname ? ` and startswith(MailNickname,'${formatting.encodeQueryParameter(args.options.groupMailNickname).replace(/'/g, `''`)}')` : '';
      const topCount: string = '&$top=100';
      const endpoint: string = `${this.resource}/v1.0/directory/deletedItems/Microsoft.Graph.Group${filter}${displayNameFilter}${mailNicknameFilter}${topCount}`;

      const recycleBinItems = await odata.getAllItems<DirectoryObject>(endpoint);
      await logger.log(recycleBinItems);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new EntraM365GroupRecycleBinItemListCommand();