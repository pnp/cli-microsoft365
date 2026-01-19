import { DirectoryObject } from '@microsoft/microsoft-graph-types';
import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import { formatting } from '../../../../utils/formatting.js';
import { odata } from '../../../../utils/odata.js';
import { zod } from '../../../../utils/zod.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

const options = globalOptionsZod
  .extend({
    groupName: zod.alias('d', z.string().optional()),
    groupMailNickname: zod.alias('m', z.string().optional())
  })
  .strict();

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class EntraM365GroupRecycleBinItemListCommand extends GraphCommand {
  public get name(): string {
    return commands.M365GROUP_RECYCLEBINITEM_LIST;
  }

  public get description(): string {
    return 'Lists Microsoft 365 Groups deleted in the current tenant';
  }

  public get schema(): z.ZodTypeAny | undefined {
    return options;
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