import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import { odata } from '../../../../utils/odata.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  siteUrl: z.string()
    .refine(url => validation.isValidSharePointUrl(url) === true, {
      error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
    }).alias('u'),
  state: z.enum(['approved', 'declined', 'pending']).optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

interface AccessRequestItem {
  Id: number;
  Title: string | null;
  RequestDate: string;
  Status: number;
  PermissionLevelRequested: number;
  PermissionType: string | null;
  IsInvitation: boolean;
  Conversation: string | null;
  RequestedObjectUrl: string | null;
  RequestedObjectTitle: string | null;
  RequestedByDisplayName: string | null;
  RequestedForDisplayName: string | null;
  StatusLabel?: string;
}

class SpoSiteAccessRequestListCommand extends SpoCommand {
  private static readonly statusMap: { [key: number]: string } = {
    0: 'pending',
    1: 'approved',
    3: 'declined'
  };

  public get name(): string {
    return commands.SITE_ACCESSREQUEST_LIST;
  }

  public get description(): string {
    return 'Lists all access requests for a specific site';
  }

  public defaultProperties(): string[] | undefined {
    return ['Id', 'RequestDate', 'RequestedForDisplayName', 'PermissionLevelRequested', 'StatusLabel'];
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (this.verbose) {
        await logger.logToStderr(`Listing access requests for site '${args.options.siteUrl}'...`);
      }

      let requestUrl = `${args.options.siteUrl}/_api/web/AccessRequestsList/Items?$select=*,Status`;

      if (args.options.state) {
        const statusValue = this.getStatusValue(args.options.state);
        requestUrl += `&$filter=Status eq ${statusValue}`;
      }

      const items = await odata.getAllItems<AccessRequestItem>(requestUrl)
        .catch((err: any) => {
          if (err?.error?.error?.code?.includes('ResourceNotFoundException')) {
            return [] as AccessRequestItem[];
          }
          throw err;
        });

      const result = items.map(item => ({
        ...item,
        StatusLabel: SpoSiteAccessRequestListCommand.statusMap[item.Status] ?? 'unknown'
      }));

      await logger.log(result);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private getStatusValue(state: string): number {
    switch (state) {
      case 'approved':
        return 1;
      case 'declined':
        return 3;
      case 'pending':
      default:
        return 0;
    }
  }
}

export default new SpoSiteAccessRequestListCommand();
