import { z } from 'zod';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { formatting } from '../../../../utils/formatting.js';
import { odata } from '../../../../utils/odata.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  webUrl: z.string().alias('u')
    .refine(url => validation.isValidSharePointUrl(url) === true, {
      error: e => `'${e.input}' is not a valid SharePoint site URL.`
    }),
  listId: z.uuid().optional(),
  listUrl: z.string().optional(),
  listTitle: z.string().optional(),
  userName: z.string().optional().refine(upn => typeof upn === 'undefined' || validation.isValidUserPrincipalName(upn as string), {
    error: e => `'${e.input}' is not a valid UPN.`
  }),
  userId: z.uuid().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoWebAlertListCommand extends SpoCommand {
  public get name(): string {
    return commands.WEB_ALERT_LIST;
  }

  public get description(): string {
    return 'Lists all SharePoint list alerts';
  }

  public defaultProperties(): string[] | undefined {
    return ['ID', 'Title', 'UserPrincipalName'];
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(options => [options.listId, options.listUrl, options.listTitle].filter(x => x !== undefined).length <= 1, {
        error: `Specify either listId, listUrl, or listTitle, but not more than one.`
      })
      .refine(options => [options.userName, options.userId].filter(x => x !== undefined).length <= 1, {
        error: `Specify either userName or userId, but not both.`
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      const listParams = args.options.listId || args.options.listTitle || args.options.listUrl;
      const userParams = args.options.userName || args.options.userId;

      let message = `Retrieving alerts from site '${args.options.webUrl}'`;

      if (listParams) {
        message += ` for list '${listParams}'`;
      }

      if (userParams) {
        message += `${listParams ? ' and' : ''} for user '${userParams}'`;
      }

      await logger.logToStderr(`${message}...`);
    }

    let requestUrl = `${args.options.webUrl}/_api/web/alerts?$expand=List,User,List/Rootfolder,Item&$select=*,List/Id,List/Title,List/Rootfolder/ServerRelativeUrl,Item/ID,Item/FileRef,Item/Guid`;

    const filters: string[] = [];
    let listId: string | undefined;

    if (args.options.listId) {
      listId = args.options.listId;
    }
    else if (args.options.listUrl || args.options.listTitle) {
      listId = await spo.getListId(args.options.webUrl, args.options.listTitle, args.options.listUrl, logger, this.verbose);
    }

    if (listId) {
      filters.push(`List/Id eq guid'${formatting.encodeQueryParameter(listId)}'`);
    }

    if (args.options.userName) {
      filters.push(`User/UserPrincipalName eq '${formatting.encodeQueryParameter(args.options.userName)}'`);
    }
    else if (args.options.userId) {
      const userPrincipalName = await entraUser.getUpnByUserId(args.options.userId);
      filters.push(`User/UserPrincipalName eq '${formatting.encodeQueryParameter(userPrincipalName)}'`);
    }

    if (filters.length > 0) {
      requestUrl += `&$filter=${filters.join(' and ')}`;
    }

    try {
      const res = await odata.getAllItems<any>(requestUrl);
      res.forEach(alert => {
        if (alert.Item) {
          delete alert.Item['ID'];
        }

        if (cli.shouldTrimOutput(args.options.output)) {
          alert.UserPrincipalName = alert.User?.UserPrincipalName;
        }
      });
      await logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoWebAlertListCommand();