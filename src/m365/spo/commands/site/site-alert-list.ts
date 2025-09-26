import commands from '../../commands.js';
import { Logger } from '../../../../cli/Logger.js';
import SpoCommand from '../../../base/SpoCommand.js';
import { globalOptionsZod } from '../../../../Command.js';
import { z } from 'zod';
import { zod } from '../../../../utils/zod.js';
import { validation } from '../../../../utils/validation.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { formatting } from '../../../../utils/formatting.js';
import { odata } from '../../../../utils/odata.js';

export const options = globalOptionsZod
  .extend({
    webUrl: zod.alias('u', z.string()
      .refine(url => validation.isValidSharePointUrl(url) === true, {
        message: 'webUrl is not a valid SharePoint site URL.'
      })),
    listId: z.string()
      .refine(id => validation.isValidGuid(id), id => ({
        message: `'${id}' is not a valid GUID.`
      })).optional(),
    listUrl: z.string().optional(),
    listTitle: z.string().optional(),
    userName: z.string().refine(upn => validation.isValidUserPrincipalName(upn), upn => ({
      message: `'${upn}' is not a valid UPN.`
    })).optional(),
    userId: z.string().refine(id => validation.isValidGuid(id), id => ({
      message: `'${id}' is not a valid GUID.`
    })).optional()
  })
  .strict();

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoSiteAlertListCommand extends SpoCommand {
  public get name(): string {
    return commands.SITE_ALERT_LIST;
  }

  public get description(): string {
    return 'Lists all SharePoint list alerts';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'title', 'userId'];
  }

  public get schema(): z.ZodTypeAny | undefined {
    return options;
  }

  public getRefinedSchema(schema: z.ZodTypeAny): z.ZodEffects<any> | undefined {
    return schema
      .refine(options => [options.listId, options.listUrl, options.listTitle].filter(x => x !== undefined).length <= 1, {
        message: `Specify either listId, listUrl, or listTitle, but not more than one.`
      })
      .refine(options => [options.userName, options.userId].filter(x => x !== undefined).length <= 1, {
        message: `Specify either userName or userId, but not both.`
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

    if (args.options.listId) {
      filters.push(`List/Id eq guid'${formatting.encodeQueryParameter(args.options.listId)}'`);
    }
    else if (args.options.listUrl) {
      const listServerRelativeUrl = urlUtil.getServerRelativePath(args.options.webUrl, args.options.listUrl);
      filters.push(`List/RootFolder/ServerRelativeUrl eq '${formatting.encodeQueryParameter(listServerRelativeUrl)}'`);
    }
    else if (args.options.listTitle) {
      filters.push(`List/Title eq '${formatting.encodeQueryParameter(args.options.listTitle)}'`);
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
      const res = await odata.getAllItems<any>(`${requestUrl}`);
      await logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoSiteAlertListCommand();