import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import { formatting } from '../../../../utils/formatting.js';
import { odata } from '../../../../utils/odata.js';
import { validation } from '../../../../utils/validation.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  webUrl: z.string()
    .refine(url => validation.isValidSharePointUrl(url) === true, {
      error: e => `'${e.input}' is not a valid SharePoint Online site URL.`
    })
    .alias('u'),
  groupId: z.number().optional(),
  groupName: z.string().optional()
});

export declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoGroupMemberListCommand extends SpoCommand {
  public get name(): string {
    return commands.GROUP_MEMBER_LIST;
  }

  public get description(): string {
    return 'Lists the members of a SharePoint Group';
  }

  public defaultProperties(): string[] | undefined {
    return ['Title', 'UserPrincipalName', 'Id', 'Email'];
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(o => [o.groupId, o.groupName].filter(x => x !== undefined).length === 1, {
        error: `Specify either 'groupId' or 'groupName', but not both.`,
        params: {
          customCode: 'optionSet',
          options: ['groupId', 'groupName']
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving the list of members from the SharePoint group :  ${args.options.groupId ? args.options.groupId : args.options.groupName}`);
    }

    const requestUrl: string = `${args.options.webUrl}/_api/web/sitegroups/${args.options.groupId
      ? `GetById('${args.options.groupId}')`
      : `GetByName('${formatting.encodeQueryParameter(args.options.groupName as string)}')`}/users`;

    try {
      const response = await odata.getAllItems<any>(requestUrl);
      await logger.log(response);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoGroupMemberListCommand();