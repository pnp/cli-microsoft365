import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
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
  name: z.string().alias('n'),
  description: z.string().optional(),
  allowMembersEditMembership: z.boolean().optional(),
  onlyAllowMembersViewMembership: z.boolean().optional(),
  allowRequestToJoinLeave: z.boolean().optional(),
  autoAcceptRequestToJoinLeave: z.boolean().optional(),
  requestToJoinLeaveEmailSetting: z.string().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoGroupAddCommand extends SpoCommand {
  public get name(): string {
    return commands.GROUP_ADD;
  }

  public get description(): string {
    return 'Creates group in the specified site';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const requestOptions: CliRequestOptions = {
      url: `${args.options.webUrl}/_api/web/sitegroups`,
      headers: {
        accept: 'application/json;odata=nometadata'
      },
      responseType: 'json',
      data: {
        Title: args.options.name,
        Description: args.options.description,
        AllowMembersEditMembership: args.options.allowMembersEditMembership,
        OnlyAllowMembersViewMembership: args.options.onlyAllowMembersViewMembership,
        AllowRequestToJoinLeave: args.options.allowRequestToJoinLeave,
        AutoAcceptRequestToJoinLeave: args.options.autoAcceptRequestToJoinLeave,
        RequestToJoinLeaveEmailSetting: args.options.requestToJoinLeaveEmailSetting
      }
    };

    try {
      const response = await request.post<any>(requestOptions);
      await logger.log(response);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoGroupAddCommand();