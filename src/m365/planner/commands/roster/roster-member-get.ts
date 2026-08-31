import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { Logger } from '../../../../cli/Logger.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { entraUser } from '../../../../utils/entraUser.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  rosterId: z.string(),
  userId: z.string()
    .refine(val => validation.isValidGuid(val), {
      message: 'The value is not a valid GUID.'
    })
    .optional(),
  userName: z.string()
    .refine(val => validation.isValidUserPrincipalName(val), {
      message: 'The value is not a valid user principal name (UPN).'
    })
    .optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class PlannerRosterMemberGetCommand extends GraphCommand {
  public get name(): string {
    return commands.ROSTER_MEMBER_GET;
  }

  public get description(): string {
    return 'Gets a member of the specified Microsoft Planner Roster';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodType | undefined {
    return schema
      .refine(opts => [opts.userId, opts.userName].filter(x => x !== undefined).length === 1, {
        message: `Specify exactly one of the following options: 'userId' or 'userName'.`,
        params: {
          customCode: 'optionSet',
          options: ['userId', 'userName']
        }
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Retrieving member ${args.options.userName || args.options.userId} from the Microsoft Planner Roster`);
    }
    try {
      const userId = await this.getUserId(args);

      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/beta/planner/rosters/${args.options.rosterId}/members/${userId}`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json'
      };

      const response = await request.get(requestOptions);
      await logger.log(response);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getUserId(args: CommandArgs): Promise<string> {
    if (args.options.userId) {
      return args.options.userId;
    }

    return entraUser.getUserIdByUpn(args.options.userName!);
  }
}

export default new PlannerRosterMemberGetCommand();
