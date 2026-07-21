import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { Logger } from '../../../../cli/Logger.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

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

class PlannerRosterMemberAddCommand extends GraphCommand {
  public get name(): string {
    return commands.ROSTER_MEMBER_ADD;
  }

  public get description(): string {
    return 'Adds a user to a Microsoft Planner Roster';
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
      await logger.logToStderr('Adding a user to a Microsoft Planner Roster');
    }

    try {
      const userId = await this.getUserId(logger, args);

      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/beta/planner/rosters/${args.options.rosterId}/members`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        data: {
          userId: userId
        },
        responseType: 'json'
      };

      const response = await request.post(requestOptions);
      await logger.log(response);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getUserId(logger: Logger, args: CommandArgs): Promise<string> {
    if (this.verbose) {
      await logger.logToStderr('Getting the user ID');
    }

    if (args.options.userId) {
      return args.options.userId;
    }

    const userId = await entraUser.getUserIdByUpn(args.options.userName!);

    return userId;
  }
}

export default new PlannerRosterMemberAddCommand();
