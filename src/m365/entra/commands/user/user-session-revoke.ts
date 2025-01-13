import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { zod } from '../../../../utils/zod.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { validation } from '../../../../utils/validation.js';
import { Logger } from '../../../../cli/Logger.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { cli } from '../../../../cli/cli.js';
import { formatting } from '../../../../utils/formatting.js';

const options = globalOptionsZod
  .extend({
    userId: zod.alias('i', z.string().refine(id => validation.isValidGuid(id), id => ({
      message: `'${id}' is not a valid GUID.`
    })).optional()),
    userName: zod.alias('n', z.string().refine(name => validation.isValidUserPrincipalName(name), name => ({
      message: `'${name}' is not a valid UPN.`
    })).optional()),
    force: zod.alias('f', z.boolean().optional())
  })
  .strict();

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class EntraUserSessionRevokeCommand extends GraphCommand {
  public get name(): string {
    return commands.USER_SESSION_REVOKE;
  }
  public get description(): string {
    return 'Revokes all sign-in sessions for a given user';
  }
  public get schema(): z.ZodTypeAny | undefined {
    return options;
  }
  public getRefinedSchema(schema: typeof options): z.ZodEffects<any> | undefined {
    return schema
      .refine(options => [options.userId, options.userName].filter(o => o !== undefined).length === 1, {
        message: `Specify either 'userId' or 'userName'.`
      });
  }
  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const revokeUserSessions = async (): Promise<void> => {
      try {
        const userIdentifier = args.options.userId ?? args.options.userName;

        if (this.verbose) {
          await logger.logToStderr(`Invalidating all the refresh tokens for user ${userIdentifier}...`);
        }

        const requestOptions: CliRequestOptions = {
          url: `${this.resource}/v1.0/users('${formatting.encodeQueryParameter(userIdentifier!)}')/revokeSignInSessions`,
          headers: {
            accept: 'application/json;odata.metadata=none'
          },
          responseType: 'json',
          data: {}
        };

        await request.post(requestOptions);
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };

    if (args.options.force) {
      await revokeUserSessions();
    }
    else {
      const result = await cli.promptForConfirmation({ message: `This will revoke all sessions for the user '${args.options.userId || args.options.userName}', requiring the user to re-sign in from all devices. Are you sure?` });

      if (result) {
        await revokeUserSessions();
      }
    }
  }
}

export default new EntraUserSessionRevokeCommand();