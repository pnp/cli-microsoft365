import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { zod } from '../../../../utils/zod.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { validation } from '../../../../utils/validation.js';
import { Logger } from '../../../../cli/Logger.js';
import { formatting } from '../../../../utils/formatting.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { cli } from '../../../../cli/cli.js';

const options = globalOptionsZod
  .extend({
    id: zod.alias('i', z.string().optional()),
    userName: zod.alias('n', z.string().optional()),
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
    return 'Revokes Microsoft Entra user sessions';
  }
  public get schema(): z.ZodTypeAny | undefined {
    return options;
  }
  public getRefinedSchema(schema: typeof options): z.ZodEffects<any> | undefined {
    return schema
      .refine(options => !options.id !== !options.userName, {
        message: 'Specify either id or userName, but not both'
      })
      .refine(options => options.id || options.userName, {
        message: 'Specify either id or userName'
      })
      .refine(options => (!options.id && !options.userName) || options.userName || (options.id && validation.isValidGuid(options.id)), options => ({
        message: `The '${options.id}' must be a valid GUID`,
        path: ['id']
      }))
      .refine(options => (!options.id && !options.userName) || options.id || (options.userName && validation.isValidUserPrincipalName(options.userName)), options => ({
        message: `The '${options.userName}' must be a valid UPN`,
        path: ['id']
      }));
  }
  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const revokeUserSessions = async (): Promise<void> => {
      try {
        let userIdOrPrincipalName = args.options.id;

        if (args.options.userName) {
          // single user can be retrieved also by user principal name
          userIdOrPrincipalName = formatting.encodeQueryParameter(args.options.userName);
        }

        if (args.options.verbose) {
          await logger.logToStderr(`Invalidating all the refresh tokens for user ${userIdOrPrincipalName}...`);
        }

        // user principal name can start with $ but it violates the OData URL convention, so it must be enclosed in parenthesis and single quotes
        const requestUrl = userIdOrPrincipalName!.startsWith('%24')
          ? `${this.resource}/v1.0/users('${userIdOrPrincipalName}')/revokeSignInSessions`
          : `${this.resource}/v1.0/users/${userIdOrPrincipalName}/revokeSignInSessions`;

        const requestOptions: CliRequestOptions = {
          url: requestUrl,
          headers: {
            accept: 'application/json;odata.metadata=none'
          }
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
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to invalidate all the refresh tokens issued to applications for a user '${args.options.id || args.options.userName}'?` });

      if (result) {
        await revokeUserSessions();
      }
    }
  }
}

export default new EntraUserSessionRevokeCommand();