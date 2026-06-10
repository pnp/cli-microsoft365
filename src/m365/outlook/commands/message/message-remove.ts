import auth from '../../../../Auth.js';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { accessToken } from '../../../../utils/accessToken.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { cli } from '../../../../cli/cli.js';
import { validation } from '../../../../utils/validation.js';
import { formatting } from '../../../../utils/formatting.js';
import { z } from 'zod';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  id: z.string().alias('i'),
  userId: z.string()
    .refine(userId => validation.isValidGuid(userId), {
      error: e => `Value '${e.input}' is not a valid GUID for option 'userId'.`
    }).optional(),
  userName: z.string()
    .refine(userName => validation.isValidUserPrincipalName(userName), {
      error: e => `Value '${e.input}' is not a valid user principal name for option 'userName'.`
    }).optional(),
  force: z.boolean().optional().alias('f')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class OutlookMessageRemoveCommand extends GraphCommand {
  public get name(): string {
    return commands.MESSAGE_REMOVE;
  }

  public get description(): string {
    return 'Permanently removes a specific message from a mailbox';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const isAppOnlyAccessToken: boolean | undefined = accessToken.isAppOnlyAccessToken(auth.connection.accessTokens[auth.defaultResource].accessToken);
    let principalUrl = '';

    if (isAppOnlyAccessToken) {
      if (!args.options.userId && !args.options.userName) {
        throw `The option 'userId' or 'userName' is required when removing a message using application permissions.`;
      }

      if (args.options.userId && args.options.userName) {
        throw `Both options 'userId' and 'userName' cannot be used together when removing a message using application permissions.`;
      }
    }
    else {
      if (args.options.userId && args.options.userName) {
        throw `Both options 'userId' and 'userName' cannot be used together when removing a message using delegated permissions.`;
      }
    }

    if (args.options.userId || args.options.userName) {
      principalUrl += `users/${args.options.userId || formatting.encodeQueryParameter(args.options.userName!)}`;
    }
    else {
      principalUrl += 'me';
    }

    const removeMessage = async (): Promise<void> => {
      try {
        if (this.verbose) {
          await logger.logToStderr(`Removing message with id '${args.options.id}' using ${isAppOnlyAccessToken ? 'application' : 'delegated'} permissions.`);
        }

        const requestOptions: CliRequestOptions = {
          url: `${this.resource}/v1.0/${principalUrl}/messages/${args.options.id}`,
          headers: {
            accept: 'application/json;odata.metadata=none'
          }
        };

        await request.delete(requestOptions);
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };

    if (args.options.force) {
      await removeMessage();
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove message with id '${args.options.id}'?` });

      if (result) {
        await removeMessage();
      }
    }
  }
}

export default new OutlookMessageRemoveCommand();