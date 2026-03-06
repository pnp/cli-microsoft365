import auth from '../../../../Auth.js';
import { Logger } from '../../../../cli/Logger.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { accessToken } from '../../../../utils/accessToken.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { cli } from '../../../../cli/cli.js';
import { globalOptionsZod } from '../../../../Command.js';
import { z } from 'zod';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  id: z.string().alias('i'),
  userId: z.uuid().optional(),
  userName: z.string()
    .refine(upn => validation.isValidUserPrincipalName(upn) === true, {
      error: e => `'${e.input}' is not a valid user principal name for option 'userName'.`
    })
    .optional(),
  permanent: z.boolean().optional(),
  force: z.boolean().optional().alias('f')
});
declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class OutlookEventRemoveCommand extends GraphCommand {
  public get name(): string {
    return commands.EVENT_REMOVE;
  }

  public get description(): string {
    return 'Removes an event from a calendar';
  }

  public get schema(): z.ZodTypeAny {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const isAppOnlyAccessToken: boolean | undefined = accessToken.isAppOnlyAccessToken(auth.connection.accessTokens[auth.defaultResource].accessToken);
    const token = auth.connection.accessTokens[auth.defaultResource].accessToken;

    if (isAppOnlyAccessToken) {
      if (!args.options.userId && !args.options.userName) {
        throw `The option 'userId' or 'userName' is required when removing an event using application permissions.`;
      }

      if (args.options.userId && args.options.userName) {
        throw `Both options 'userId' and 'userName' cannot be used together when removing an event using application permissions.`;
      }
    }
    else {
      if (args.options.userId && args.options.userName) {
        throw `Both options 'userId' and 'userName' cannot be used together when removing an event using delegated permissions.`;
      }

      if (args.options.userId) {
        const currentUserId = accessToken.getUserIdFromAccessToken(token);
        if (args.options.userId !== currentUserId) {
          throw `You can only remove your own events when using delegated permissions. The specified userId '${args.options.userId}' does not match the current user '${currentUserId}'.`;
        }
      }

      if (args.options.userName) {
        const currentUserName = accessToken.getUserNameFromAccessToken(token);
        if (args.options.userName.toLowerCase() !== currentUserName.toLowerCase()) {
          throw `You can only remove your own events when using delegated permissions. The specified userName '${args.options.userName}' does not match the current user '${currentUserName}'.`;
        }
      }
    }

    let principalUrl = '';
    if (args.options.userId || args.options.userName) {
      principalUrl += `users/${args.options.userId || formatting.encodeQueryParameter(args.options.userName!)}`;
    }
    else {
      principalUrl += 'me';
    }

    const removeEvent = async (): Promise<void> => {
      try {
        if (this.verbose) {
          await logger.logToStderr(`Removing event with id '${args.options.id}'...`);
        }

        const requestOptions: CliRequestOptions = {
          url: `${this.resource}/v1.0/${principalUrl}/events/${args.options.id}`,
          headers: {
            accept: 'application/json;odata.metadata=none'
          }
        };

        if (args.options.permanent) {
          requestOptions.url += '/permanentDelete';
          await request.post(requestOptions);
        }
        else {
          await request.delete(requestOptions);
        }
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };

    if (args.options.force) {
      await removeEvent();
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove event with id '${args.options.id}'?` });

      if (result) {
        await removeEvent();
      }
    }
  }
}

export default new OutlookEventRemoveCommand();
