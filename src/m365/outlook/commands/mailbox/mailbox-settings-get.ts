import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import GraphCommand from '../../../base/GraphCommand.js';
import { Logger } from '../../../../cli/Logger.js';
import commands from '../../commands.js';
import { validation } from '../../../../utils/validation.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { MailboxSettings } from '@microsoft/microsoft-graph-types';
import { accessToken } from '../../../../utils/accessToken.js';
import auth from '../../../../Auth.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  userId: z.uuid().optional().alias('i'),
  userName: z.string().refine(name => validation.isValidUserPrincipalName(name), {
    error: e => `'${e.input}' is not a valid UPN.`
  }).optional().alias('n')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class OutlookMailboxSettingsGetCommand extends GraphCommand {
  public get name(): string {
    return commands.MAILBOX_SETTINGS_GET;
  }

  public get description(): string {
    return `Get the user's mailbox settings`;
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(options => !(options.userId && options.userName), {
        error: 'Specify either userId or userName, but not both'
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const isAppOnlyAccessToken = accessToken.isAppOnlyAccessToken(auth.connection.accessTokens[auth.defaultResource].accessToken);

    let requestUrl = `${this.resource}/v1.0/me/mailboxSettings`;

    if (isAppOnlyAccessToken) {
      if (!args.options.userId && !args.options.userName) {
        throw 'When running with application permissions either userId or userName is required';
      }

      const userIdentifier = args.options.userId ?? args.options.userName;

      if (this.verbose) {
        await logger.logToStderr(`Retrieving mailbox settings for user ${userIdentifier}...`);
      }

      requestUrl = `${this.resource}/v1.0/users('${userIdentifier}')/mailboxSettings`;
    }
    else {
      if (args.options.userId || args.options.userName) {
        throw 'You can retrieve mailbox settings of other users only if CLI is authenticated in app-only mode';
      }
    }

    const requestOptions: CliRequestOptions = {
      url: requestUrl,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    try {
      const result = await request.get<MailboxSettings>(requestOptions);
      await logger.log(result);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new OutlookMailboxSettingsGetCommand();