import auth from '../../../../Auth.js';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { accessToken } from '../../../../utils/accessToken.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { cli } from '../../../../cli/cli.js';
import { validation } from '../../../../utils/validation.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id: string;
  userId?: string;
  userName?: string;
  force?: boolean
}

class OutlookMessageRemoveCommand extends GraphCommand {
  public get name(): string {
    return commands.MESSAGE_REMOVE;
  }

  public get description(): string {
    return 'Removes a specifc message from a mailbox';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        userId: typeof args.options.userId !== 'undefined',
        userName: typeof args.options.userName !== 'undefined',
        force: !!args.options.force
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --id <id>'
      },
      {
        option: '--userId [userId]'
      },
      {
        option: '--userName [userName]'
      },
      {
        option: '-f, --force'
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const isAppOnlyAccessToken: boolean | undefined = accessToken.isAppOnlyAccessToken(auth.service.accessTokens[this.resource].accessToken);
    let requestUrl = '';

    if (isAppOnlyAccessToken) {
      if (!args.options.userId && !args.options.userName) {
        throw `The option 'userId' or 'userName' is required when removing a message using application permissions`;
      }

      if (args.options.userId && args.options.userName) {
        throw `Both options 'userId' and 'userName' cannot be set when removing a message using application permissions`;
      }

      if (args.options.userId && !validation.isValidGuid(args.options.userId)) {
        throw `The value '${args.options.userId}' for 'userId' option is not a valid GUID`;
      }

      if (args.options.userName && !validation.isValidUserPrincipalName(args.options.userName)) {
        throw `The value '${args.options.userName}' for 'userName' option is not a valid user principal name`;
      }

      requestUrl += `users/${args.options.userId ? args.options.userId : args.options.userName}`;
    }
    else {
      if (args.options.userId && args.options.userName) {
        throw `Both options 'userId' and 'userName' cannot be set when removing a message using delegated permissions`;
      }

      if (args.options.userId && !validation.isValidGuid(args.options.userId)) {
        throw `The value '${args.options.userId}' for 'userId' option is not a valid GUID`;
      }

      if (args.options.userName && !validation.isValidUserPrincipalName(args.options.userName)) {
        throw `The value '${args.options.userName}' for 'userName' option is not a valid user principal name`;
      }

      if (args.options.userId || args.options.userName) {
        requestUrl += `users/${args.options.userId ? args.options.userId : args.options.userName}`;
      }
      else {
        requestUrl += 'me';
      }
    }

    const removeMessage = async (): Promise<void> => {
      try {
        if (this.verbose) {
          await logger.logToStderr(`Removing message with id ${args.options.id} using ${isAppOnlyAccessToken ? 'application permissions' : 'delegated permissions'}`);
        }

        const requestOptions: CliRequestOptions = {
          url: `${this.resource}/v1.0/${requestUrl}/messages/${args.options.id}`,
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