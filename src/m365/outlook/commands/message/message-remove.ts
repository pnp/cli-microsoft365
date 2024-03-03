import auth from '../../../../Auth.js';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { accessToken } from '../../../../utils/accessToken.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { cli } from '../../../../cli/cli.js';
import { validation } from '../../../../utils/validation.js';
import { formatting } from '../../../../utils/formatting.js';

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
    return 'Removes a specific message from a mailbox';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initTypes();
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

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.userId && !validation.isValidGuid(args.options.userId)) {
          return `Value '${args.options.userId}' is not a valid GUID for option 'userId'.`;
        }

        if (args.options.userName && !validation.isValidUserPrincipalName(args.options.userName)) {
          return `Value '${args.options.userName}' is not a valid user principal name for option 'userName'.`;
        }

        return true;
      }
    );
  }

  #initTypes(): void {
    this.types.string.push('id', 'userId', 'userName');
    this.types.boolean.push('force');
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