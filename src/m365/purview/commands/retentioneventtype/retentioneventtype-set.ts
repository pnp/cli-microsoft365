import auth from '../../../../Auth';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import { validation } from '../../../../utils/validation';
import request, { CliRequestOptions } from '../../../../request';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';
import { accessToken } from '../../../../utils/accessToken';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id: string;
  newDisplayName?: string;
  description?: string;
}

class PurviewRetentionEventTypeSetCommand extends GraphCommand {
  public get name(): string {
    return commands.RETENTIONEVENTTYPE_SET;
  }

  public get description(): string {
    return 'Update a retention event type';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        newDisplayName: typeof args.options.newDisplayName !== 'undefined',
        description: typeof args.options.description !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --id <id>'
      },
      {
        option: '-n, --newDisplayName [newDisplayName]'
      },
      {
        option: '-d, --description [description]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (!validation.isValidGuid(args.options.id)) {
          return `'${args.options.id}' is not a valid GUID.`;
        }

        if (!args.options.newDisplayName && !args.options.description) {
          return `Specify atleast one option to update.`;
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (accessToken.isAppOnlyAccessToken(auth.service.accessTokens[this.resource].accessToken)) {
      this.handleError('This command does not support application permissions.');
    }

    if (this.verbose) {
      logger.log(`Starting to update retention event type with id ${args.options.id}`);
    }

    const requestBody = {
      displayName: args.options.newDisplayName,
      description: args.options.description
    };

    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/beta/security/triggerTypes/retentionEventTypes/${args.options.id}`,
      headers: {
        accept: 'application/json'
      },
      responseType: 'json',
      data: requestBody
    };

    await request.patch(requestOptions);
  }
}

module.exports = new PurviewRetentionEventTypeSetCommand();