import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id: string;
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

        if (!args.options.description) {
          return 'Specify at least one option to update.';
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.log(`Updating retention event type with id ${args.options.id}`);
    }

    try {
      const requestBody = {
        description: args.options.description
      };

      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/security/triggerTypes/retentionEventTypes/${args.options.id}`,
        headers: {
          accept: 'application/json'
        },
        responseType: 'json',
        data: requestBody
      };

      await request.patch(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new PurviewRetentionEventTypeSetCommand();