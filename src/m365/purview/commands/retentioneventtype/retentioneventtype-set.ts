import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import { validation } from '../../../../utils/validation';
import request, { CliRequestOptions } from '../../../../request';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

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
      logger.log(`Updating retention event type with id ${args.options.id}`);
    }

    try {
      const requestBody = {
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
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new PurviewRetentionEventTypeSetCommand();