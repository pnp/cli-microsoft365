import GlobalOptions from '../../../../GlobalOptions.js';
import { Logger } from '../../../../cli/Logger.js';
import request, { CliRequestOptions } from '../../../../request.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { MultitenantOrganization } from './MultitenantOrganization.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  displayName?: string;
  description?: string;
}

class EntraMultitenantSetCommand extends GraphCommand {
  public get name(): string {
    return commands.MULTITENANT_SET;
  }

  public get description(): string {
    return 'Updates the properties of a multitenant organization';
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
        displayName: typeof args.options.displayName !== 'undefined',
        description: typeof args.options.description !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-n, --displayName [displayName]'
      },
      {
        option: '-d, --description [description]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (!args.options.displayName && !args.options.description) {
          return 'Specify either displayName or description or both.';
        }

        return true;
      }
    );
  }

  #initTypes(): void {
    this.types.string.push('displayName', 'description');
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr('Updating multitenant organization...');
    }

    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/tenantRelationships/multiTenantOrganization`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json',
      data: {
        description: args.options.description,
        displayName: args.options.displayName
      }
    };

    try {
      await request.patch<MultitenantOrganization>(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new EntraMultitenantSetCommand();