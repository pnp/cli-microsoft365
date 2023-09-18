import { Cli } from '../../../../cli/Cli.js';
import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import AzmgmtCommand from '../../../base/AzmgmtCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  environmentName: string;
  flowName: string;
  name: string;
}

class FlowRunCancelCommand extends AzmgmtCommand {
  public get name(): string {
    return commands.RUN_CANCEL;
  }

  public get description(): string {
    return 'Cancels a specific run of the specified Microsoft Flow';
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
        force: !!args.options.force
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-n, --name <name>'
      },
      {
        option: '--flowName <flowName>'
      },
      {
        option: '-e, --environmentName <environmentName>'
      },
      {
        option: '-f, --force'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (!validation.isValidGuid(args.options.flowName)) {
          return `${args.options.flowName} is not a valid GUID`;
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.log(`Cancelling run ${args.options.name} of Microsoft Flow ${args.options.flowName}...`);
    }

    const cancelFlow = async (): Promise<void> => {
      const requestOptions: CliRequestOptions = {
        url: `${this.resource}providers/Microsoft.ProcessSimple/environments/${formatting.encodeQueryParameter(args.options.environmentName)}/flows/${formatting.encodeQueryParameter(args.options.flowName)}/runs/${formatting.encodeQueryParameter(args.options.name)}/cancel?api-version=2016-11-01`,
        headers: {
          accept: 'application/json'
        },
        responseType: 'json'
      };

      try {
        await request.post(requestOptions);
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };

    if (args.options.force) {
      await cancelFlow();
    }
    else {
      const result = await Cli.promptForConfirmation(`Are you sure you want to cancel the flow run ${args.options.name}?`);

      if (result) {
        await cancelFlow();
      }
    }
  }
}

export default new FlowRunCancelCommand();