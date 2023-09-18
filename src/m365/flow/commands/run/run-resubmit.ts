import chalk from 'chalk';
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
  force: boolean;
  environmentName: string;
  flowName: string;
  name: string;
}

class FlowRunResubmitCommand extends AzmgmtCommand {
  public get name(): string {
    return commands.RUN_RESUBMIT;
  }

  public get description(): string {
    return 'Resubmits a specific flow run for the specified Microsoft Flow';
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
        force: args.options.force
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
      await logger.logToStderr(`Resubmitting run ${args.options.name} of Microsoft Flow ${args.options.flowName}...`);
    }

    const resubmitFlow = async (): Promise<void> => {
      try {
        const triggerName = await this.getTriggerName(args.options.environmentName, args.options.flowName);

        if (this.debug) {
          await logger.logToStderr(chalk.yellow(`Retrieved trigger: ${triggerName}`));
        }

        const requestOptions: CliRequestOptions = {
          url: `${this.resource}providers/Microsoft.ProcessSimple/environments/${formatting.encodeQueryParameter(args.options.environmentName)}/flows/${formatting.encodeQueryParameter(args.options.flowName)}/triggers/${formatting.encodeQueryParameter(triggerName)}/histories/${formatting.encodeQueryParameter(args.options.name)}/resubmit?api-version=2016-11-01`,
          headers: {
            accept: 'application/json'
          },
          responseType: 'json'
        };

        await request.post(requestOptions);
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };

    if (args.options.force) {
      await resubmitFlow();
    }
    else {
      const result = await Cli.promptForConfirmation(`Are you sure you want to resubmit the flow with run ${args.options.name}?`);

      if (result) {
        await resubmitFlow();
      }
    }
  }

  private async getTriggerName(environment: string, flow: string): Promise<string> {
    const requestOptions: CliRequestOptions = {
      url: `${this.resource}providers/Microsoft.ProcessSimple/environments/${formatting.encodeQueryParameter(environment)}/flows/${formatting.encodeQueryParameter(flow)}/triggers?api-version=2016-11-01`,
      headers: {
        accept: 'application/json'
      },
      responseType: 'json'
    };

    const res = await request.get<{ value: { name: string; }[]; }>(requestOptions);
    return res.value[0].name;
  }
}

export default new FlowRunResubmitCommand();