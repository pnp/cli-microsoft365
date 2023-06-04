import * as chalk from 'chalk';
import { Cli } from '../../../../cli/Cli';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import { validation } from '../../../../utils/validation';
import AzmgmtCommand from '../../../base/AzmgmtCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  confirm: boolean;
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
        confirm: args.options.confirm
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-n, --name <name>'
      },
      {
        option: '-f, --flowName <flowName>'
      },
      {
        option: '-e, --environmentName <environmentName>'
      },
      {
        option: '--confirm'
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
      logger.logToStderr(`Resubmitting run ${args.options.name} of Microsoft Flow ${args.options.flowName}...`);
    }

    const resubmitFlow = async (): Promise<void> => {
      try {
        const triggerName = await this.getTriggerName(args.options.environmentName, args.options.flowName);

        if (this.debug) {
          logger.logToStderr(chalk.yellow(`Retrieved trigger: ${triggerName}`));
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

    if (args.options.confirm) {
      await resubmitFlow();
    }
    else {
      const result = await Cli.prompt<{ continue: boolean }>({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to resubmit the flow with run ${args.options.name}?`
      });

      if (result.continue) {
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

module.exports = new FlowRunResubmitCommand();