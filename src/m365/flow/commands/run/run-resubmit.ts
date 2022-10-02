import * as chalk from 'chalk';
import { Cli } from '../../../../cli/Cli';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { validation } from '../../../../utils/validation';
import AzmgmtCommand from '../../../base/AzmgmtCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  confirm: boolean;
  environment: string;
  flow: string;
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
        option: '-f, --flow <flow>'
      },
      {
        option: '-e, --environment <environment>'
      },
      {
        option: '--confirm'
      }
    );
  }
  
  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (!validation.isValidGuid(args.options.flow)) {
          return `${args.options.flow} is not a valid GUID`;
        }
    
        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      logger.logToStderr(`Resubmitting run ${args.options.name} of Microsoft Flow ${args.options.flow}...`);
    }

    const resubmitFlow: () => Promise<void> = async (): Promise<void> => {
      try {
        const triggerName = await this.getTriggerName(args.options.environment, args.options.flow);

        if (this.debug) {
          logger.logToStderr(chalk.yellow(`Retrieved trigger: ${triggerName}`));
        }

        const requestOptions: any = {
          url: `${this.resource}providers/Microsoft.ProcessSimple/environments/${encodeURIComponent(args.options.environment)}/flows/${encodeURIComponent(args.options.flow)}/triggers/${encodeURIComponent(triggerName)}/histories/${encodeURIComponent(args.options.name)}/resubmit?api-version=2016-11-01`,
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

  private getTriggerName(environment: string, flow: string): Promise<string> {
    const requestOptions: any = {
      url: `${this.resource}providers/Microsoft.ProcessSimple/environments/${encodeURIComponent(environment)}/flows/${encodeURIComponent(flow)}/triggers?api-version=2016-11-01`,
      headers: {
        accept: 'application/json'
      },
      responseType: 'json'
    };

    return request
      .get<{ value: { name: string; }[]; }>(requestOptions)
      .then((res: { value: { name: string }[]; }): Promise<string> => Promise.resolve(res.value[0].name));
  }
}

module.exports = new FlowRunResubmitCommand();