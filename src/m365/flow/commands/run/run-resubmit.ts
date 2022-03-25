import * as chalk from 'chalk';
import { Cli, Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { validation } from '../../../../utils';
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

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.confirm = args.options.confirm;
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    if (this.verbose) {
      logger.logToStderr(`Resubmitting run ${args.options.name} of Microsoft Flow ${args.options.flow}...`);
    }

    const resubmitFlow: () => void = (): void => {
      this
        .getTriggerName(args.options.environment, args.options.flow)
        .then((triggerName: string): Promise<void> => {
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

          return request.post(requestOptions);
        })
        .then(_ => cb(), (rawRes: any): void => this.handleRejectedODataJsonPromise(rawRes, logger, cb));
    };

    if (args.options.confirm) {
      resubmitFlow();
    }
    else {
      Cli.prompt({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to resubmit the flow with run ${args.options.name}?`
      }, (result: { continue: boolean }): void => {
        if (!result.continue) {
          cb();
        }
        else {
          resubmitFlow();
        }
      });
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

  public options(): CommandOption[] {
    const options: CommandOption[] = [
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
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (!validation.isValidGuid(args.options.flow)) {
      return `${args.options.flow} is not a valid GUID`;
    }

    return true;
  }
}

module.exports = new FlowRunResubmitCommand();