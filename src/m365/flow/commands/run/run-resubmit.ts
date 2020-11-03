import * as chalk from 'chalk';
import { Cli, Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import Utils from '../../../../Utils';
import AzmgmtCommand from '../../../base/AzmgmtCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  environment: string;
  flow: string;
  name: string;
}

class FlowRunResubmitCommand extends AzmgmtCommand {
  public get name(): string {
    return commands.FLOW_RUN_RESUBMIT;
  }

  public get description(): string {
    return 'Resubmits a specific flow run for the specified Microsoft Flow';
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    if (this.verbose) {
      logger.log(`Resubmitting run ${args.options.name} of Microsoft Flow ${args.options.flow}...`);
    }

    const resubmitFlow: () => void = (): void => {
      this._getTriggerName(args.options.environment, args.options.flow).then((triggerName: string) => {
        const requestOptions: any = {
          url: `${this.resource}providers/Microsoft.ProcessSimple/environments/${encodeURIComponent(args.options.environment)}/flows/${encodeURIComponent(args.options.flow)}/triggers/${encodeURIComponent(triggerName)}/histories/${encodeURIComponent(args.options.name)}/resubmit?api-version=2016-11-01`,
          headers: {
            accept: 'application/json'
          },
          responseType: 'json'
        };

        request
          .post(requestOptions)
          .then((): void => {
            if (this.verbose) {
              logger.log(chalk.green('DONE'));
            }

            cb();
          }, (rawRes: any): void => this.handleRejectedODataJsonPromise(rawRes, logger, cb));
      }, (rawRes: any): void => this.handleRejectedODataJsonPromise(rawRes, logger, cb));
    };

    if (args.options.confirm) {
      resubmitFlow();
    }
    else {
      Cli.prompt({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to cancel the flow with run ${args.options.name}?`
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

  private _getTriggerName(environment: string, flow: string): Promise<string> {
    return new Promise<string>((resolve, reject) => {
      const requestOptions: any = {
        url: `${this.resource}providers/Microsoft.ProcessSimple/environments/${encodeURIComponent(environment)}/flows/${encodeURIComponent(flow)}/triggers?api-version=2016-11-01`,
        headers: {
          accept: 'application/json'
        },
        responseType: 'json'
      };

      request
        .get(requestOptions)
        .then((res: any): void => {
          return resolve(res['value'][0]['name']);
        }).catch((ex: ExceptionInformation) => {
          reject(ex);
        });
    });
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-n, --name <name>',
        description: 'The name of the run to resubmit'
      },
      {
        option: '-f, --flow <flow>',
        description: 'he name of the Microsoft Flow to resubmit'
      },
      {
        option: '-e, --environment <environment>',
        description: 'The name of the environment where the Flow is located'
      },
      {
        option: '--confirm',
        description: 'Don\'t prompt for confirming cancelling the Flow'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (!Utils.isValidGuid(args.options.flow)) {
      return `${args.options.flow} is not a valid GUID`;
    }
    return true;
  }
}

module.exports = new FlowRunResubmitCommand();