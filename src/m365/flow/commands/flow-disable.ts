import * as chalk from 'chalk';
import { Logger } from '../../../cli';
import {
    CommandOption
} from '../../../Command';
import GlobalOptions from '../../../GlobalOptions';
import request from '../../../request';
import AzmgmtCommand from '../../base/AzmgmtCommand';
import commands from '../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  name: string;
  environment: string;
  asAdmin: boolean;
}

class FlowDisableCommand extends AzmgmtCommand {
  public get name(): string {
    return commands.FLOW_DISABLE;
  }

  public get description(): string {
    return 'Disables specified Microsoft Flow';
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    if (this.verbose) {
      logger.log(`Disables Microsoft Flow ${args.options.name}...`);
    }

    const requestOptions: any = {
      url: `${this.resource}providers/Microsoft.ProcessSimple/${args.options.asAdmin ? 'scopes/admin/' : ''}environments/${encodeURIComponent(args.options.environment)}/flows/${encodeURIComponent(args.options.name)}/stop?api-version=2016-11-01`,
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
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-n, --name <name>',
        description: 'The name of the Microsoft Flow to disable'
      },
      {
        option: '-e, --environment <environment>',
        description: 'The name of the environment for which to disable Flow'
      },
      {
        option: '--asAdmin',
        description: 'Set, to disable the Flow as admin'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }
}

module.exports = new FlowDisableCommand();