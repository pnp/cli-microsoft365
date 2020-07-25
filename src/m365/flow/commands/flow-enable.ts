import commands from '../commands';
import GlobalOptions from '../../../GlobalOptions';
import {
  CommandOption
} from '../../../Command';
import request from '../../../request';
import AzmgmtCommand from '../../base/AzmgmtCommand';
import * as chalk from 'chalk';
import { CommandInstance } from '../../../cli';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  name: string;
  environment: string;
  asAdmin: boolean;
}

class FlowEnableCommand extends AzmgmtCommand {
  public get name(): string {
    return commands.FLOW_ENABLE;
  }

  public get description(): string {
    return 'Enables specified Microsoft Flow';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    if (this.verbose) {
      cmd.log(`Enables Microsoft Flow ${args.options.name}...`);
    }

    const requestOptions: any = {
      url: `${this.resource}providers/Microsoft.ProcessSimple/${args.options.asAdmin ? 'scopes/admin/' : ''}environments/${encodeURIComponent(args.options.environment)}/flows/${encodeURIComponent(args.options.name)}/start?api-version=2016-11-01`,
      headers: {
        accept: 'application/json'
      },
      json: true
    };

    request
      .post(requestOptions)
      .then((): void => {

        if (this.verbose) {
          cmd.log(chalk.green('DONE'));
        }

        cb();
      }, (rawRes: any): void => this.handleRejectedODataJsonPromise(rawRes, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-n, --name <name>',
        description: 'The name of the Microsoft Flow to enable'
      },
      {
        option: '-e, --environment <environment>',
        description: 'The name of the environment for which to enable Flow'
      },
      {
        option: '--asAdmin',
        description: 'Set, to enable the Flow as admin'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }
}

module.exports = new FlowEnableCommand();