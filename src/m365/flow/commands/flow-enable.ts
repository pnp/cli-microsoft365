import commands from '../commands';
import GlobalOptions from '../../../GlobalOptions';
import {
  CommandOption,
  CommandValidate
} from '../../../Command';
import request from '../../../request';
import AzmgmtCommand from '../../base/AzmgmtCommand';
import * as os from 'os';

const vorpal: Vorpal = require('../../../vorpal-init');

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
          cmd.log(vorpal.chalk.green('DONE'));
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

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.name) {
        return 'Required option name missing';
      }

      if (!args.options.environment) {
        return 'Required option environment missing';
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(commands.FLOW_ENABLE).helpInformation());
    log(
      `  Remarks:

    ${chalk.yellow('Attention:')} This command is based on an API that is currently
    in preview and is subject to change once the API reached general
    availability.
  
    By default, the command will try to enable Microsoft Flows you own.
    If you want to enable Flow owned by another user, use the ${chalk.blue('asAdmin')}
    flag.

    If the environment with the name you specified doesn't exist, you will get
    the ${chalk.grey('Access to the environment \'xyz\' is denied.')} error.

    If the Microsoft Flow with the name you specified doesn't exist, you will
    get the ${chalk.grey(`The caller with object id \'abc\' does not have permission${os.EOL}` +
        '    for connection \'xyz\' under Api \'shared_logicflows\'.')} error.
    If you try to enable a non-existing flow as admin, you will get the
    ${chalk.grey('Could not find flow \'xyz\'.')} error.
   
  Examples:
  
    Enables Microsoft Flow owned by the currently signed-in user
      ${this.getCommandName()} --environment Default-d87a7535-dd31-4437-bfe1-95340acd55c5 --name 3989cb59-ce1a-4a5c-bb78-257c5c39381d

    Enables Microsoft Flow owned by another user
      ${this.getCommandName()} --environment Default-d87a7535-dd31-4437-bfe1-95340acd55c5 --name 3989cb59-ce1a-4a5c-bb78-257c5c39381d --asAdmin
`);
  }
}

module.exports = new FlowEnableCommand();