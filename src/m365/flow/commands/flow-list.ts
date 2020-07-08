import commands from '../commands';
import GlobalOptions from '../../../GlobalOptions';
import {
  CommandOption,
  CommandValidate
} from '../../../Command';
import { AzmgmtItemsListCommand } from '../../base/AzmgmtItemsListCommand';

const vorpal: Vorpal = require('../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  environment: string;
  asAdmin: boolean;
}

class FlowListCommand extends AzmgmtItemsListCommand<{ name: string, properties: { displayName: string } }> {
  public get name(): string {
    return commands.FLOW_LIST;
  }

  public get description(): string {
    return 'Lists Microsoft Flows in the given environment';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.asAdmin = args.options.asAdmin === true;
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const url: string = `${this.resource}providers/Microsoft.ProcessSimple${args.options.asAdmin ? '/scopes/admin' : ''}/environments/${encodeURIComponent(args.options.environment)}/flows?api-version=2016-11-01`;

    this
      .getAllItems(url, cmd, true)
      .then((): void => {
        if (this.items.length > 0) {
          if (args.options.output === 'json') {
            cmd.log(this.items);
          }
          else {
            cmd.log(this.items.map(f => {
              return {
                name: f.name,
                displayName: f.properties.displayName
              };
            }));
          }
        }
        else {
          if (this.verbose) {
            cmd.log('No Flows found');
          }
        }

        cb();
      }, (rawRes: any): void => this.handleRejectedODataJsonPromise(rawRes, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-e, --environment <environment>',
        description: 'The name of the environment for which to retrieve available Flows'
      },
      {
        option: '--asAdmin',
        description: 'Set, to list all Flows as admin. Otherwise will return only your own Flows'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.environment) {
        return 'Required option environment missing';
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(commands.FLOW_LIST).helpInformation());
    log(
      `  Remarks:

    ${chalk.yellow('Attention:')} This command is based on an API that is currently
    in preview and is subject to change once the API reached general
    availability.
  
    If the environment with the name you specified doesn't exist, you will get
    the ${chalk.grey('Access to the environment \'xyz\' is denied.')} error.

    By default, the ${chalk.blue(this.getCommandName())} command returns only your
    Flows. To list all Flows, use the ${chalk.blue('asAdmin')} option.
   
  Examples:
  
    List all your Flows in the given environment
      ${this.getCommandName()} --environment Default-d87a7535-dd31-4437-bfe1-95340acd55c5

    List all Flows in the given environment
      ${this.getCommandName()} --environment Default-d87a7535-dd31-4437-bfe1-95340acd55c5 --asAdmin
`);
  }
}

module.exports = new FlowListCommand();