import commands from '../commands';
import GlobalOptions from '../../../GlobalOptions';
import {
  CommandOption
} from '../../../Command';
import request from '../../../request';
import AzmgmtCommand from '../../base/AzmgmtCommand';
import { CommandInstance } from '../../../cli';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  environment: string;
  name: string;
  asAdmin: boolean;
}

class FlowGetCommand extends AzmgmtCommand {
  public get name(): string {
    return commands.FLOW_GET;
  }

  public get description(): string {
    return 'Gets information about the specified Microsoft Flow';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    if (this.verbose) {
      cmd.log(`Retrieving information about Microsoft Flow ${args.options.name}...`);
    }

    const requestOptions: any = {
      url: `${this.resource}providers/Microsoft.ProcessSimple/${args.options.asAdmin ? 'scopes/admin/' : ''}environments/${encodeURIComponent(args.options.environment)}/flows/${encodeURIComponent(args.options.name)}?api-version=2016-11-01`,
      headers: {
        accept: 'application/json'
      },
      json: true
    };

    request
      .get(requestOptions)
      .then((res: any): void => {
        if (args.options.output === 'json') {
          cmd.log(res);
        }
        else {
          const summary: any = {
            name: res.name,
            displayName: res.properties.displayName,
            description: res.properties.definitionSummary.description || '',
            triggers: Object.keys(res.properties.definition.triggers).join(', '),
            actions: Object.keys(res.properties.definition.actions).join(', ')
          };
          cmd.log(summary);
        }

        cb();
      }, (rawRes: any): void => this.handleRejectedODataJsonPromise(rawRes, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-n, --name <name>',
        description: 'The name of the Microsoft Flow to get information about'
      },
      {
        option: '-e, --environment <environment>',
        description: 'The name of the environment for which to retrieve available Flows'
      },
      {
        option: '--asAdmin',
        description: 'Set, to retrieve the Flow as admin'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }
}

module.exports = new FlowGetCommand();