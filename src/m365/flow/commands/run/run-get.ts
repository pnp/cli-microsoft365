import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption
} from '../../../../Command';
import request from '../../../../request';
import AzmgmtCommand from '../../../base/AzmgmtCommand';
import { CommandInstance } from '../../../../cli';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  environment: string;
  flow: string;
  name: string;
}

class FlowRunGetCommand extends AzmgmtCommand {
  public get name(): string {
    return commands.FLOW_RUN_GET;
  }

  public get description(): string {
    return 'Gets information about a specific run of the specified Microsoft Flow';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    if (this.verbose) {
      cmd.log(`Retrieving information about run ${args.options.name} of Microsoft Flow ${args.options.flow}...`);
    }

    const requestOptions: any = {
      url: `${this.resource}providers/Microsoft.ProcessSimple/environments/${encodeURIComponent(args.options.environment)}/flows/${encodeURIComponent(args.options.flow)}/runs/${encodeURIComponent(args.options.name)}?api-version=2016-11-01`,
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
            startTime: res.properties.startTime,
            endTime: res.properties.endTime || '',
            status: res.properties.status,
            triggerName: res.properties.trigger.name
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
        description: 'The name of the run to get information about'
      },
      {
        option: '-f, --flow <flow>',
        description: 'The name of the Microsoft Flow for which to retrieve information'
      },
      {
        option: '-e, --environment <environment>',
        description: 'The name of the environment where the Flow is located'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }
}

module.exports = new FlowRunGetCommand();