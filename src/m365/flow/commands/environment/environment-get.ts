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
  name: string;
}

class FlowEnvironmentGetCommand extends AzmgmtCommand {
  public get name(): string {
    return commands.FLOW_ENVIRONMENT_GET;
  }

  public get description(): string {
    return 'Gets information about the specified Microsoft Flow environment';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    if (this.verbose) {
      cmd.log(`Retrieving information about Microsoft Flow environment ${args.options.name}...`);
    }

    const requestOptions: any = {
      url: `${this.resource}providers/Microsoft.ProcessSimple/environments/${encodeURIComponent(args.options.name)}?api-version=2016-11-01`,
      headers: {
        accept: 'application/json'
      },
      json: true
    };

    request
      .get(requestOptions)
      .then((res: any): void => {
        cmd.log(res);

        cb();
      }, (rawRes: any): void => this.handleRejectedODataJsonPromise(rawRes, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-n, --name <name>',
        description: 'The name of the environment to get information about'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }
}

module.exports = new FlowEnvironmentGetCommand();