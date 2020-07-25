import commands from '../commands';
import GlobalOptions from '../../../GlobalOptions';
import {
  CommandOption
} from '../../../Command';
import { AzmgmtItemsListCommand } from '../../base/AzmgmtItemsListCommand';
import { CommandInstance } from '../../../cli';

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
}

module.exports = new FlowListCommand();