import { Logger } from '../../../cli';
import {
    CommandOption
} from '../../../Command';
import GlobalOptions from '../../../GlobalOptions';
import { AzmgmtItemsListCommand } from '../../base/AzmgmtItemsListCommand';
import commands from '../commands';

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

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    const url: string = `${this.resource}providers/Microsoft.ProcessSimple${args.options.asAdmin ? '/scopes/admin' : ''}/environments/${encodeURIComponent(args.options.environment)}/flows?api-version=2016-11-01`;

    this
      .getAllItems(url, logger, true)
      .then((): void => {
        if (this.items.length > 0) {
          if (args.options.output === 'json') {
            logger.log(this.items);
          }
          else {
            logger.log(this.items.map(f => {
              return {
                name: f.name,
                displayName: f.properties.displayName
              };
            }));
          }
        }
        else {
          if (this.verbose) {
            logger.log('No Flows found');
          }
        }

        cb();
      }, (rawRes: any): void => this.handleRejectedODataJsonPromise(rawRes, logger, cb));
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