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
  environment: string;
  name: string;
  asAdmin: boolean;
}

class FlowGetCommand extends AzmgmtCommand {
  public get name(): string {
    return commands.GET;
  }

  public get description(): string {
    return 'Gets information about the specified Microsoft Flow';
  }

  public defaultProperties(): string[] | undefined {
    return ['name', 'displayName', 'description', 'triggers', 'actions'];
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    if (this.verbose) {
      logger.logToStderr(`Retrieving information about Microsoft Flow ${args.options.name}...`);
    }

    const requestOptions: any = {
      url: `${this.resource}providers/Microsoft.ProcessSimple/${args.options.asAdmin ? 'scopes/admin/' : ''}environments/${encodeURIComponent(args.options.environment)}/flows/${encodeURIComponent(args.options.name)}?api-version=2016-11-01`,
      headers: {
        accept: 'application/json'
      },
      responseType: 'json'
    };

    request
      .get(requestOptions)
      .then((res: any): void => {
        res.displayName = res.properties.displayName;
        res.description = res.properties.definitionSummary.description || '';
        res.triggers = Object.keys(res.properties.definition.triggers).join(', ');
        res.actions = Object.keys(res.properties.definition.actions).join(', ');

        logger.log(res);
        cb();
      }, (rawRes: any): void => this.handleRejectedODataJsonPromise(rawRes, logger, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-n, --name <name>'
      },
      {
        option: '-e, --environment <environment>'
      },
      {
        option: '--asAdmin'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }
}

module.exports = new FlowGetCommand();