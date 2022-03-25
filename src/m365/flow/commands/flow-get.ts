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

interface Trigger {
  type: string;
  kind?: string;
}

interface Action {
  type: string;
  swaggerOperationId?: string;
}

interface Flow {
  actions?: string;
  description?: string;
  displayName?: string;
  properties: {
    displayName: string;
    definitionSummary: {
      actions: Action[];
      description: string;
      triggers: Trigger[];
    };
  },
  triggers: string;
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
      .get<Flow>(requestOptions)
      .then((res): void => {
        res.displayName = res.properties.displayName;
        res.description = res.properties.definitionSummary.description || '';
        res.triggers = res.properties.definitionSummary.triggers.map((t: Trigger) => (t.type + (t.kind ? "-" + t.kind : '')) as string).join(', ');
        res.actions = res.properties.definitionSummary.actions.map((a: Action) => (a.type + (a.swaggerOperationId ? "-" + a.swaggerOperationId : '')) as string).join(', ');

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
