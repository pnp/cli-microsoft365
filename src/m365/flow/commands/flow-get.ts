import { Logger } from '../../../cli/Logger';
import GlobalOptions from '../../../GlobalOptions';
import request from '../../../request';
import AzmgmtCommand from '../../base/AzmgmtCommand';
import commands from '../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  environmentName: string;
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

  constructor() {
    super();
  
    this.#initOptions();
  }
  
  #initOptions(): void {
    this.options.unshift(
      {
        option: '-n, --name <name>'
      },
      {
        option: '-e, --environmentName <environmentName>'
      },
      {
        option: '--asAdmin'
      }      
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      logger.logToStderr(`Retrieving information about Microsoft Flow ${args.options.name}...`);
    }

    const requestOptions: any = {
      url: `${this.resource}providers/Microsoft.ProcessSimple/${args.options.asAdmin ? 'scopes/admin/' : ''}environments/${encodeURIComponent(args.options.environmentName)}/flows/${encodeURIComponent(args.options.name)}?api-version=2016-11-01`,
      headers: {
        accept: 'application/json'
      },
      responseType: 'json'
    };

    try {
      const res = await request.get<Flow>(requestOptions);

      res.displayName = res.properties.displayName;
      res.description = res.properties.definitionSummary.description || '';
      res.triggers = res.properties.definitionSummary.triggers.map((t: Trigger) => (t.type + (t.kind ? "-" + t.kind : '')) as string).join(', ');
      res.actions = res.properties.definitionSummary.actions.map((a: Action) => (a.type + (a.swaggerOperationId ? "-" + a.swaggerOperationId : '')) as string).join(', ');

      logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new FlowGetCommand();
