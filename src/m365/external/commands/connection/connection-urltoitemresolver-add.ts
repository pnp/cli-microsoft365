import { ExternalConnectors } from '@microsoft/microsoft-graph-types/microsoft-graph';
import GlobalOptions from '../../../../GlobalOptions.js';
import { Logger } from '../../../../cli/Logger.js';
import request, { CliRequestOptions } from '../../../../request.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  externalConnectionId: string;
  baseUrls: string;
  urlPattern: string;
  itemId: string;
  priority: number;
}

class ExternalConnectionUrlToItemResolverAddCommand extends GraphCommand {
  public get name(): string {
    return commands.CONNECTION_URLTOITEMRESOLVER_ADD;
  }

  public get description(): string {
    return 'Adds a URL to item resolver to an external connection';
  }

  constructor() {
    super();

    this.#initOptions();
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-c, --externalConnectionId <externalConnectionId>'
      },
      {
        option: '--baseUrls <baseUrls>'
      },
      {
        option: '--urlPattern <urlPattern>'
      },
      {
        option: '-i, --itemId <itemId>'
      },
      {
        option: '-p, --priority <priority>'
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const baseUrls: string[] = args.options.baseUrls.split(',').map(b => b.trim());

    const itemIdResolver: ExternalConnectors.ItemIdResolver = {
      itemId: args.options.itemId,
      priority: args.options.priority,
      urlMatchInfo: {
        baseUrls: baseUrls,
        urlPattern: args.options.urlPattern
      }
    };
    // not a part of the type definition, but required by the API
    (itemIdResolver as any)['@odata.type'] = '#microsoft.graph.externalConnectors.itemIdResolver';

    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/external/connections/${args.options.externalConnectionId}`,
      headers: {
        accept: 'application/json;odata.metadata=none',
        'content-type': 'application/json'
      },
      responseType: 'json',
      data: {
        activitySettings: {
          urlToItemResolvers: [itemIdResolver]
        }
      } as ExternalConnectors.ExternalConnection
    };

    try {
      await request.patch(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new ExternalConnectionUrlToItemResolverAddCommand();