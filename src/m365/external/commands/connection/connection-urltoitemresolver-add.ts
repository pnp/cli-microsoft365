import { ExternalConnectors } from '@microsoft/microsoft-graph-types/microsoft-graph';
import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { Logger } from '../../../../cli/Logger.js';
import request, { CliRequestOptions } from '../../../../request.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  externalConnectionId: z.string().alias('c'),
  baseUrls: z.string(),
  urlPattern: z.string(),
  itemId: z.string().alias('i'),
  priority: z.number().alias('p')
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class ExternalConnectionUrlToItemResolverAddCommand extends GraphCommand {
  public get name(): string {
    return commands.CONNECTION_URLTOITEMRESOLVER_ADD;
  }

  public get description(): string {
    return 'Adds a URL to item resolver to an external connection';
  }

  public get schema(): z.ZodType | undefined {
    return options;
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