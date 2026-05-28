import { ExternalConnectors } from '@microsoft/microsoft-graph-types/microsoft-graph';
import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { Logger } from '../../../../cli/Logger.js';
import request from '../../../../request.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

const invalidIds: string[] = ['None',
  'Directory',
  'Exchange',
  'ExchangeArchive',
  'LinkedIn',
  'Mailbox',
  'OneDriveBusiness',
  'SharePoint',
  'Teams',
  'Yammer',
  'Connectors',
  'TaskFabric',
  'PowerBI',
  'Assistant',
  'TopicEngine',
  'MSFT_All_Connectors'
];

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  id: z.string()
    .min(3, 'ID must be between 3 and 32 characters in length.')
    .max(32, 'ID must be between 3 and 32 characters in length.')
    .refine(id => !/[^\w]|_/g.test(id), {
      message: 'ID must only contain alphanumeric characters.'
    })
    .refine(id => !(id.length > 9 && id.startsWith('Microsoft')), {
      message: 'ID cannot begin with Microsoft'
    })
    .refine(id => !invalidIds.includes(id), {
      error: () => `ID cannot be one of the following values: ${invalidIds.join(', ')}.`
    })
    .alias('i'),
  name: z.string().alias('n'),
  description: z.string().alias('d'),
  authorizedAppIds: z.string().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class ExternalConnectionAddCommand extends GraphCommand {
  public get name(): string {
    return commands.CONNECTION_ADD;
  }

  public get description(): string {
    return 'Adds a new external connection';
  }

  public alias(): string[] | undefined {
    return [commands.EXTERNALCONNECTION_ADD];
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    let appIds: string[] = [];

    if (args.options.authorizedAppIds !== undefined &&
      args.options.authorizedAppIds !== '') {
      appIds = args.options.authorizedAppIds.split(',');
    }

    const commandData: ExternalConnectors.ExternalConnection = {
      id: args.options.id,
      name: args.options.name,
      description: args.options.description,
      configuration: {
        authorizedAppIds: appIds
      }
    };

    const requestOptions: any = {
      url: `${this.resource}/v1.0/external/connections`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json',
      data: commandData
    };

    try {
      await request.post(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new ExternalConnectionAddCommand();
