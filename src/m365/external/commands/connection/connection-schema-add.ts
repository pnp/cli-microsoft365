import { ExternalConnectors, NullableOption } from '@microsoft/microsoft-graph-types';
import { AxiosResponse } from 'axios';
import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import { Logger } from '../../../../cli/Logger.js';
import request, { CliRequestOptions } from '../../../../request.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  externalConnectionId: z.string()
    .min(3, 'externalConnectionId must be between 3 and 32 characters in length.')
    .max(32, 'externalConnectionId must be between 3 and 32 characters in length.')
    .refine(id => !/[^\w]|_/g.test(id), {
      message: 'externalConnectionId must only contain alphanumeric characters.'
    })
    .refine(id => !(id.length > 9 && id.startsWith('Microsoft')), {
      message: 'ID cannot begin with Microsoft'
    })
    .alias('i'),
  schema: z.string()
    .refine(val => {
      try {
        JSON.parse(val);
        return true;
      }
      catch {
        return false;
      }
    }, {
      message: 'The schema is not a valid JSON string'
    })
    .refine(val => {
      try {
        const obj = JSON.parse(val);
        return obj.baseType === 'microsoft.graph.externalItem';
      }
      catch {
        return true;
      }
    }, {
      message: `The schema needs a required property 'baseType' with value 'microsoft.graph.externalItem'`
    })
    .refine(val => {
      try {
        const obj = JSON.parse(val);
        return obj.properties && obj.properties.length > 0 && obj.properties.length <= 128;
      }
      catch {
        return true;
      }
    }, {
      message: 'We need at least one property and a maximum of 128 properties in the schema object'
    })
    .alias('s'),
  wait: z.boolean().optional()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class ExternalConnectionSchemaAddCommand extends GraphCommand {
  public get name(): string {
    return commands.CONNECTION_SCHEMA_ADD;
  }

  public get description(): string {
    return 'Allows the administrator to add a schema to a specific external connection';
  }

  public alias(): string[] | undefined {
    return [commands.EXTERNALCONNECTION_SCHEMA_ADD];
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Adding schema to external connection with id ${args.options.externalConnectionId}`);
    }

    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/external/connections/${args.options.externalConnectionId}/schema`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json',
      data: args.options.schema,
      fullResponse: true
    };

    try {
      const res = await request.patch<AxiosResponse>(requestOptions);

      const location: string = res.headers.location;
      await logger.log(location);

      if (!args.options.wait) {
        return;
      }

      let status: NullableOption<ExternalConnectors.ConnectionOperationStatus> | undefined;
      do {
        if (this.verbose) {
          await logger.logToStderr(`Waiting 60 seconds...`);
        }

        // waiting 60s before polling as recommended by Microsoft
        await new Promise(resolve => setTimeout(resolve, 60000));

        if (this.debug) {
          await logger.logToStderr(`Checking schema operation status...`);
        }

        const operation = await request.get<ExternalConnectors.ConnectionOperation>({
          url: location,
          headers: {
            accept: 'application/json;odata.metadata=none'
          },
          responseType: 'json'
        });
        status = operation.status;

        if (this.verbose) {
          await logger.logToStderr(`Schema operation status: ${status}`);
        }

        if (status === 'failed') {
          throw `Provisioning schema failed: ${operation.error?.message}`;
        }
      }
      while (status === 'inprogress');
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new ExternalConnectionSchemaAddCommand();