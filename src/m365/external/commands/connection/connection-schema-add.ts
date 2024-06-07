import { ExternalConnectors, NullableOption } from '@microsoft/microsoft-graph-types';
import { AxiosResponse } from 'axios';
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
  schema: string;
  wait: boolean;
}

interface ExternalItem {
  baseType: string;
  properties: Property[];
}

interface Property {
  aliases?: string[];
  isQueryable?: boolean;
  isRefinable?: boolean;
  isRetrievable?: boolean;
  isSearchable?: boolean;
  labels?: string[];
  name: string;
  type: string;
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

  constructor() {
    super();

    this.#initOptions();
    this.#initValidators();
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --externalConnectionId <externalConnectionId>'
      },
      {
        option: '-s, --schema <schema>'
      },
      {
        option: '--wait'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.externalConnectionId.length < 3 || args.options.externalConnectionId.length > 32) {
          return 'externalConnectionId must be between 3 and 32 characters in length.';
        }

        const alphaNumericRegEx = /[^\w]|_/g;

        if (alphaNumericRegEx.test(args.options.externalConnectionId)) {
          return 'externalConnectionId must only contain alphanumeric characters.';
        }

        if (args.options.externalConnectionId.length > 9 &&
          args.options.externalConnectionId.startsWith('Microsoft')) {
          return 'ID cannot begin with Microsoft';
        }

        const schemaObject: ExternalItem = JSON.parse(args.options.schema);
        if (schemaObject.baseType === undefined || schemaObject.baseType !== 'microsoft.graph.externalItem') {
          return `The schema needs a required property 'baseType' with value 'microsoft.graph.externalItem'`;
        }

        if (!schemaObject.properties || schemaObject.properties.length > 128) {
          return `We need at least one property and a maximum of 128 properties in the schema object`;
        }

        return true;
      }
    );
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