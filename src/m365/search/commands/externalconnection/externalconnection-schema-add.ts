import { AxiosRequestConfig } from 'axios';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  externalConnectionId: string;
  schema: string;
}

interface SearchExternalItem {
  baseType: string;
  properties: Property[];
}

interface Property {
  aliasses?: string[];
  isQueryable?: boolean;
  isRefinable?: boolean;
  isRetrievable?: boolean;
  isSearchable?: boolean;
  labels?: string[];
  name: string;
  type: string;
}

class SearchExternalConnectionSearchAddCommand extends GraphCommand {
  public get name(): string {
    return commands.EXTERNALCONNECTION_SCHEMA_ADD;
  }

  public get description(): string {
    return 'This command will allow the administrator to add a schema to a specific external connection for use in Microsoft Search.';
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

        const schemaObject: SearchExternalItem = JSON.parse(args.options.schema);
        if (schemaObject.baseType === undefined || schemaObject.baseType !== 'microsoft.graph.externalItem') {
          return `The schema needs a required property 'baseType' with value 'microsoft.graph.externalItem'`;
        }
        if (!schemaObject.properties || schemaObject.properties.length > 128) {
          return `We need atleast one property and a maximum of 128 properties in the schema object`;
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      logger.logToStderr(`Adding schema to external connection with id ${args.options.externalConnectionId}`);
    }

    const requestOptions: AxiosRequestConfig = {
      url: `${this.resource}/v1.0/external/connections/${args.options.externalConnectionId}/schema`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json',
      data: args.options.schema
    };

    try {
      await request.post(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new SearchExternalConnectionSearchAddCommand();