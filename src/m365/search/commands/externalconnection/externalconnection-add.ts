import { ExternalConnectors } from '@microsoft/microsoft-graph-types/microsoft-graph';
import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id: string;
  name: string;
  description: string;
  authorizedAppIds?: string;
}

class SearchExternalConnectionAddCommand extends GraphCommand {
  public get name(): string {
    return commands.EXTERNALCONNECTION_ADD;
  }

  public get description(): string {
    return 'Adds a new External Connection for Microsoft Search';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        authorizedAppIds: typeof args.options.authorizedAppIds !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --id <id>'
      },
      {
        option: '-n, --name <name>'
      },
      {
        option: '-d, --description <description>'
      },
      {
        option: '--authorizedAppIds [authorizedAppIds]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        const id = args.options.id;
        if (id.length < 3 || id.length > 32) {
          return 'ID must be between 3 and 32 characters in length.';
        }

        const alphaNumericRegEx = /[^\w]|_/g;

        if (alphaNumericRegEx.test(id)) {
          return 'ID must only contain alphanumeric characters.';
        }

        if (id.length > 9 &&
          id.startsWith('Microsoft')) {
          return 'ID cannot begin with Microsoft';
        }

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

        if (invalidIds.indexOf(id) > -1) {
          return `ID cannot be one of the following values: ${invalidIds.join(', ')}.`;
        }

        return true;
      }
    );
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
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

    request
      .post(requestOptions)
      .then(_ => cb(), err => this.handleRejectedODataJsonPromise(err, logger, cb));
  }
}

module.exports = new SearchExternalConnectionAddCommand();
