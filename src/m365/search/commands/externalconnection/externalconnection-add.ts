//import { parseConfigFileTextToJson } from 'typescript';
import { Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
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
  description?: string;
  authorizedAppIds?: string;
}

class SearchExternalConnectionAddCommand extends GraphCommand {
  public get name(): string {
    return commands.EXTERNALCONNECTION_ADD;
  }

  public get description(): string {
    return 'Adds a new External Connection for Microsoft Search';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    return telemetryProps;
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'name', 'description'];
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    if (this.verbose) {
      logger.logToStderr(`Adding new external connections...`);
    }

    let appIds:string[] = [];

    if (args.options.authorizedAppIds !== undefined && args.options.authorizedAppIds !== '') {
      appIds = args.options.authorizedAppIds?.split(',');
    }

    const commandData = {
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
        'accept': 'application/json;odata.metadata=none'
      },
      responseType: 'json',
      data: commandData
    };
    if (this.verbose) { logger.logToStderr(`Adding new external connections...`); }
    request
      .post(requestOptions)
      .then(_ => cb(), (err: any) => {
        this.handleRejectedODataJsonPromise(err, logger, cb);
      });
  }


  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: "--id <id>"
      },
      {
        option: "--name <name>"
      },
      {
        option: "--description <description>"
      },
      {
        option: "--authorizedAppIds [authorizedAppIds]"
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (args.options.id.length < 3 || args.options.id.length > 32) {
      return 'ID field must be between 3 and 32 characters in length.';
    }

    //var alphanumeric = "someStringHere";
    const alphaNumericRegEx  = /[^\w]|_/g;

    if (alphaNumericRegEx.test(args.options.id)) {
      return 'ID field must only contain alphanumeric characters.';
    }

    if (args.options.id.length > 9 && args.options.id.startsWith('Microsoft')) {
      return 'ID field cannot begin with Microsoft';
    }

    if (args.options.id === 'None'
        || args.options.id === 'Directory'
        || args.options.id === 'Exchange'
        || args.options.id === 'ExchangeArchive'
        || args.options.id === 'LinkedIn'
        || args.options.id === 'Mailbox'
        || args.options.id === 'OneDriveBusiness'
        || args.options.id === 'SharePoint'
        || args.options.id === 'Teams'
        || args.options.id === 'Yammer'
        || args.options.id === 'Connectors'
        || args.options.id === 'TaskFabric'
        || args.options.id === 'PowerBI'
        || args.options.id === 'Assistant'
        || args.options.id === 'TopicEngine'
        || args.options.id === 'MSFT_All_Connectors'
    ) {
      return 'ID field cannot be one of the following values: None, Directory, Exchange, ExchangeArchive, LinkedIn, Mailbox, OneDriveBusiness, SharePoint, Teams, Yammer, Connectors, TaskFabric, PowerBI, Assistant, TopicEngine, MSFT_All_Connectors.';
    }

    return true;
  }
}

module.exports = new SearchExternalConnectionAddCommand();