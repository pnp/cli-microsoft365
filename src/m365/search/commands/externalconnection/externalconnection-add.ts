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
  name: string;
  connectionId?: string;
  connectionDescription?: string;
  authorisedAppIds?: string[];
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
    telemetryProps.externalConnectionId = typeof args.options.externalConnectionId !== 'undefined';
    telemetryProps.externalConnectionName = typeof args.options.externalConnectionName !== 'undefined';
    telemetryProps.externalConnectionDescription = typeof args.options.externalConnectionDescription !== 'undefined';
    return telemetryProps;
  }

  public defaultProperties(): string[] | undefined {
    return ['externalConnectionId', 'externalConnectionName', 'externalConnectionDescription'];
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    if (this.verbose) {
      logger.logToStderr(`Adding new external connections...`);
    }
    const requestOptions: any = {
      url: `${this.resource}/v1.0/external/connections`,
      headers: {
        'accept': 'application/json;odata.metadata=none'
      },
      responseType: 'json',
      data: {
        id: args.options.externalConnectionId,
        name: args.options.externalConnectionName,
        description: args.options.externalConnectionDescription
      }
    };
    logger.logToStderr(`Adding new external connections...`);
    request
      .post(requestOptions)
      .then(_ => cb(), (err: any) => {
        logger.logToStderr(`Errored adding new external connections...`);
        logger.logToStderr(err);
        this.handleRejectedODataJsonPromise(err, logger, cb);
      });
  }


  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: "--externalConnectionId [externalConnectionId]"
      },
      {
        option: "--externalConnectionName [externalConnectionName]"
      },
      {
        option: "--externalConnectionDescription [externalConnectionDescription]"
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (!args.options.externalConnectionId || !args.options.externalConnectionName || !args.options.externalConnectionDescription) {
      return 'Specify externalConnectionId, externalConnectionName and externalConnectionDescription';
    }


    return true;
  }
}

module.exports = new SearchExternalConnectionAddCommand();