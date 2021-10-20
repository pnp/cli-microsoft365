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
    telemetryProps.id = typeof args.options.id !== 'undefined';
    telemetryProps.name = typeof args.options.name !== 'undefined';
    telemetryProps.description = typeof args.options.description !== 'undefined';
    return telemetryProps;
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'name', 'description'];
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    if (this.verbose) {
      logger.logToStderr(`Adding new external connections...`);
    }

    const appIds:string[] = [];

    if (args.options.authorizedAppIds !== undefined && args.options.authorizedAppIds !== '') {
      const splitAppIds = args.options.authorizedAppIds?.split(',');
      
      splitAppIds.forEach(appId => {
        appIds.push(appId);
      });
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
        option: "--id [id]"
      },
      {
        option: "--name [name]"
      },
      {
        option: "--description [description]"
      },
      {
        option: "--authorizedAppIds [authorizedAppIds]"
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (!args.options.id || !args.options.name || !args.options.description) {
      return 'Specify id, name and description';
    }


    return true;
  }
}

module.exports = new SearchExternalConnectionAddCommand();