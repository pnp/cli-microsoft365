import { Logger } from '../../../../cli';
import {
  CommandOption
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import Utils from '../../../../Utils';
import { GraphItemsListCommand } from '../../../base/GraphItemsListCommand';
import commands from '../../commands';
import { Tab } from '../../Tab';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  teamId: string;
  channelId: string;
  appId: string;
  appName: string;
  contentUrl: string;
  entityId?: string;
  removeUrl?: string;
  websiteUrl?: string;
}

class TeamsTabAddCommand extends GraphItemsListCommand<Tab> {
  public get name(): string {
    return commands.TAB_ADD;
  }
  public allowUnknownOptions(): boolean | undefined {
    return true;
  }

  public get description(): string {
    return 'Add a tab to the specified channel';
  }
  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.entityId = typeof args.options.entityId !== 'undefined';
    telemetryProps.removeUrl = typeof args.options.removeUrl !== 'undefined';
    telemetryProps.websiteUrl = typeof args.options.websiteUrl !== 'undefined';
    return telemetryProps;
  }
  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {

    const data: any = this.mapRequestBody(args.options);
    const requestOptions: any = {
      url: `${this.resource}/v1.0/teams/${encodeURIComponent(args.options.teamId)}/channels/${args.options.channelId}/tabs`,
      headers: {
        accept: 'application/json;odata.metadata=none',
        'content-type': 'application/json;odata=nometadata'
      },
      data: data,
      responseType: 'json'
    };
    request
      .post(requestOptions)
      .then((res: any): void => {
        logger.log(res);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, logger, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --teamId <teamId>'
      },
      {
        option: '-c, --channelId <channelId>'
      },
      {
        option: '--appId <appId>'
      },
      {
        option: '--appName <appName>'
      },
      {
        option: '--contentUrl <contentUrl>'
      },
      {
        option: '--entityId [entityId]'
      },
      {
        option: '--removeUrl [removeUrl]'
      },
      {
        option: '--websiteUrl [websiteUrl]'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (!Utils.isValidGuid(args.options.teamId as string)) {
      return `${args.options.teamId} is not a valid GUID`;
    }
    if (!Utils.isValidTeamsChannelId(args.options.channelId as string)) {
      return `${args.options.channelId} is not a valid Teams ChannelId`;
    }

    return true;
  }

  private mapRequestBody(options: Options): any {
    const requestBody: any = {};
    requestBody['configuration'] = {};
    const excludeOptions: string[] = [
      'debug',
      'verbose',
      'teamId',
      'channelId',
      'appId',
      'appName',
      'entityId',
      'contentUrl',
      'removeUrl',
      'websiteUrl',
      'output'
    ];
    if (options.appName) {
      requestBody.displayName = options.appName;
    }
    if (options.appId) {
      requestBody['teamsApp@odata.bind'] = `https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/${options.appId}`;
    }
    if (options.contentUrl) {
      requestBody.configuration.contentUrl = options.contentUrl;
    }
    options.entityId ? requestBody.configuration.entityId = options.entityId : null;
    options.removeUrl ? requestBody.configuration.removeUrl = options.removeUrl : null;
    options.websiteUrl ? requestBody.configuration.websiteUrl = options.websiteUrl : null;
    Object.keys(options).forEach(key => {
      if (excludeOptions.indexOf(key) === -1) {
        requestBody.configuration[key] = `${(<any>options)[key]}`;
      }
    });
    return requestBody;
  }
}

module.exports = new TeamsTabAddCommand();