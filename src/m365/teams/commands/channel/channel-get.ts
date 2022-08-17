import { Group } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import { validation } from '../../../../utils';
import { aadGroup } from '../../../../utils/aadGroup';
import GraphCommand from '../../../base/GraphCommand';
import { Channel } from '../../Channel';
import commands from '../../commands';

interface ExtendedGroup extends Group {
  resourceProvisioningOptions: string[];
}

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  teamId?: string;
  teamName?: string;
  channelId?: string;
  channelName?: string;
  primary?: boolean;
}

class TeamsChannelGetCommand extends GraphCommand {
  private teamId: string = "";

  public get name(): string {
    return commands.CHANNEL_GET;
  }

  public get description(): string {
    return 'Gets information about the specific Microsoft Teams team channel';
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
        teamId: typeof args.options.teamId !== 'undefined',
        teamName: typeof args.options.teamName !== 'undefined',
        channelId: typeof args.options.channelId !== 'undefined',
        channelName: typeof args.options.channelName !== 'undefined',
        primary: (!(!args.options.primary)).toString()
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --teamId [teamId]'
      },
      {
        option: '--teamName [teamName]'
      },
      {
        option: '-c, --channelId [channelId]'
      },
      {
        option: '--channelName [channelName]'
      },
      {
        option: '--primary'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.teamId && args.options.teamName) {
          return 'Specify either teamId or teamName, but not both.';
        }

        if (!args.options.teamId && !args.options.teamName) {
          return 'Specify teamId or teamName, one is required';
        }

        if (args.options.teamId && !validation.isValidGuid(args.options.teamId as string)) {
          return `${args.options.teamId} is not a valid GUID`;
        }

        if (args.options.channelId && args.options.channelName && args.options.primary) {
          return 'Specify channelId, channelName or primary';
        }

        if (!args.options.channelId && args.options.channelName && args.options.primary) {
          return 'Specify channelId, channelName or primary.';
        }

        if (args.options.channelId && !args.options.channelName && args.options.primary) {
          return 'Specify channelId, channelName or primary.';
        }

        if (args.options.channelId && args.options.channelName && !args.options.primary) {
          return 'Specify channelId, channelName or primary.';
        }

        if (!args.options.channelId && !args.options.channelName && !args.options.primary) {
          return 'Specify channelId, channelName or primary, one is required';
        }

        if (args.options.channelId && !validation.isValidTeamsChannelId(args.options.channelId as string)) {
          return `${args.options.channelId} is not a valid Teams ChannelId`;
        }

        return true;
      }
    );
  }

  private getTeamId(args: CommandArgs): Promise<string> {
    if (args.options.teamId) {
      return Promise.resolve(args.options.teamId);
    }

    return aadGroup
      .getGroupByDisplayName(args.options.teamName!)
      .then(group => {
        if ((group as ExtendedGroup).resourceProvisioningOptions.indexOf('Team') === -1) {
          return Promise.reject(`The specified team does not exist in the Microsoft Teams`);
        }

        return group.id!;
      });
  }

  private getChannelId(args: CommandArgs): Promise<string> {
    if (args.options.channelId) {
      return Promise.resolve(args.options.channelId);
    }

    if (args.options.primary) {
      return Promise.resolve('');
    }

    const channelRequestOptions: any = {
      url: `${this.resource}/v1.0/teams/${encodeURIComponent(this.teamId)}/channels?$filter=displayName eq '${encodeURIComponent(args.options.channelName as string)}'`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    return request
      .get<{ value: Channel[] }>(channelRequestOptions)
      .then(response => {
        const channelItem: Channel | undefined = response.value[0];

        if (!channelItem) {
          return Promise.reject(`The specified channel does not exist in the Microsoft Teams team`);
        }

        return Promise.resolve(channelItem.id);
      });
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: () => void): void {
    this
      .getTeamId(args)
      .then((teamId: string): Promise<string> => {
        this.teamId = teamId;
        return this.getChannelId(args);
      })
      .then((channelId: string): Promise<Channel> => {
        let url: string = '';
        if (args.options.primary) {
          url = `${this.resource}/v1.0/teams/${encodeURIComponent(this.teamId)}/primaryChannel`;
        }
        else {
          url = `${this.resource}/v1.0/teams/${encodeURIComponent(this.teamId)}/channels/${encodeURIComponent(channelId)}`;
        }

        const requestOptions: any = {
          url: url,
          headers: {
            accept: 'application/json;odata.metadata=none'
          },
          responseType: 'json'
        };

        return request.get<Channel>(requestOptions);
      })
      .then((res: Channel): void => {
        logger.log(res);
        cb();
      }, (err: any) => this.handleRejectedODataJsonPromise(err, logger, cb));
  }
}

module.exports = new TeamsChannelGetCommand();