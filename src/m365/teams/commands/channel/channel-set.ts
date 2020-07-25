import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import GraphCommand from '../../../base/GraphCommand';
import Utils from '../../../../Utils';
import request from '../../../../request';
import { Channel } from '../../Channel';
import * as chalk from 'chalk';
import { CommandInstance } from '../../../../cli';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  channelName: string;
  description?: string
  newChannelName?: string;
  teamId: string;
}

class TeamsChannelSetCommand extends GraphCommand {
  public get name(): string {
    return `${commands.TEAMS_CHANNEL_SET}`;
  }
  public get description(): string {
    return 'Updates properties of the specified channel in the given Microsoft Teams team';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.newChannelName = typeof args.options.newChannelName !== 'undefined';
    telemetryProps.description = typeof args.options.description !== 'undefined';
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const requestOptions: any = {
      url: `${this.resource}/v1.0/teams/${encodeURIComponent(args.options.teamId)}/channels?$filter=displayName eq '${encodeURIComponent(args.options.channelName)}'`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      json: true
    }

    request
      .get<{ value: Channel[] }>(requestOptions)
      .then((res: { value: Channel[] }): Promise<void> => {
        const channelItem: Channel | undefined = res.value[0];

        if (!channelItem) {
          return Promise.reject(`The specified channel does not exist in the Microsoft Teams team`);
        }

        const channelId: string = res.value[0].id;
        const body: any = this.mapRequestBody(args.options);
        const requestOptions: any = {
          url: `${this.resource}/v1.0/teams/${encodeURIComponent(args.options.teamId)}/channels/${channelId}`,
          headers: {
            'accept': 'application/json;odata.metadata=none'
          },
          json: true,
          body: body
        };

        return request.patch(requestOptions);
      })
      .then((): void => {
        if (this.verbose) {
          cmd.log(chalk.green('DONE'));
        }

        cb();
      }, (err: any) => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --teamId <teamId>',
        description: 'The ID of the team where the channel to update is located'
      },
      {
        option: '--channelName <channelName>',
        description: 'The name of the channel to update'
      },
      {
        option: '--newChannelName [newChannelName]',
        description: 'The new name of the channel'
      },
      {
        option: '--description [description]',
        description: 'The description of the channel'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!Utils.isValidGuid(args.options.teamId)) {
        return `${args.options.teamId} is not a valid GUID`;
      }

      if (args.options.channelName.toLowerCase() === "general") {
        return 'General channel cannot be updated';
      }

      return true;
    };
  }

  private mapRequestBody(options: Options): any {
    const requestBody: any = {};

    if (options.newChannelName) {
      requestBody.displayName = options.newChannelName;
    }

    if (options.description) {
      requestBody.description = options.description;
    }

    return requestBody;
  }
}

module.exports = new TeamsChannelSetCommand();