import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import Utils from '../../../../Utils';
import request from '../../../../request';
import GraphCommand from '../../../base/GraphCommand';
import { Channel } from '../../Channel';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  channelId?: string;
  channelName?: string;
  teamId: string;
  confirm?: boolean;
}

class TeamsChannelRemoveCommand extends GraphCommand {
  public get name(): string {
    return commands.TEAMS_CHANNEL_REMOVE;
  }

  public get description(): string {
    return 'Removes the specified channel in the Microsoft Teams team';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.channelId = typeof args.options.channelId !== 'undefined';
    telemetryProps.channelName = typeof args.options.channelName !== 'undefined';
    telemetryProps.confirm = (!(!args.options.confirm)).toString();
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {

    const removeChannel: () => void = (): void => {
      if (args.options.channelName) {
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

            const requestOptions: any = {
              url: `${this.resource}/v1.0/teams/${encodeURIComponent(args.options.teamId)}/channels/${encodeURIComponent(channelId)}`,
              headers: {
                accept: 'application/json;odata.metadata=none'
              },
              json: true
            };

            return request.delete(requestOptions);
          })
          .then((): void => {
            if (this.verbose) {
              cmd.log(vorpal.chalk.green('DONE'));
            }

            cb();
          }, (err: any) => this.handleRejectedODataJsonPromise(err, cmd, cb));
      }

      if (args.options.channelId) {
        const requestOptions: any = {
          url: `${this.resource}/v1.0/teams/${encodeURIComponent(args.options.teamId)}/channels/${encodeURIComponent(args.options.channelId)}`,
          headers: {
            accept: 'application/json;odata.metadata=none'
          },
          json: true
        };

        request
          .delete(requestOptions)
          .then((): void => {
            if (this.verbose) {
              cmd.log(vorpal.chalk.green('DONE'));
            }

            cb();
          }, (err: any) => this.handleRejectedODataJsonPromise(err, cmd, cb));
      }
    };

    if (args.options.confirm) {
      removeChannel();
    }
    else {
      const channelName = args.options.channelName ? args.options.channelName : args.options.channelId;
      cmd.prompt({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove the channel ${channelName} from team ${args.options.teamId}?`
      }, (result: { continue: boolean }): void => {
        if (!result.continue) {
          cb();
        }
        else {
          removeChannel();
        }
      });
    }
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-c, --channelId [channelId]',
        description: 'The ID of the channel to remove'
      },
      {
        option: '-n, --channelName [channelName]',
        description: 'The name of the channel to remove. Specify channelId or channelName but not both'
      },
      {
        option: '-i, --teamId <teamId>',
        description: 'The ID of the team to which the channel to remove belongs'
      },
      {
        option: '--confirm',
        description: 'Don\'t prompt for confirmation'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (args.options.channelId && args.options.channelName) {
        return 'Specify channelId or channelName but not both';
      }

      if (!args.options.channelId && !args.options.channelName) {
        return 'Specify channelId or channelName';
      }

      if (args.options.channelId && !Utils.isValidTeamsChannelId(args.options.channelId)) {
        return `${args.options.channelId} is not a valid Teams Channel Id`;
      }

      if (!args.options.teamId) {
        return 'Required parameter teamId missing';
      }

      if (args.options.teamId && !Utils.isValidGuid(args.options.teamId)) {
        return `${args.options.teamId} is not a valid GUID`;
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Remarks:
      
    When deleted, Microsoft Teams channels are moved to a recycle bin and can be
    restored within 30 days. After that time, they are permanently deleted.
      
  Examples:
    
    Remove the specified Microsoft Teams channel by Id
      ${this.name} --channelId 19:f3dcbb1674574677abcae89cb626f1e6@thread.skype --teamId d66b8110-fcad-49e8-8159-0d488ddb7656

    Remove the specified Microsoft Teams channel by Id without confirmation
      ${this.name} --channelId 19:f3dcbb1674574677abcae89cb626f1e6@thread.skype --teamId d66b8110-fcad-49e8-8159-0d488ddb7656 --confirm

    Remove the specified Microsoft Teams channel by Name
      ${this.name} --channelName 'channelName' --teamId d66b8110-fcad-49e8-8159-0d488ddb7656

    Remove the specified Microsoft Teams channel by Name without confirmation
      ${this.name} --channelName 'channelName' --teamId d66b8110-fcad-49e8-8159-0d488ddb7656 --confirm  

  More information:

    directory resource type (deleted items)
      https://docs.microsoft.com/en-us/graph/api/resources/directory?view=graph-rest-1.0
`);
  }
}

module.exports = new TeamsChannelRemoveCommand();